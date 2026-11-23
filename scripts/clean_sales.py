#!/usr/bin/env python3
import argparse
import hashlib
import json
import re
import subprocess
import time
import unicodedata
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

ORDER_LINE = '销售订单行号'
CUSTOMER = '客户编号'
MATERIAL = '物料编码'
CUSTOMER_DATE = '客户建档日期'
QUANTITY = '数量'
NUMBERS = set('序号|数量|集团货币金额|集团货币含税金额|公司货币金额|公司货币含税金额|不含税金额（交易货币）|含税金额（交易货币）'.split('|'))
DATES = set('最早审核时间|最后审核时间|最后修改时间|要求交货日期|客户建档日期|创建时间'.split('|'))
ORDER = '销售订单编号'
FY = '客户建档财年'
SOURCE = '客户来源'
KEY = '导入记录键'
QUALITY = '数据质量标记'
TRADE = '含税金额（交易货币）'
GROUP = '集团货币含税金额'
CHANNEL = ('渠道大类', '渠道子来源')
DETAIL_CHANNEL = ('客户来源（映射）', '客户子来源')
DETAIL_CHANNEL_ALIASES = ('客户渠道', '客户子渠道')
MATERIAL_SCHEME = '解决方案'
MATERIAL_CODE = '物料编码'
MATERIAL_MAP_KEY = '完整物料编码'
MATERIAL_MAP_VALUE = '归属方案'
MATERIAL_MAP_NAMES = ('解决方案映射表', '物料方案映射表')
SELECT_FIELDS = {FY, *CHANNEL}
MAP_TABLE = '渠道映射表'
MAP_SOURCE = 'ERP客户来源'
MAP_FIELDS = (MAP_SOURCE, *CHANNEL)
DETAIL_VISIBLE = (ORDER, '销售订单行号', '客户简称', '城市', *DETAIL_CHANNEL, '业务员', '销售组', '物料编码', '物料名称', '物料大类', '物料分类', MATERIAL_SCHEME, '数量', '单位', TRADE, '交易货币', GROUP, '集团货币', '客户建档日期')


def encode(value):
    return json.dumps(value, ensure_ascii=False, separators=(',', ':'), allow_nan=False)


def normalize(value):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(value or ''))).casefold()


def channel_key(value):
    return str(value or '').strip().replace(' ', '').replace('（', '(').replace('）', ')').lower()


def channel_index(rows):
    counts = Counter(channel_key(row[MAP_SOURCE]) for row in rows)
    return {channel_key(row[MAP_SOURCE]): (row[CHANNEL[0]] or '未分类', row[CHANNEL[1]])
            for row in rows if counts[channel_key(row[MAP_SOURCE])] == 1}


def channel_formula(name, source, index):
    def key(ref):
        return f'LOWER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM({ref})," ",""),"（","("),"）",")"))'
    match = f'[{MAP_TABLE}].FILTER({key(f"CurrentValue.[{MAP_SOURCE}]")}={key(f"[{SOURCE}]")})'
    result = f'{match}.[{source}].ARRAYJOIN("")'
    fallback = '"未分类"' if index == 0 else '""'
    expr = (f'IF(ISBLANK([{SOURCE}]),{fallback},IF({match}.[{MAP_SOURCE}].COUNTA()=1,'
            f'IF(ISBLANK({result}),{fallback},{result}),{fallback}))')
    return {'name': name, 'type': 'formula', 'expression': expr}


def ensure_channel_formulas(base, table, source_fields=()):
    args = ('--base-token', base, '--table-id', table)
    fields = {f['name']: f for f in cli('field-list', *args)['fields']}
    names = []
    for i, (default, alias) in enumerate(zip(DETAIL_CHANNEL, DETAIL_CHANNEL_ALIASES)):
        existing_alias = alias in fields and alias not in source_fields
        if default in fields and existing_alias:
            raise ValueError(f'{table} 同时存在「{default}」与「{alias}」，请先确认渠道映射列')
        name = alias if existing_alias else default
        names.append(name)
        definition = channel_formula(name, CHANNEL[i], i)
        field = fields.get(name)
        if not field:
            cli('field-create', *args, '--json', encode(definition), '--i-have-read-guide')
        elif field['type'] != 'formula':
            raise ValueError(f'{table} 的「{name}」须先迁移成联动公式')
        actual = cli('field-get', *args, '--field-id', name)['field']
        if actual.get('expression') != definition['expression']:
            raise ValueError(f'{table} 的「{name}」映射公式与规则约定不一致')
    return tuple(names)


def material_scheme_formula(table):
    match = (f'[{table}].FILTER(UPPER(TRIM(CurrentValue.[{MATERIAL_MAP_KEY}]))'
             f'=UPPER(TRIM([{MATERIAL_CODE}])))')
    result = f'{match}.[{MATERIAL_MAP_VALUE}].ARRAYJOIN("")'
    expression = (f'IF(ISBLANK([{MATERIAL_CODE}]),"待确认",'
                  f'IF({match}.[{MATERIAL_MAP_KEY}].COUNTA()=1,'
                  f'IF(ISBLANK({result}),"待确认",{result}),"待确认"))')
    return {'name': MATERIAL_SCHEME, 'type': 'formula', 'expression': expression}


def material_scheme_counts(base, table, detail):
    counts, schemes = Counter(), {}
    args = ('--base-token', base, '--table-id', table['id'])
    offset = 0
    while True:
        page = cli('record-list', *args, '--format', 'json', '--limit', '200',
                   '--offset', str(offset), '--field-id', MATERIAL_MAP_KEY,
                   '--field-id', MATERIAL_MAP_VALUE)
        if page['fields'] != [MATERIAL_MAP_KEY, MATERIAL_MAP_VALUE]:
            raise ValueError('物料方案映射字段回读不一致')
        for code, scheme in page['data']:
            key = str(code or '').strip().upper()
            counts[key] += 1
            schemes[key] = select_value(scheme) or '待确认'
        if not page['has_more']:
            break
        if not page['data']:
            raise ValueError('物料方案映射表分页异常')
        offset += len(page['data'])

    def resolve(row):
        key = str(row.get(MATERIAL_CODE) or '').strip().upper()
        return schemes[key] if key and counts[key] == 1 else '待确认'

    return {fy: Counter(resolve(row) for row in rows) for fy, rows in detail.items()}


def ensure_material_scheme_formula(base, table, mapping_table):
    args = ('--base-token', base, '--table-id', table)
    definition = material_scheme_formula(mapping_table)
    fields = {f['name']: f for f in cli('field-list', *args)['fields']}
    if MATERIAL_SCHEME not in fields:
        cli('field-create', *args, '--json', encode(definition), '--i-have-read-guide')
    actual = cli('field-get', *args, '--field-id', MATERIAL_SCHEME)['field']
    if actual['type'] != 'formula' or actual.get('expression') != definition['expression']:
        raise ValueError(f'{table} 的「解决方案」字段与物料映射规则不一致')


def clean_cell(name, value):
    if value is None:
        return None
    if name in DATES:
        if not str(value).strip():
            return None
        return datetime.fromisoformat(str(value).strip()).strftime('%Y-%m-%d %H:%M:%S')
    if name in NUMBERS and isinstance(value, (int, float)):
        return value
    value = str(value).strip()
    return (float(value) if name in NUMBERS else value) if value else None


def read_source(path, company_column, reversal_column):
    required = {ORDER, ORDER_LINE, CUSTOMER_DATE, SOURCE, company_column}
    if reversal_column:
        required.update({reversal_column, CUSTOMER, MATERIAL, '单位', '交易货币',
                         '集团货币', TRADE, GROUP, QUANTITY})
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        candidates = []
        for sheet in book:
            sheet.reset_dimensions()
            rows = sheet.values
            header = tuple(str(x or '').strip() for x in next(rows, ()))
            if required.issubset(header):
                candidates.append((header, rows))
        if len(candidates) != 1:
            raise ValueError('需要一张包含必填字段的订单明细工作表')
        header, rows = candidates[0]
        if not all(header) or len(set(header)) != len(header):
            raise ValueError('源字段名不能为空或重复')
        reserved = {FY, KEY, QUALITY, *CHANNEL, *DETAIL_CHANNEL}
        if reserved.intersection(header):
            raise ValueError('源字段名包含脚本保留字段，请先在副本中重命名')
        records = []
        for n, values in enumerate(rows, 2):
            if not any(v is not None for v in values):
                continue
            try:
                row = {name: clean_cell(name, values[i] if i < len(values) else None)
                       for i, name in enumerate(header)}
                if not row[ORDER] or not row[CUSTOMER_DATE]:
                    raise ValueError('销售订单编号或客户建档日期为空')
                records.append(row)
            except ValueError as exc:
                raise ValueError(f'源表第 {n} 行：{exc}') from exc
        return header, records
    finally:
        book.close()


def details(rows, columns, index, company_column, fiscal_start_month):
    def order_line(row):
        return row[company_column], row[ORDER], row[ORDER_LINE]

    key_columns = tuple(name for name in columns if name != '序号')
    by_key = defaultdict(set)
    master = defaultdict(set)
    for row in rows:
        key = order_line(row)
        if all(key):
            by_key[key].add(tuple(row[name] for name in key_columns))
        for field, anchor in (('物料分类', '物料编码'), ('物料大类', '物料编码'), (SOURCE, '客户编号')):
            if row.get(anchor) and normalize(row.get(field)) not in ('', '0', '#n/a'):
                master[field, row.get(anchor)].add(row.get(field))
    result = defaultdict(list)
    seen = set()
    counts = Counter()
    unmapped = Counter()
    for row in rows:
        key = order_line(row)
        payload = encode([row[name] for name in key_columns])
        if all(key) and payload in seen:
            counts['完全重复移除'] += 1
            continue
        if all(key):
            seen.add(payload)
        digest = hashlib.sha256(payload.encode()).hexdigest()
        day = date.fromisoformat(row['客户建档日期'][:10])
        year = day.year - (day.month < fiscal_start_month)
        fiscal = f'FY{year % 100:02}-{(year + 1) % 100:02}'
        record = {**row, FY: fiscal, KEY: digest if all(key) else f"{digest}-{row.get('序号')}"}
        tags = []
        for field, anchor in (('物料分类', '物料编码'), ('物料大类', '物料编码'), (SOURCE, '客户编号')):
            value = row.get(field)
            if normalize(value) in ('', '0', '#n/a'):
                choices = master[field, row.get(anchor)]
                if len(choices) == 1:
                    value = next(iter(choices))
                    counts[f'{field}补齐'] += 1
                    tags.append(f'{field}按同编码补齐')
                else:
                    tags.append(f'{field}待核查' if choices else f'{field}缺失')
        raw_source = channel_key(row[SOURCE])
        if raw_source and raw_source in index:
            record.update(zip(CHANNEL, index[raw_source]))
        else:
            record.update(zip(CHANNEL, ('未分类', None)))
            unmapped[row[SOURCE] or '<空>'] += 1
        record.update(zip(DETAIL_CHANNEL, (record[name] for name in CHANNEL)))
        if row.get(TRADE) is not None and row[TRADE] <= 0:
            tags.append('交易含税金额为零' if row[TRADE] == 0 else '交易含税金额为负')
        if '交易货币' in row and not row['交易货币']:
            tags.append('交易币种缺失')
        if not all(key):
            tags.append('订单行键不完整')
        if len(by_key.get(key, ())) > 1:
            tags.append('同订单行字段冲突')
        record[QUALITY] = '；'.join(tags) or None
        result[fiscal].append(record)
    keys = [r[KEY] for rows in result.values() for r in rows]
    if len(set(keys)) != len(keys):
        raise ValueError('明细匹配标识重复，无法安全定位记录')
    return dict(sorted(result.items())), counts, unmapped


def cancel_matched_offsets(detail, company_column, reversal_column, reversal_value):
    removed = {}
    audit = Counter()
    for fiscal, rows in detail.items():
        groups = defaultdict(list)
        for row in rows:
            groups[(row[company_column], row[ORDER])].append(row)
        matched = set()
        for order in groups.values():
            offsets = [r for r in order if r[reversal_column] == reversal_value]
            if not offsets:
                continue
            normal = [r for r in order if not r[reversal_column]]
            if len(order) == len(normal) + len(offsets) and offsets_match(normal, offsets, company_column):
                matched.update(r[KEY] for r in order)
                audit['已抵消订单'] += 1
                audit['移除正常明细'] += len(normal)
                audit['移除补单明细'] += len(offsets)
            else:
                audit['待核查订单'] += 1
        removed[fiscal] = [r for r in rows if r[KEY] in matched]
        detail[fiscal] = [r for r in rows if r[KEY] not in matched]
    return removed, audit


def offsets_match(normal, offsets, company_column):
    dimensions = (company_column, '客户编号', '物料编码', '单位', '交易货币', '集团货币')
    amounts = (TRADE, GROUP, '数量')
    if not normal or len(normal) != len(offsets):
        return False
    if any(r[name] is None for r in (*normal, *offsets) for name in (*dimensions, *amounts)):
        return False
    if any(r[GROUP] < 0 for r in normal) or any(r[GROUP] > 0 for r in offsets):
        return False

    def signature(row, sign):
        return (*[row[name] for name in dimensions], Decimal(str(row[TRADE])),
                sign * Decimal(str(row[GROUP])), sign * Decimal(str(row['数量'])))

    return Counter(signature(r, 1) for r in normal) == Counter(signature(r, -1) for r in offsets)



def cli(command, *args):
    proc = subprocess.run(['lark-cli', 'base', f'+{command}', '--as', 'user', *args],
                          capture_output=True, text=True, encoding='utf-8', timeout=120)
    result = json.loads(proc.stdout) if proc.stdout else {}
    if proc.returncode or not result.get('ok'):
        raise RuntimeError(f'{command} 执行失败：{proc.stderr[-900:] or str(result.get("error", result))[:900]}')
    return result['data']


def schema(names, options):
    result = []
    for name in names:
        field = {'name': name, 'type': 'datetime' if name in DATES else
                 'number' if name in NUMBERS else 'text'}
        if name == '数量':
            field['style'] = {'type': 'plain', 'precision': 0}
        if name in SELECT_FIELDS:
            field.update(type='select', multiple=False,
                         options=[{'name': value} for value in sorted(options[name])])
        result.append(field)
    return result


def select_value(value):
    if isinstance(value, list):
        if len(value) > 1:
            raise ValueError('单选字段含有多个值')
        return value[0] if value else None
    return value


def writable_cells(values):
    return {name: [value] if name in SELECT_FIELDS and value is not None else value
            for name, value in values.items()}


def read_remote(args, names, expected_count):
    records = {}
    selectors = [arg for name in names for arg in ('--field-id', name)]
    offset = 0
    while True:
        page = cli('record-list', *args, '--format', 'json', '--limit', '200',
                   '--offset', str(offset), *selectors)
        assert page['fields'] == list(names), page['fields']
        for record_id, row in zip(page['record_id_list'], page['data'], strict=True):
            if row[0] is None or row[0] in records:
                raise ValueError(f'飞书表存在空白或重复的行标识：{args[-1]}')
            cells = dict(zip(names, row, strict=True))
            for name in DATES.intersection(cells):
                if isinstance(cells[name], str):
                    cells[name] = cells[name][:19].replace('T', ' ')
            for name in SELECT_FIELDS.intersection(cells):
                cells[name] = select_value(cells[name])
            records[row[0]] = (record_id, cells)
        offset += len(page['data'])
        if not page['has_more']:
            return records
        if offset > expected_count or not page['data']:
            raise ValueError(f'分页或表内记录数异常：{args[-1]}')


def batches(items):
    batch = []
    size = 0
    for item in items:
        weight = len(encode(item).encode('utf-8')) + 4
        if weight > 110_000:
            raise ValueError('单条记录过大，无法通过飞书 CLI 导入')
        if batch and (len(batch) == 180 or size + weight > 110_000):
            yield batch
            batch = []
            size = 0
        batch.append(item)
        size += weight
    if batch:
        yield batch


def ensure_schema(base, table, names, options):
    args = ('--base-token', base, '--table-id', table)
    fields = schema(names, options)
    tables = cli('table-list', '--base-token', base)['tables']
    if not any(item['name'] == table for item in tables):
        cli('table-create', '--base-token', base, '--name', table, '--fields', encode(fields))
    current = {f['name']: f['type'] for f in cli('field-list', *args)['fields']}
    missing = [f for f in fields if f['name'] not in current]
    if any(current.get(f['name'], f['type']) != f['type'] for f in fields):
        raise ValueError(f'字段类型冲突：{table}')
    for start in range(0, len(missing), 4):
        cli('field-create', *args, '--json', encode(missing[start:start+4]))
        time.sleep(1.1)
    if missing:
        present = {f['name'] for f in cli('field-list', *args)['fields']}
        if not set(names).issubset(present):
            raise ValueError(f'字段回读不完整：{table}')


def sync_table(base, table, records, names, writable, visible, identity=KEY):
    args = ('--base-token', base, '--table-id', table)
    projected = (identity, *(n for n in writable if n != identity))
    expected = {r[identity]: r for r in records}
    remote = read_remote(args, projected, len(expected))
    extra = remote.keys() - expected.keys()
    if extra:
        raise ValueError(f'{table} 有 {len(extra)} 条不在本次导出的记录，可能为不同快照，请用新 Base')
    to_create, to_update = [], {}
    for key, row in expected.items():
        if key not in remote:
            to_create.append(writable_cells({name: row[name] for name in names if name not in DETAIL_CHANNEL and row.get(name) is not None}))
        else:
            rid, actual = remote[key]
            delta = {n: row.get(n) for n in projected[1:] if actual[n] != row.get(n)}
            if delta:
                to_update[rid] = writable_cells(delta)
    for batch in batches(to_create):
        result = cli('record-batch-create', *args, '--json', encode({'create_records': batch}))
        if len(result['record_id_list']) != len(batch):
            raise ValueError(f'飞书返回写入数异常：{table}')
    for batch in batches(to_update.items()):
        cli('record-batch-update', *args, '--json', encode({'update_records': dict(batch)}))
    readback = read_remote(args, projected, len(expected))
    if readback.keys() != expected.keys():
        raise ValueError(f'飞书记录数或记录标识核对失败：{table}')
    for key, (_, values) in readback.items():
        if any(value != expected[key].get(name) for name, value in values.items()):
            raise ValueError(f'飞书字段核对失败：{table}')
    if visible:
        set_views(args, names, visible)
    print(encode({'表': table, '回读': len(readback), '新增': len(to_create), '更新': len(to_update)}), flush=True)


def set_views(args, names, visible):
    views = {v['name']: v for v in cli('view-list', *args)['views']}
    if '表格' in views and '订单明细' not in views:
        default = views.pop('表格')
        cli('view-rename', *args, '--view-id', default['id'], '--name', '订单明细')
        views['订单明细'] = default
    for name, columns in (
        ('订单明细', visible), ('全部字段', tuple(n for n in names if n != KEY))
    ):
        if name not in views:
            cli('view-create', *args, '--json', encode({'name': name, 'type': 'grid'}))
        current = cli('view-get-visible-fields', *args, '--view-id', name)['visible_fields']
        if tuple(n for n in current if n != KEY) != columns:
            cli('view-set-visible-fields', *args, '--view-id', name,
                '--json', encode({'visible_fields': columns}))
            current = cli('view-get-visible-fields', *args, '--view-id', name)['visible_fields']
            if tuple(n for n in current if n != KEY) != columns:
                raise ValueError(f'视图可见字段回读不一致：{args[-1]} / {name}')
    view = '订单明细'
    configured = cli('view-get-group', *args, '--view-id', view)['group']
    if not configured:
        cli('view-set-group', *args, '--view-id', view,
            '--json', encode({'group_config': [{'field': ORDER, 'desc': False}]}))
        configured = cli('view-get-group', *args, '--view-id', view)['group']
    if len(configured) != 1 or configured[0]['field'] != ORDER:
        raise ValueError(f'订单分组核对失败：{args[-1]} / {view} / {configured}')


def main():
    parser = argparse.ArgumentParser(description='按客户建档财年清洗 ERP 销售订单明细')
    parser.add_argument('xlsx', type=Path)
    parser.add_argument('--company-column', default='公司名称', help='用于隔离公司记录的源字段名')
    parser.add_argument('--fiscal-start-month', type=int, default=7, choices=range(1, 13))
    parser.add_argument('--reversal-column', help='经业务验证的冲销标记字段')
    parser.add_argument('--reversal-value', help='经业务验证的冲销标记值')
    target = parser.add_mutually_exclusive_group(required=True)
    target.add_argument('--base', help='读取该 Base 的渠道映射表，并写入同一个 Base')
    target.add_argument('--mapping-base', help='仅从该 Base 读取渠道映射表并核查 Excel')
    args = parser.parse_args()
    if bool(args.reversal_column) != bool(args.reversal_value):
        parser.error('冲销字段和标记值须一起指定')
    columns, rows = read_source(args.xlsx, args.company_column, args.reversal_column)
    base = args.base or args.mapping_base
    mapping_args = ('--base-token', base, '--table-id', MAP_TABLE)
    tables = cli('table-list', '--base-token', base)['tables']
    mapping_table = next((t for t in tables if t['name'] == MAP_TABLE), None)
    if not mapping_table:
        raise ValueError('目标 Base 缺少「渠道映射表」，请使用已有唯一映射规则的 Base')
    mapping = read_remote(mapping_args, MAP_FIELDS, mapping_table['records_count'])
    detail, quality, unmapped = details(rows, columns, channel_index([r[1] for r in mapping.values()]),
                                       args.company_column, args.fiscal_start_month)
    offset_audit = {}
    if args.reversal_column:
        _, offset_audit = cancel_matched_offsets(detail, args.company_column,
                                                 args.reversal_column, args.reversal_value)
    if not any(detail.values()):
        raise ValueError('源工作簿没有明细')
    print(encode({'财年明细': {k: len(v) for k, v in detail.items()},
                  '抵消': offset_audit, '清洗': quality, '未映射来源种类': len(unmapped),
                  '渠道明细': Counter(r['渠道大类'] for rows in detail.values() for r in rows)}), flush=True)
    material_tables = [t for t in tables if t['name'] in MATERIAL_MAP_NAMES]
    if len(material_tables) > 1:
        raise ValueError('检测到两张物料方案映射表，请确认唯一规则来源')
    material_mapping = material_tables[0] if material_tables and MATERIAL_CODE in columns else None
    if material_mapping:
        if MATERIAL_SCHEME in columns:
            raise ValueError('ERP 源表已有「解决方案」字段，请先确认如何保留源字段')
        print(encode({'物料方案明细': material_scheme_counts(base, material_mapping, detail)}), flush=True)
    if not args.base:
        return
    storage = (ORDER, KEY, FY, *(n for n in columns if n != ORDER))
    scheme_fields = (MATERIAL_SCHEME,) if material_mapping else ()
    options = {FY: set(detail)}
    jobs = []
    for fiscal, rows in detail.items():
        ensure_schema(base, fiscal, storage, options)
        channel_names = ensure_channel_formulas(base, fiscal, columns)
        if material_mapping:
            ensure_material_scheme_formula(base, fiscal, material_mapping['name'])
        names = (*storage, *channel_names, *scheme_fields)
        channel_headers = dict(zip(DETAIL_CHANNEL, channel_names))
        visible = tuple(channel_headers.get(n, n) for n in DETAIL_VISIBLE
                        if channel_headers.get(n, n) in names)
        jobs.append((base, fiscal, rows, names, (KEY, FY), visible))
    with ThreadPoolExecutor(max_workers=3) as pool:
        for _ in pool.map(lambda job: sync_table(*job), jobs):
            pass


if __name__ == '__main__':
    main()
