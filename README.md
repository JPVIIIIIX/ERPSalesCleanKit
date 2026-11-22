# ERP 销售订单清洗 Skill

把中文 ERP 销售明细拆成财年表。识别精确重复，保留原字段和明细金额，通过飞书渠道子表联动显示两个渠道结果。仓库提供脚本和使用说明，使用者自备 ERP 导出和飞书权限。

## 获取

```sh
git clone https://github.com/JPVIIIIIX/erp-sales-clean-skill.git
cd erp-sales-clean-skill
python3 -m pip install -r requirements.txt
```

需要 Python 3.10+、[官方 lark-cli](https://github.com/larksuite/cli) 以及目标多维表格权限。将项目作为 Skill 使用时，提供本目录的 `SKILL.md` 和配套脚本。

## 准备数据

Excel 第一行至少包括「销售订单编号」「销售订单行号」「客户建档日期」「客户来源」和一个公司字段。公司字段默认叫「公司名称」，可用 `--company-column` 指定现有表头。源文件其他列随明细保留。请保持标识列为 Excel 文本以保留前导零。

在自己的目标 Base 里准备一张名为「渠道映射表」的数据表，包含「ERP客户来源」（文本主字段）、「渠道大类」和「渠道子来源」（单选）。每种原始来源只留一条规则。用户自行维护渠道规则。

## 核查及导入

```sh
python3 scripts/clean_sales.py "/path/orders.xlsx" --mapping-base "YOUR_BASE_TOKEN"
python3 scripts/clean_sales.py "/path/orders.xlsx" --base "YOUR_BASE_TOKEN"
```

第一条读取飞书规则并输出核查统计。第二条在指定 Base 中按客户建档财年写入明细；额外选项见 `--help`。同一导出重跑保持表头顺序一致。目标财年表已经有其他范围的数据时，脚本会拒绝合并。

如已核实业务系统的冲销语义，可选用 `--reversal-column "调整标记" --reversal-value "冲销"`。这两个名称只是示例。启用时需要客户编号、物料编码、单位、集团及交易币种、两种含税金额和数量。先只读核查排除范围，再写飞书。
