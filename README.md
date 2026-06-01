# Order Consolidation Checker

自动分析 NetSuite 订单，识别可合并发货的订单组，并将结果写回 Google Sheet。

---

## 功能概览

| 步骤 | 说明 |
|------|------|
| 读取数据 | 从 Google Sheet 追踪表 + NetSuite 导出的 IF 表（`.xls`）读取订单信息 |
| 维护历史 | 将 IF 数据 upsert 到 `IF_history` tab，自动清理已完成的 SO |
| 合单分析 | 按账号 + ZIP 分组，基于发货日 / 抓货日时间窗口判断哪些订单可以合并发货 |
| 写回结果 | 将合单建议写入 Google Sheet 的 `Note` 列，保留人工已有备注 |
| 本地报告 | 导出 CSV 报告至当前目录 |

---

## 环境要求

- Python 3.8+
- 依赖库：

```bash
pip install pandas gspread google-auth
```

---

## 使用方法

### 1. 准备文件

| 文件 | 说明 |
|------|------|
| `ItemFulfillments.xls` | 从 NetSuite 导出的当日 IF 表（Excel XML 格式） |
| `credentials.json` | Google Service Account 凭证文件 |

> NetSuite 导出时请选择 **Excel**（SpreadsheetML `.xls`）格式，不是 CSV 或二进制 xlsx。

### 2. 配置 `CONFIG`

打开 `consolidation_checker.py`，根据实际情况修改顶部的 `CONFIG` 字典：

```python
CONFIG = {
    "sheet_id":       "你的 Google Sheet ID",
    "tracking_tab":   "2026",           # 追踪表 tab 名
    "if_history_tab": "IF_history",     # IF 历史 tab（首次运行自动创建）
    "if_xls_path":    "ItemFulfillments.xls",
    "credentials_file": "credentials.json",
    ...
}
```

### 3. 运行

```bash
# 交互模式（运行后询问是否写回）
python consolidation_checker.py

# 直接写回 Google Sheet
python consolidation_checker.py --write

# 仅分析，不写回
python consolidation_checker.py --dry-run
```

---

## Google Sheet 格式要求

| 要求 | 说明 |
|------|------|
| 表头行 | 第 **3** 行为列名，第 4 行起为数据 |
| 必要列 | `Order Date`、`ACCT#`、`ZIP CODE`、`SO#`、`Status`、`Note`、`normal shipped date`、`QTY` |

---

## 合单逻辑

```
同一 ACCT# + ZIP CODE
    └─ 发货日相差 ≤ 3 天  或  抓货日相差 ≤ 3 天
        └─ 排除：一方有 IF 记录、另一方无 IF 且 QTY > 300
            └─ Clique 检验：组内任意两张都满足时间窗口才归为同一合单组
```

**参与分析的 Status（精确匹配，不区分大小写）：**

| Status | 说明 |
|--------|------|
| `stock order` | 未处理，可合并 |
| `fcsrepleshing` | 补货中，可合并 |
| `hold` | 等待中，可合并 |
| `send`（今日） | 当天给了 shipping，还来得及合并 |

---

## Note 列写入格式

脚本以 `[自动]` 为标识追加建议，不覆盖人工内容：

```
人工备注  [自动]可与 SO12345、SO67890 合发 | 已抓货:5/28
```

再次运行时，`[自动]` 之后的内容会被刷新，人工部分保持不变。

---

## 输出示例

```
=======================================================
  ✅ 可合单 SO：6 张，共 3 组
  ⚠️  放弃合并警告：1 张
  🔴 紧急/过期：2 张
=======================================================

合单明细：
  合单组1  🔴 紧急
    📦有IF  QTY:120  截止6/2  SO:100001
      无IF  QTY:80   截止6/3  SO:100045
  合单组2
    📦有IF  QTY:200  截止6/5  SO:100078
    📦有IF  QTY:150  截止6/6  SO:100089
```

---

## 文件结构

```
.
├── consolidation_checker.py   # 主程序
├── credentials.json           # Google 凭证（不提交到 Git）
├── ItemFulfillments.xls       # 当日 NetSuite 导出（不提交到 Git）
└── consolidation_report_YYYYMMDD.csv  # 自动生成的本地报告
```

> `credentials.json` 和 `.xls` 文件包含敏感信息或临时数据，请添加到 `.gitignore`，不要提交。

---

## 常见错误

| 错误信息 | 原因 & 解决方法 |
|----------|----------------|
| `❌ 找不到 IF 表文件` | `ItemFulfillments.xls` 不在当前目录，检查路径或 `CONFIG["if_xls_path"]` |
| `❌ IF 表解析失败` | 导出格式不是 SpreadsheetML，重新从 NetSuite 选择 Excel 格式导出 |
| `❌ 找不到 credentials.json` | 缺少 Google Service Account 凭证，参考 [官方文档](https://docs.gspread.org/en/latest/oauth2.html) 创建 |
| `⚠️ 追踪表缺少以下列` | `CONFIG` 中的列名与 Sheet 实际列名不一致，对照 Sheet 第3行修改 `CONFIG` |
