---
name: ccb-bill-analyzer
description: Analyze China Construction Bank (CCB) credit card bills and generate Excel summaries with interactive HTML visualization reports. Use when the user provides a CCB bill file (.txt or .csv), asks to parse credit card transactions, categorize expenses, create spending dashboards, calendar views, or get optimization suggestions. Trigger: 建行信用卡, 信用卡账单, CCB bill, 消费分析, 账单可视化.
---

# CCB Bill Analyzer

Parses China Construction Bank credit card transaction records and bank account transaction statements, producing structured Excel summaries and interactive HTML visualization reports.

## Workflow

### Step 1: Locate the input file

Ask the user for the bill file path if not provided. Supported formats:
- **CSV**: Comma-separated, GBK encoded (CCB credit card bill format)
- **TXT**: Fixed-width, GBK encoded (CCB credit card bill format)
- **PDF**: CCB savings/debit account transaction statement (交易流水明细清单), using pdfplumber table extraction

Default paths to check:
- `C:\Users\Admin\Desktop\downfile.csv`
- `C:\Users\Admin\Downloads\downfile.txt`

### Step 2: LLM-based merchant classification (New!)

Instead of relying on simple keyword matching, the skill now uses an LLM to analyze all 140+ merchant names in the raw data and build accurate category rules before visualization:

1. **Extract**: Parse the CSV/TXT to extract all unique merchant names
2. **Analyze**: Use the LLM to classify each merchant into the correct category based on the actual business name (stripping payment method prefixes like `财付通-微信支付-`, `支付宝-支付宝-消费-` etc.)
3. **Update**: Write the comprehensive `MERCHANT_CATEGORY_RULES` dictionary into `ccb_bill_analyzer.py`
4. **Generate**: Run the analyzer script with the updated rules

This ensures categories like "网购电商" don't incorrectly capture every `微信支付`/`支付宝` transaction regardless of what was actually purchased.

### Step 3: Run the analyzer

```bash
python "<skill_dir>/ccb_bill_analyzer.py" "<input_file_path>"
```

Replace `<skill_dir>` with the directory containing this SKILL.md file.

### Step 4: Verify output

Two files are generated in the same directory as the input:
- `信用卡账单汇总_YYYYMMDD_HHMMSS.xlsx`
- `信用卡账单可视化_YYYYMMDD_HHMMSS.html`

Open the HTML file in a browser to view the interactive report.

## Script

The analyzer script is `ccb_bill_analyzer.py` in this same directory.

### Dependencies

```bash
pip install pandas numpy openpyxl pdfplumber
```

### Classification architecture

The classification uses a three-layer approach:

1. **Prefix stripping**: Payment gateway prefixes (`财付通-微信支付-`, `支付宝-支付宝-消费-`, etc.) are stripped to reveal the actual merchant name
2. **Cleaned name matching**: The stripped merchant name is matched against `MERCHANT_CATEGORY_RULES` keywords
3. **Fallback matching**: If no match is found, the original (unstripped) name is matched as fallback
4. **Special cases**: Cross-bank consumption (`跨行消费`), refunds, and AMEX rewards are handled explicitly before keyword matching

### Category rules

Transactions are classified by matching the cleaned merchant name (after stripping payment method prefixes) against these keywords, prioritized in list order:

| Category | Keywords |
|----------|----------|
| 餐饮美食 | Cotti, 咖啡, 餐厅, 美食, 餐饮, 火锅, 快餐, 奶茶, 面包, 蛋糕, 零食, 小吃, 面, 抄手, 米线, 汉堡, 肯德基, 金拱门, 烧烤, 麻辣烫, 蛋烘糕, 冰淇淋, 蜜雪, 烘焙, 餐馆, 酒, 厨, 锅, 煲, 餐吧, 星巴克, 拉扎斯, 饿了么, 外卖, and 50+ more food keywords |
| 日用百货 | 超市, 便利店, 沃尔玛, 山姆, Sam, 永辉, 罗森, 大润发, RT-Mart, 邻你, 百货, 购物中心, 万象城, 康成投资, 汽车用品 |
| 网购电商 | 拼多多, 天猫, 京东, 淘宝, 电商, 平台商户, Nuvei, Smart2Pay, 得物, 小红书, 抖音, 直播, 电子商务, 天猫供应链 |
| 交通出行 | 地铁, 公交, 停车, 加油, 打车, 滴滴, 嘀嘀, 高德, 携程, 中石油, 充电, 新能源, 智慧交通, 交投, 通行宝 |
| 医疗健康 | 医院, 药房, 医药, 健康, 体检, 诊所 |
| 教育培训 | 大学, 学校, 教育, 培训, 课程, 图书 |
| 休闲娱乐 | 电影, 游戏, KTV, 酒吧, 旅游, 景点, 博物馆, 三星堆, 武侯祠, 大熊猫, 毕棚沟, 玩具, 泡泡玛特 |
| 住房物业 | 物业, 房租, 水电, 燃气, 宽带, 国网, 国家电网, 机电工程 |
| 通讯数码 | 手机, 电脑, 数码, 小米, 华为, 网易UU, 加速器 |
| 服饰美容 | 服装, 服饰, 美容, 美发, 鞋, 包, 饰品, 优衣库, 周大福, 欧莱雅, IFS, 际华 |
| 投资理财 | 投资, 理财, 基金, 保险, 证券, 银行转账 |
| 还款退款 | 还款, 退款, 退货, 返还, 入账, 退税, 返现 (automatically detected from negative amounts) |

## HTML Report features

- **Stat cards**: Total spending, transaction count, monthly average, largest transaction
- **6 global charts**: Monthly trend, category pie, stacked comparison, TOP10 merchants, weekday distribution, transaction count
- **Large expense table**: All transactions >= 500 CNY or top 5%
- **Interactive calendar**: Month navigation, color-coded daily spending, click any day for transaction details modal
- **Monthly drill-down**: Per-month stats, daily trend, category breakdown, TOP merchants, weekday distribution
- **Optimization suggestions**: Auto-generated based on spending patterns

## Color palette (Nature journal style)

| Role | Color |
|------|-------|
| Primary | `#0F4D92` |
| Secondary | `#3775BA` |
| Green | `#8BCF8B` |
| Red | `#B64342` |
| Orange | `#E28E2C` |
| Teal | `#42949E` |
| Violet | `#9A4D8E` |

## Troubleshooting

- **GBK decode error**: The file must be GBK encoded (standard for CCB exports). If parsing fails, try the CSV format instead of TXT.
- **No data parsed**: Check that the file contains actual transaction rows, not just headers.
- **Missing dependencies**: Run `pip install pandas numpy openpyxl pdfplumber`.
- **Unicode print error**: The script uses ASCII-safe print statements for Windows console compatibility.
- **Classification off**: Rerun Step 2 (LLM re-analysis) to rebuild `MERCHANT_CATEGORY_RULES` with updated merchant names from the new bill file. Different billing periods may contain different merchants.
- **PDF garbled text**: Some CCB PDFs use fonts with incomplete ToUnicode CMaps, causing garbled Chinese characters. The table structure (dates, amounts, balances) is always preserved. Merchant name matching may be less accurate for PDF inputs.
- **PDF is not a credit card bill**: The PDF parser is designed for CCB savings account transaction statements (交易流水明细清单). The output labels will still say "信用卡" but the analysis will work correctly on the transaction data.
