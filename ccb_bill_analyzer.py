#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
建行信用卡账单分析工具 - Binance专业交易终端风格
支持TXT和CSV两种格式输入
输出:Excel汇总表+HTML交互式可视化报告
"""

import re
import os
import json
import requests
import pandas as pd
import numpy as np
from datetime import datetime
from collections import defaultdict
import pdfplumber

# Payment prefixes to strip before classification
PAYMENT_PREFIXES = [
    '财付通-微信支付-', '财付通-财付通-', '财付通-',
    '支付宝-支付宝-消费-', '支付宝-支付宝-', '支付宝-',
    '手机银行龙支付-',
    '上海 跨行消费 ', '上海 跨行退货 ',
]

# Merchant prefixes that indicate specific categories
MERCHANT_CATEGORY_RULES = {
    '餐饮美食': [
        'Cotti', '咖啡', '餐厅', '美食', '餐饮', '火锅', '快餐', '奶茶', '面包', '蛋糕',
        '零食', '小吃', '饭', '面', '抄手', '馄饨', '饺子', '粥', '茶', '饮',
        '汉堡', '肯德基', '麦当劳', '金拱门', '比萨', '匹萨',
        '烤鱼', '烧烤', '麻辣烫', '冒菜', '米线', '米粉',
        '蛋烘糕', '烘焙', '甜品', '冰淇淋', '蜜雪', '鲜果', '果汁',
        '鸡柳', '鸡排', '炸鸡', '花甲', '龙虾', '海鲜',
        '牛肉', '羊肉', '鸡肉', '烤鸭', '烧鹅',
        '炒饭', '盖饭', '拌饭', '便当', '外卖',
        '拉扎斯', '饿了么', '食堂', '餐馆', '菜馆', '食府',
        '酒', '酒吧', '厨', '锅', '煲', '烤',
        '烘焙', '糕点', '面包店', '蛋糕店',
        '牛肉火锅', '鸡煲', '打边炉', '清远鸡',
        '卤味', '鸭脖', '熟食', '凉菜',
        '藏餐', '永利藏餐', '野人',
        '鑫鑫回味', '良品', '味',
        '生榨', '树夏', '椰子',
        '杂吃', '酸菜鱼', '水煮',
        '川菜', '湘菜',
        '半坡', '无恙', '鹤鸣', '灵感之茶',
        '石棉', '杨记', '木炭',
        '泉盛餐饮', '餐吧', '星巴克',
    ],
    '日用百货': [
        '超市', '便利店', '沃尔玛', '山姆', 'Sam', '永辉', '红旗', '舞东风',
        '全家', '罗森', '7-11', '711', '大润发', 'RT-Mart', '华润',
        '生活', '日用', '杂货', '百货', '购物中心',
        '万象汇', '万象城', '商场', 'mall', 'Mall',
        '邻你超市', '华渝', '上蔬',
        '康成投资', '沃尔玛',
        '广州可备', '汽车用品', '车品',
    ],
    '网购电商': [
        '拼多多', '天猫', '京东', '淘宝', '电商', '平台商户', 'Nuvei', 'Smart2Pay',
        '得物', '小红书', '抖音', '快手', '直播',
        '两只小马', '电子商务',
        '浙江天猫供应链',
        '际华官方旗舰店',
    ],
    '交通出行': [
        '地铁', '公交', '停车', '加油', '打车', '滴滴', '嘀嘀', '高德', '百度地图',
        '交通', '出行', '租车', '火车', '飞机', '携程', '去哪儿', '航',
        '中石油', '昆仑网电', '充电', '新能源',
        '通行宝', '智慧交通', '数字交通',
        '交投', '停车公司', '停车场',
        '四川蜀道', '湖北交投',
        '江苏通行宝', '重庆数字交通',
        '骑安', '滴滴',
        '顺易通', '微泊云',
    ],
    '医疗健康': [
        '医院', '药房', '医药', '健康', '体检', '诊所', '牙科', '眼科',
        '都江堰市人民医院', '兴福药房',
    ],
    '教育培训': [
        '大学', '学校', '教育', '培训', '课程', '书本', '图书', '文具',
        '四川农业大学',
    ],
    '休闲娱乐': [
        '电影', '游戏', '娱乐', 'KTV', '酒吧', '旅游', '景点', '门票',
        '视频', '音乐', '演出', '展览', '博物馆', '乐园',
        '三星堆', '武侯祠', '大熊猫', '熊猫基地',
        '毕棚沟', '汶川无忧谷', '北湖旅游',
        '东方佛都', '奇趣实验室',
        '玩具', '泡泡玛特',
        '星岩文化', '文化传播',
    ],
    '住房物业': [
        '物业', '房租', '水电', '燃气', '宽带', '电信', '移动', '联通', '住房',
        '国网', '国家电网', '网上国网',
        '机电工程', '民盛机电',
        '汇通金财',
    ],
    '通讯数码': [
        '手机', '电脑', '数码', '苹果', '华为', '小米', '通讯', '话费',
        '网易UU', '加速器',
        '速卡科技',
        '景明科技',
    ],
    '服饰美容': [
        '服装', '服饰', '美容', '美发', '化妆', '护肤', '鞋', '包', '饰品',
        '优衣库', '周大福', '欧莱雅',
        '兵立王', '品牌管理',
        '文旅宽窄巷子',
        'IFS', '际华',
    ],
    '投资理财': [
        '投资', '理财', '基金', '股票', '保险', '证券', '银行转账',
        '投', '资本',
    ],
    '还款退款': [
        '还款', '退款', '退货', '返还', '入账', '退税', '返现', '历史',
        '银联云闪付',
    ],
    '其他': []
}


def parse_csv(filepath):
    records = []
    with open(filepath, 'r', encoding='gbk', errors='ignore') as f:
        lines = f.readlines()
    header_idx = None
    for i, line in enumerate(lines):
        if '交易日期' in line or 'ճ' in line:
            header_idx = i
            break
    if header_idx is None:
        header_idx = 13
    for line in lines[header_idx+1:]:
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split(',')]
        if len(parts) < 7:
            continue
        try:
            record = {
                '交易日期': parts[0].strip(),
                '记账日期': parts[1].strip(),
                '卡号': parts[2].strip().replace("'", ""),
                '交易类型': parts[3].strip(),
                '币种': parts[4].strip(),
                '金额': float(parts[5].strip()),
                '商户名称': parts[6].strip()
            }
            if len(record['交易日期']) == 8 and record['交易日期'].isdigit():
                records.append(record)
        except (ValueError, IndexError):
            continue
    return pd.DataFrame(records)


def parse_txt(filepath):
    records = []
    with open(filepath, 'r', encoding='gbk', errors='ignore') as f:
        content = f.read()
    lines = content.split('\n')
    current_merchant = ''
    current_record = None
    for line in lines:
        line = line.strip()
        if not line:
            continue
        date_pattern = r'^(\d{8})\s+(\d{8})\s+(\d+)\s+'
        match = re.match(date_pattern, line)
        if match:
            if current_record and current_merchant:
                current_record['商户名称'] = current_merchant
                records.append(current_record)
            parts = line.split()
            if len(parts) >= 6:
                try:
                    current_record = {
                        '交易日期': parts[0],
                        '记账日期': parts[1],
                        '卡号': parts[2],
                        '交易类型': parts[3],
                        '币种': parts[4],
                        '金额': float(parts[5]),
                        '商户名称': ''
                    }
                    if len(parts) > 6:
                        current_merchant = ' '.join(parts[6:])
                    else:
                        current_merchant = ''
                except (ValueError, IndexError):
                    current_record = None
                    current_merchant = ''
        else:
            if current_record is not None:
                current_merchant += ' ' + line
    if current_record and current_merchant:
        current_record['商户名称'] = current_merchant
        records.append(current_record)
    return pd.DataFrame(records)


def parse_pdf(filepath):
    records = []
    with pdfplumber.open(filepath) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if not tables:
                continue
            for table in tables:
                for row in table:
                    if not row or len(row) < 9:
                        continue
                    # Column 0 = date (YYYY-MM-DD), Column 3 = amount, Column 8 = counterparty name
                    date_str = str(row[0]).strip() if row[0] else ''
                    amt_str = str(row[3]).strip() if row[3] else ''
                    merchant = str(row[8]).strip() if row[8] else ''
                    # Clean merchant name: remove newlines and extra spaces
                    merchant = re.sub(r'\s+', '', merchant)
                    # Detect if this is a header row (starts with non-date text)
                    if not re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
                        continue
                    try:
                        amount = float(amt_str.replace(',', ''))
                    except (ValueError, AttributeError):
                        continue
                    if not merchant or merchant == '-------------------':
                        merchant = ''
                    # Convert YYYY-MM-DD to YYYYMMDD for consistency
                    date_compact = date_str.replace('-', '')
                    # Flip sign: debit card outflows (negative in source) → positive (spending),
                    # incomes (positive in source) → negative, matching credit card convention
                    unified_amount = -amount
                    record = {
                        '交易日期': date_compact,
                        '记账日期': date_compact,
                        '卡号': 'PDF',
                        '交易类型': '消费' if amount < 0 else '转入',
                        '币种': '人民币',
                        '金额': unified_amount,
                        '商户名称': merchant
                    }
                    records.append(record)
    if not records:
        raise ValueError("未从PDF解析到有效数据")
    return pd.DataFrame(records)


def parse_bill(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.csv':
        df = parse_csv(filepath)
    elif ext == '.pdf':
        df = parse_pdf(filepath)
    else:
        df = parse_txt(filepath)
    if df.empty:
        raise ValueError("未解析到有效数据")
    df['交易日期'] = pd.to_datetime(df['交易日期'], format='%Y%m%d')
    df['记账日期'] = pd.to_datetime(df['记账日期'], format='%Y%m%d')
    df['年月'] = df['交易日期'].dt.to_period('M')
    df['月份'] = df['交易日期'].dt.strftime('%Y-%m')
    df['星期'] = df['交易日期'].dt.day_name()
    return df


def clean_merchant_name(name):
    if not name or pd.isna(name):
        return ''
    name = str(name)
    for prefix in PAYMENT_PREFIXES:
        if name.startswith(prefix):
            return name[len(prefix):]
    return name


def classify_merchant(merchant_name, amount):
    if not merchant_name or pd.isna(merchant_name):
        return '其他'
    merchant_str = str(merchant_name)
    merchant_lower = merchant_str.lower()

    # Handle negative amounts (refunds/repayments)
    if amount < 0:
        if any(kw in merchant_str for kw in ['还款', '转入', '转账']):
            return '还款退款'
        if any(kw in merchant_str for kw in ['退款', '退货', '返还', '入账', '退税', '返现', '历史']):
            return '还款退款'
        return '还款退款'

    # Clean the merchant name by stripping payment prefixes
    clean_name = clean_merchant_name(merchant_name)

    # Also check the original name for payment-platform-specific categories
    if '运通版欢享卡' in merchant_str:
        return '还款退款'
    if '网银跨行还款' in merchant_str:
        return '投资理财'
    if '跨行消费' in merchant_str:
        # These are specific utility/bill payments
        if '网上国网' in clean_name:
            return '住房物业'
        if '银联云闪付' in merchant_str:
            return '网购电商'
    if '跨行退货' in merchant_str:
        return '还款退款'

    # Try matching the cleaned name first
    if clean_name:
        for category, keywords in MERCHANT_CATEGORY_RULES.items():
            if category in ('其他', '还款退款'):
                continue
            for kw in keywords:
                if kw.lower() in clean_name.lower():
                    return category

    # Fallback: match against original name
    for category, keywords in MERCHANT_CATEGORY_RULES.items():
        if category in ('其他', '还款退款'):
            continue
        for kw in keywords:
            if kw.lower() in merchant_lower:
                return category

    return '其他'


def apply_classification(df):
    df['分类'] = df.apply(lambda row: classify_merchant(row['商户名称'], row['金额']), axis=1)
    return df


def generate_excel(df, output_path):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        expense_df = df[df['金额'] > 0].copy()
        monthly_category = expense_df.groupby(['月份', '分类'])['金额'].sum().unstack(fill_value=0)
        monthly_category['月度总计'] = monthly_category.sum(axis=1)
        monthly_category = monthly_category.sort_index()
        monthly_category.to_excel(writer, sheet_name='月度分类汇总')
        
        detail_df = df[['交易日期', '记账日期', '商户名称', '交易类型', '金额', '分类']].copy()
        detail_df = detail_df.sort_values('交易日期', ascending=False)
        detail_df.to_excel(writer, sheet_name='交易明细', index=False)
        
        monthly_stats = expense_df.groupby('月份').agg(
            总支出=('金额', 'sum'),
            交易笔数=('金额', 'count')
        ).round(2)
        monthly_stats['日均消费'] = monthly_stats['总支出'] / expense_df.groupby('月份')['交易日期'].nunique()
        monthly_stats = monthly_stats.round(2)
        monthly_stats.to_excel(writer, sheet_name='月度趋势')
        
        category_summary = expense_df.groupby('分类').agg(
            总金额=('金额', 'sum'),
            交易笔数=('金额', 'count'),
            占比=('金额', lambda x: x.sum() / expense_df[expense_df['金额'] > 0]['金额'].sum() * 100)
        ).round(2)
        category_summary = category_summary.sort_values('总金额', ascending=False)
        category_summary.to_excel(writer, sheet_name='分类汇总')
        
        threshold = max(500, expense_df['金额'].quantile(0.95))
        large_exp = expense_df[expense_df['金额'] >= threshold].sort_values('金额', ascending=False).copy()
        large_exp['交易日期'] = large_exp['交易日期'].dt.strftime('%Y-%m-%d')
        large_exp_out = large_exp[['交易日期', '商户名称', '分类', '金额']].copy()
        large_exp_out.to_excel(writer, sheet_name='大额支出明细', index=False)
    return output_path


def generate_suggestions(df, category_summary, monthly_category):
    suggestions = []
    total = df['金额'].sum()
    if '餐饮美食' in category_summary.index:
        food_pct = category_summary['餐饮美食'] / total * 100
        if food_pct > 30:
            suggestions.append({'type': 'warning', 'text': '餐饮支出占比{:.1f}%,建议控制外食/外卖频率以优化资金利用率'.format(food_pct)})
    if '网购电商' in category_summary.index:
        shop_pct = category_summary['网购电商'] / total * 100
        if shop_pct > 25:
            suggestions.append({'type': 'warning', 'text': '网购支出占比{:.1f}%,建议引入冷却期机制避免冲动消费'.format(shop_pct)})
    large_tx = df[df['金额'] > df['金额'].quantile(0.95)]
    if len(large_tx) > 0:
        suggestions.append({'type': 'danger', 'text': '监测到{}笔大额异常支出(均>¥{:.0f}),请复盘资金流向'.format(len(large_tx), large_tx['金额'].min())})
    if len(monthly_category) >= 2:
        monthly_totals = monthly_category.sum(axis=1)
        if monthly_totals.iloc[-1] > monthly_totals.iloc[-2] * 1.3:
            suggestions.append({'type': 'danger', 'text': '近期支出波动剧烈,环比增长超30%,建议启动预算控制'})
    suggestions.append({'type': 'success', 'text': '系统计算得出您的基础月均开支基准线为¥{:,.0f},可作为下季度预算设定参考'.format(df.groupby("月份")["金额"].sum().mean())})
    return suggestions


def generate_html(df, output_path):
    expense_df = df[df['金额'] > 0].copy()
    total_expense = expense_df['金额'].sum()
    total_transactions = len(expense_df)
    months = sorted(expense_df['月份'].unique())
    avg_monthly = total_expense / len(months) if months else 0
    
    monthly_category = expense_df.groupby(['月份', '分类'])['金额'].sum().unstack(fill_value=0)
    monthly_category = monthly_category.reindex(sorted(monthly_category.index))
    category_summary = expense_df.groupby('分类')['金额'].sum().sort_values(ascending=False)
    top_merchants = expense_df.groupby('商户名称')['金额'].sum().sort_values(ascending=False).head(10)
    
    weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    weekday_cn = {'Monday': '周一', 'Tuesday': '周二', 'Wednesday': '周三', 'Thursday': '周四',
                  'Friday': '周五', 'Saturday': '周六', 'Sunday': '周日'}
    weekday_data = expense_df.groupby('星期')['金额'].sum()
    suggestions = generate_suggestions(expense_df, category_summary, monthly_category)
    threshold = max(500, expense_df['金额'].quantile(0.95))
    large_expenses = expense_df[expense_df['金额'] >= threshold].sort_values('金额', ascending=False)
    max_amount = expense_df['金额'].max()
    
    monthly_data_js = {}
    for m in months:
        m_df = expense_df[expense_df['月份'] == m].copy()
        m_total = m_df['金额'].sum()
        m_count = len(m_df)
        m_avg = m_total / m_df['交易日期'].nunique() if m_df['交易日期'].nunique() > 0 else 0
        m_cat = m_df.groupby('分类')['金额'].sum().sort_values(ascending=False)
        m_cat_data = [{'name': str(k), 'value': round(float(v), 2)} for k, v in m_cat.items()]
        m_top = m_df.groupby('商户名称')['金额'].sum().sort_values(ascending=False).head(10)
        m_top_names = [str(n)[:60] for n in m_top.index[::-1]]
        m_top_values = [round(float(v), 2) for v in m_top.values[::-1]]
        m_daily = m_df.groupby('交易日期')['金额'].sum()
        m_daily_dates = [str(d.date()) if hasattr(d, 'date') else str(d)[:10] for d in m_daily.index]
        m_daily_values = [round(float(v), 2) for v in m_daily.values]
        m_weekday = m_df.groupby('星期')['金额'].sum()
        m_weekday_values = [round(float(m_weekday.get(d, 0)), 2) for d in weekday_order]
        
        # 提取全月所有流水清单，供侧滑抽屉展示
        all_month_txns = []
        for _, row in m_df.sort_values('交易日期', ascending=False).iterrows():
            d = row['交易日期'].strftime('%Y-%m-%d') if hasattr(row['交易日期'], 'strftime') else str(row['交易日期'])[:10]
            all_month_txns.append({
                'date': d,
                'merchant': str(row['商户名称']),
                'category': str(row['分类']),
                'amount': round(float(row['金额']), 2)
            })

        calendar_data = {}
        for date_str, group in m_df.groupby('交易日期'):
            d = date_str.strftime('%Y-%m-%d') if hasattr(date_str, 'strftime') else str(date_str)[:10]
            day_total = group['金额'].sum()
            calendar_data[d] = {'total': round(day_total, 2), 'count': len(group)}
        
        m_large = m_df[m_df['金额'] >= threshold].sort_values('金额', ascending=False)
        m_large_rows = []
        for _, row in m_large.iterrows():
            d = row['交易日期'].strftime('%Y-%m-%d') if hasattr(row['交易日期'], 'strftime') else str(row['交易日期'])[:10]
            m_large_rows.append({'date': d, 'merchant': str(row['商户名称'])[:30], 'category': row['分类'], 'amount': round(float(row['金额']), 2)})
        
        monthly_data_js[m] = {
            'total': round(m_total, 2), 'count': m_count, 'avg': round(m_avg, 2),
            'max': round(float(m_df['金额'].max()), 2) if not m_df.empty else 0,
            'max_merchant': str(m_df.loc[m_df['金额'].idxmax(), '商户名称'])[:25] if not m_df.empty else "无",
            'cat_data': m_cat_data, 'top_names': m_top_names, 'top_values': m_top_values,
            'daily_dates': m_daily_dates, 'daily_values': m_daily_values,
            'weekday_values': m_weekday_values, 'calendar': calendar_data,
            'large_rows': m_large_rows, 'all_txns': all_month_txns,
            'year': int(m[:4]), 'month': int(m[5:7])
        }
    
    months_list = json.dumps(list(monthly_category.index))
    monthly_totals = json.dumps([round(x, 2) for x in monthly_category.sum(axis=1)])
    pie_data = json.dumps([{'name': str(k), 'value': round(float(v), 2)} for k, v in category_summary.items()])
    stacked_series = json.dumps([
        {'name': str(col), 'type': 'bar', 'stack': 'total', 'data': [round(float(x), 2) for x in monthly_category[col]]}
        for col in monthly_category.columns
    ])
    top_merchant_names = json.dumps([str(m)[:60] for m in top_merchants.index[::-1]])
    top_merchant_values = json.dumps([round(float(x), 2) for x in top_merchants.values[::-1]])
    weekday_names = json.dumps([weekday_cn.get(d, d) for d in weekday_order])
    weekday_values = json.dumps([round(float(weekday_data.get(d, 0)), 2) for d in weekday_order])
    
    max_merchant = "无数据"
    if not expense_df.empty:
        max_amount_row = expense_df.loc[expense_df['金额'].idxmax()]
        max_merchant = str(max_amount_row['商户名称'])[:25]
        
    monthly_data_json = json.dumps(monthly_data_js, ensure_ascii=False)
    
    large_expense_rows = ''
    for _, row in large_expenses.iterrows():
        date_str = row['交易日期'].strftime('%Y-%m-%d') if hasattr(row['交易日期'], 'strftime') else str(row['交易日期'])[:10]
        merchant = str(row['商户名称'])[:20]
        amount = "{:,.2f}".format(row['金额'])
        category = row['分类']
        large_expense_rows += '<tr><td class="col-date">{}</td><td class="col-cat"><span class="badge">{}</span></td><td class="col-merch">{}</td><td class="col-amt">¥{}</td></tr>'.format(date_str, category, merchant, amount)
    
    suggestions_html = ''.join(
        '<div class="alert-box {}"><div class="alert-title">{}</div><div class="alert-desc">{}</div></div>'.format(
            s["type"], 
            "风险提示" if s["type"] == "danger" else ("系统建议" if s["type"] == "warning" else "运行结果"), 
            s["text"]
        )
        for s in suggestions
    )
    
    date_range_start = months[0] if months else 'N/A'
    date_range_end = months[-1] if months else 'N/A'
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    html_content = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>账单分析终端 | Bill Analytics Terminal</title>
    <script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
    <style>
        :root {
            --bg-base: #0B0E11;
            --bg-elevated: #181A20;
            --bg-highlight: #2B3139;
            --text-primary: #EAECEF;
            --text-secondary: #848E9C;
            --text-disabled: #5E6673;
            --border-color: #2B3139;
            --brand-color: #FCD535;
            --brand-hover: #F0B90B;
            --up-color: #0ECB81;
            --down-color: #F6465D;
            --warning-color: #C99400;
            --font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { background-color: var(--bg-base); color: var(--text-primary); font-family: var(--font-family); display: flex; height: 100vh; overflow: hidden; font-size: 22px; }
        ::-webkit-scrollbar { width: 4px; height: 4px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: var(--bg-highlight); border-radius: 2px; }
        ::-webkit-scrollbar-thumb:hover { background: var(--text-disabled); }
        
        /* Sidebar Navigation */
        .sidebar { width: 260px; background-color: var(--bg-elevated); border-right: 1px solid var(--border-color); display: flex; flex-direction: column; z-index: 10; }
        .logo-area { padding: 20px; font-size: 22px; font-weight: 700; color: var(--brand-color); letter-spacing: 1px; border-bottom: 1px solid var(--border-color); display: flex; align-items: center; gap: 8px; }
        .logo-area::before { content: ''; display: block; width: 12px; height: 12px; background-color: var(--brand-color); border-radius: 3px; transform: rotate(45deg); }
        .nav-menu { list-style: none; padding: 12px 0; flex: 1; }
        .nav-item { padding: 14px 20px; cursor: pointer; color: var(--text-secondary); font-weight: 500; transition: all 0.2s; display: flex; align-items: center; gap: 10px; border-left: 3px solid transparent; font-size: 20px; }
        .nav-item:hover { color: var(--text-primary); background-color: var(--bg-highlight); }
        .nav-item.active { color: var(--text-primary); background-color: rgba(252, 213, 53, 0.05); border-left-color: var(--brand-color); }
        .sidebar-footer { padding: 20px; border-top: 1px solid var(--border-color); color: var(--text-disabled); font-size: 20px; line-height: 1.5; }
        
        /* Main Content Area */
        .main-wrapper { flex: 1; display: flex; flex-direction: column; overflow: hidden; background-color: var(--bg-base); }
        .topbar { height: 68px; border-bottom: 1px solid var(--border-color); display: flex; align-items: center; justify-content: space-between; padding: 0 24px; background-color: var(--bg-elevated); }
        .topbar-title { font-size: 22px; font-weight: 600; color: var(--text-primary); text-transform: uppercase; letter-spacing: 0.5px;}
        .topbar-info { color: var(--text-secondary); display: flex; gap: 16px; align-items: center;}
        .info-tag { background: var(--bg-highlight); padding: 2px 10px; border-radius: 4px; color: var(--text-primary); }
        
        .content-scroll { flex: 1; overflow-y: auto; padding: 24px; scroll-behavior: smooth; }
        .view-section { display: none; animation: fadeIn 0.3s ease; }
        .view-section.active { display: block; }
        
        /* Dashboard Cards */
        .grid-row { display: grid; gap: 16px; margin-bottom: 16px; }
        .col-4 { grid-template-columns: repeat(4, 1fr); }
        .col-2 { grid-template-columns: repeat(2, 1fr); }
        .col-1 { grid-template-columns: 1fr; }
        
        .card { background-color: var(--bg-elevated); border: 1px solid var(--border-color); border-radius: 4px; padding: 18px; transition: border-color 0.2s;}
        .card.clickable { cursor: pointer; }
        .card.clickable:hover { border-color: var(--text-disabled); background-color: #1a1c24; }
        .card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
        .card-title { font-size: 22px; font-weight: 600; color: var(--text-secondary); text-transform: uppercase;}
        
        /* Stats Display */
        .stat-label { color: var(--text-secondary); margin-bottom: 4px; font-size: 22px;}
        .stat-value { font-size: 30px; font-weight: 700; color: var(--text-primary); font-variant-numeric: tabular-nums; line-height: 1.2;}
        .stat-value span { font-size: 22px; color: var(--text-secondary); font-weight: 500; margin-left: 2px; }
        .stat-trend { margin-top: 4px; font-size: 21px; }
        .txt-up { color: var(--up-color); }
        .txt-down { color: var(--down-color); }
        
        /* Charts */
        .chart-box { width: 100%; height: 280px; }
        .chart-box-tall { width: 100%; height: 300px; }
        .chart-box-xl { width: 100%; height: 400px; }
        
        /* Data Tables */
        .table-wrap { width: 100%; overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; text-align: left; }
        th { padding: 12px 14px; color: var(--text-secondary); font-weight: 500; border-bottom: 1px solid var(--border-color); white-space: nowrap; font-size: 21px;}
        td { padding: 14px; border-bottom: 1px solid var(--border-color); color: var(--text-primary); font-size: 22px; }
        tr:hover td { background-color: var(--bg-highlight); }
        tr:last-child td { border-bottom: none; }
        .col-amt { text-align: right; color: var(--down-color); font-weight: 600; font-variant-numeric: tabular-nums; }
        .badge { background-color: var(--bg-highlight); color: var(--text-secondary); padding: 4px 10px; border-radius: 4px; font-size: 20px; }
        
        /* Controls */
        .month-selector { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 16px; }
        .btn-tab { background: var(--bg-elevated); border: 1px solid var(--border-color); color: var(--text-secondary); padding: 10px 18px; border-radius: 4px; cursor: pointer; transition: 0.2s; font-weight: 500; font-size: 21px; }
        .btn-tab:hover { border-color: var(--text-disabled); color: var(--text-primary); }
        .btn-tab.active { background: var(--brand-color); border-color: var(--brand-color); color: var(--bg-base); font-weight: 600; }
        
        /* Calendar Heatmap (High Density) */
        .cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 4px; width: 100%; }
        .cal-header { text-align: center; color: var(--text-secondary); padding-bottom: 6px; font-size: 21px; font-weight: 600;}
        .cal-day { background-color: var(--bg-highlight); border-radius: 4px; padding: 6px; min-height: 68px; display: flex; flex-direction: column; justify-content: space-between; cursor: pointer; transition: 0.1s; border: 1px solid transparent; }
        .cal-day:hover { border-color: var(--brand-color); transform: scale(1.08); z-index: 2; box-shadow: 0 0 12px rgba(0,0,0,0.5);}
        .cal-day.empty { background: transparent; cursor: default; }
        .cal-day.empty:hover { border-color: transparent; transform: none; box-shadow: none;}
        .d-num { font-size: 21px; color: var(--text-disabled); font-weight: 600; line-height: 1;}
        .d-amt { font-size: 18px; text-align: right; font-variant-numeric: tabular-nums; line-height: 1; margin-top: 4px; font-weight: 500;}
        .legend-bar { display: flex; gap: 14px; margin-top: 14px; justify-content: flex-end; align-items: center; flex-wrap: wrap; }
        .legend-item { display: flex; align-items: center; gap: 5px; color: var(--text-secondary); font-size: 21px;}
        .l-box { width: 16px; height: 16px; border-radius: 2px; }
        
        /* Alerts */
        .alert-box { padding: 16px 20px; border-radius: 4px; margin-bottom: 12px; border-left: 4px solid; background-color: var(--bg-elevated); }
        .alert-box.danger { border-color: var(--down-color); }
        .alert-box.warning { border-color: var(--warning-color); }
        .alert-box.success { border-color: var(--up-color); }
        .alert-title { font-weight: 600; margin-bottom: 4px; color: var(--text-primary); font-size: 22px; }
        .alert-desc { color: var(--text-secondary); line-height: 1.6; font-size: 22px; }
        
        /* Right Drawer (Transaction History) */
        .drawer-overlay { display: none; position: fixed; inset: 0; background: rgba(11, 14, 17, 0.6); z-index: 100; backdrop-filter: blur(2px);}
        .drawer-overlay.show { display: block; animation: fadeIn 0.2s; }
        .drawer-panel { position: absolute; right: -480px; top: 0; bottom: 0; width: 480px; background: var(--bg-elevated); border-left: 1px solid var(--border-color); box-shadow: -10px 0 30px rgba(0,0,0,0.5); transition: right 0.3s cubic-bezier(0.2, 0.8, 0.2, 1); display: flex; flex-direction: column; }
        .drawer-overlay.show .drawer-panel { right: 0; }
        .drawer-head { padding: 20px 24px; border-bottom: 1px solid var(--border-color); display: flex; justify-content: space-between; align-items: center; }
        .drawer-title { font-size: 22px; font-weight: 600; color: var(--text-primary); }
        .btn-close { background: none; border: none; color: var(--text-secondary); font-size: 28px; cursor: pointer; padding: 4px; line-height: 1;}
        .btn-close:hover { color: var(--down-color); }
        .drawer-body { flex: 1; overflow-y: auto; padding: 0; }
        .txn-row { padding: 16px 24px; border-bottom: 1px solid var(--bg-highlight); display: flex; justify-content: space-between; align-items: center; transition: background 0.2s;}
        .txn-row:hover { background-color: var(--bg-highlight); }
        .txn-info { display: flex; flex-direction: column; gap: 4px; }
        .txn-date { color: var(--text-disabled); font-size: 21px; }
        .txn-merchant { color: var(--text-primary); font-weight: 500; font-size: 22px;}
        
        @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
    </style>
</head>
<body>
    <nav class="sidebar">
        <div class="logo-area">Analytics</div>
        <ul class="nav-menu">
            <li class="nav-item active" onclick="switchTab('overview', this)">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7"></rect><rect x="14" y="3" width="7" height="7"></rect><rect x="14" y="14" width="7" height="7"></rect><rect x="3" y="14" width="7" height="7"></rect></svg>
                总览面板
            </li>
            <li class="nav-item" onclick="switchTab('monthly', this)">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle></svg>
                月度数据与流水
            </li>
            <li class="nav-item" onclick="switchTab('large-exp', this)">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="1" x2="12" y2="23"></line><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"></path></svg>
                大额支出监控
            </li>
            <li class="nav-item" onclick="switchTab('suggestions', this)">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>
                策略建议
            </li>
        </ul>
        <div class="sidebar-footer">
            Build __BUILD_TIME__<br>Powered by ccb_bill_analyzer
        </div>
    </nav>
    
    <main class="main-wrapper">
        <header class="topbar">
            <div class="topbar-title" id="page-title">OVERVIEW 总览</div>
            <div class="topbar-info">
                <span>范围: __DATE_RANGE_START__ 至 __DATE_RANGE_END__</span>
                <span class="info-tag">脱敏环境</span>
            </div>
        </header>
        
        <div class="content-scroll">
            <section id="view-overview" class="view-section active">
                <div class="grid-row col-4">
                    <div class="card">
                        <div class="stat-label">累计总支出 (CNY)</div>
                        <div class="stat-value">__TOTAL_EXPENSE__</div>
                        <div class="stat-trend txt-down">周期内累计流出</div>
                    </div>
                    <div class="card">
                        <div class="stat-label">总交易笔数</div>
                        <div class="stat-value">__TOTAL_TRANSACTIONS__<span>笔</span></div>
                        <div class="stat-trend txt-up">有效消费记录</div>
                    </div>
                    <div class="card">
                        <div class="stat-label">月均开支水平 (CNY)</div>
                        <div class="stat-value">__AVG_MONTHLY__</div>
                        <div class="stat-trend">基于 __MONTH_COUNT__ 个自然月</div>
                    </div>
                    <div class="card">
                        <div class="stat-label">全期最大单笔 (CNY)</div>
                        <div class="stat-value">__MAX_AMOUNT__</div>
                        <div class="stat-trend" style="color:var(--text-secondary)">__MAX_MERCHANT__</div>
                    </div>
                </div>
                
                <div class="grid-row col-2">
                    <div class="card">
                        <div class="card-header"><div class="card-title">月度资金流出趋势</div></div>
                        <div id="chart-trend" class="chart-box-tall"></div>
                    </div>
                    <div class="card">
                        <div class="card-header"><div class="card-title">全期消费结构分布</div></div>
                        <div id="chart-pie" class="chart-box-tall"></div>
                    </div>
                </div>
                
                <div class="grid-row col-1">
                    <div class="card">
                        <div class="card-header"><div class="card-title">商户资金流向 TOP10</div></div>
                        <div id="chart-bar" class="chart-box-xl"></div>
                    </div>
                </div>
            </section>
            
            <section id="view-monthly" class="view-section">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
                    <div class="month-selector" id="month-tabs-container" style="margin-bottom:0;"></div>
                    <button class="btn-tab active" onclick="openDrawer('本月全量流水', t => true)">查看本月全量流水 &rarr;</button>
                </div>
                
                <div class="grid-row col-4">
                    <div class="card clickable" onclick="openDrawer('本月全量流水', t => true)" title="点击查看本月全量流水">
                        <div class="stat-label">当月支出 (点击查看详情)</div>
                        <div class="stat-value" id="m-total-val">0.00</div>
                    </div>
                    <div class="card">
                        <div class="stat-label">交易笔数</div>
                        <div class="stat-value" id="m-count-val">0</div>
                    </div>
                    <div class="card">
                        <div class="stat-label">日均消费</div>
                        <div class="stat-value" id="m-avg-val">0.00</div>
                    </div>
                    <div class="card">
                        <div class="stat-label">当月最大单笔</div>
                        <div class="stat-value" id="m-max-val">0.00</div>
                    </div>
                </div>
                
                <div class="grid-row col-2">
                    <div class="card">
                        <div class="card-header"><div class="card-title">活动热力图 (点击区块查看单日流水)</div></div>
                        <div class="cal-grid" id="calendar-grid"></div>
                        <div class="legend-bar">
                            <div class="legend-item"><div class="l-box" style="background:var(--bg-highlight)"></div>无</div>
                            <div class="legend-item"><div class="l-box" style="background:#203a27"></div>&lt;50</div>
                            <div class="legend-item"><div class="l-box" style="background:#2e5c3b"></div>50-200</div>
                            <div class="legend-item"><div class="l-box" style="background:#0ECB81"></div>200-500</div>
                            <div class="legend-item"><div class="l-box" style="background:#C99400"></div>500-1000</div>
                            <div class="legend-item"><div class="l-box" style="background:#F6465D"></div>&gt;1000</div>
                        </div>
                    </div>
                    <div class="card">
                        <div class="card-header"><div class="card-title">分类结构 (点击切片筛选流水)</div></div>
                        <div id="chart-m-pie" class="chart-box-tall"></div>
                    </div>
                </div>
                
                <div class="grid-row col-2">
                    <div class="card">
                        <div class="card-header"><div class="card-title">每日支出分布 (点击柱体筛选流水)</div></div>
                        <div id="chart-m-daily" class="chart-box-xl"></div>
                    </div>
                    <div class="card">
                        <div class="card-header"><div class="card-title">高频商户 TOP10 (点击柱体筛选流水)</div></div>
                        <div id="chart-m-top" class="chart-box-xl"></div>
                    </div>
                </div>
            </section>
            
            <section id="view-large-exp" class="view-section">
                <div class="card">
                    <div class="card-header">
                        <div class="card-title">大额异动记录监控</div>
                        <div class="info-tag">阈值: ≥ ¥__THRESHOLD__</div>
                    </div>
                    <div class="table-wrap">
                        <table>
                            <thead><tr><th>执行日期</th><th>资产/类目</th><th>对手方/商户</th><th style="text-align:right">成交金额(CNY)</th></tr></thead>
                            <tbody>__LARGE_EXPENSE_ROWS__</tbody>
                        </table>
                    </div>
                </div>
            </section>
            
            <section id="view-suggestions" class="view-section">
                <div class="card">
                    <div class="card-header"><div class="card-title">系统分析报告与策略建议</div></div>
                    <div style="max-width:800px">__SUGGESTIONS_HTML__</div>
                </div>
            </section>
            
        </div>
    </main>
    
    <div class="drawer-overlay" id="drawer-overlay" onclick="closeDrawer(event)">
        <div class="drawer-panel" onclick="event.stopPropagation()">
            <div class="drawer-head">
                <div class="drawer-title" id="drawer-title">流水明细 Transaction History</div>
                <button class="btn-close" onclick="closeDrawer()">×</button>
            </div>
            <div class="drawer-body" id="drawer-content">
                </div>
        </div>
    </div>

    <script>
    const monthlyData = """ + monthly_data_json + """;
    const months = """ + months_list + """;
    const weekdayNames = """ + weekday_names + """;
    
    let currentMonth = months[months.length - 1];
    const charts = {};
    
    // Theme Config for ECharts (Binance Dark Style)
    const echartsTheme = {
        textStyle: { color: '#848E9C' },
        title: { textStyle: { color: '#EAECEF' } },
        line: { smooth: true },
        tooltip: { backgroundColor: '#2B3139', borderColor: '#2B3139', textStyle: { color: '#EAECEF', fontSize: 12 }, padding: 8 },
        categoryAxis: { axisLine: { lineStyle: { color: '#2B3139' } }, splitLine: { show: false }, axisLabel: {fontSize: 20} },
        valueAxis: { axisLine: { show: false }, splitLine: { lineStyle: { color: '#2B3139', type: 'dashed' } }, axisLabel: {fontSize: 20} }
    };
    echarts.registerTheme('binance-dark', echartsTheme);
    
    function initOverviewCharts() {
        if(!charts.trend) charts.trend = echarts.init(document.getElementById('chart-trend'), 'binance-dark');
        charts.trend.setOption({
            tooltip: { trigger: 'axis' },
            grid: { left: '2%', right: '4%', bottom: '2%', top: '10%', containLabel: true },
            xAxis: { type: 'category', data: """ + months_list + """ },
            yAxis: { type: 'value' },
            series: [{ data: """ + monthly_totals + """, type: 'line', 
                itemStyle: { color: '#FCD535' }, areaStyle: { color: new echarts.graphic.LinearGradient(0,0,0,1, [{offset:0,color:'rgba(252,213,53,0.2)'},{offset:1,color:'rgba(252,213,53,0)'}]) }, symbolSize: 4, lineStyle: { width: 2 } }]
        });
        
        if(!charts.pie) charts.pie = echarts.init(document.getElementById('chart-pie'), 'binance-dark');
        charts.pie.setOption({
            tooltip: { trigger: 'item' },
            legend: { orient: 'vertical', right: 10, top: 'middle', textStyle: { color: '#EAECEF', fontSize: 12 }, icon: 'roundRect', itemWidth: 14, itemHeight: 14, itemGap: 10 },
            color: ['#0F4D92', '#B64342', '#E28E2C', '#8BCF8B', '#42949E', '#9A4D8E', '#3775BA'],
            series: [{ type: 'pie', radius: ['25%', '75%'], center: ['48%', '50%'],
                itemStyle: { borderColor: '#181A20', borderWidth: 2 }, label: { show: true, color: '#EAECEF', fontSize: 11, formatter: '{b}: {d}%' }, emphasis: { label: { show: true, fontSize: 13, fontWeight: 'bold' } }, data: """ + pie_data + """ }]
        });
        
        if(!charts.bar) charts.bar = echarts.init(document.getElementById('chart-bar'), 'binance-dark');
        charts.bar.setOption({
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
            grid: { left: 80, right: 60, top: 20, bottom: 30, containLabel: true },
            xAxis: { type: 'value', axisLabel: { fontSize: 20 } },
            yAxis: { type: 'category', data: """ + top_merchant_names + """, axisLabel: { color: '#EAECEF', fontSize: 20 } },
            series: [{ type: 'bar', data: """ + top_merchant_values + """, itemStyle: { color: '#2B3139' }, emphasis: { itemStyle: { color: '#FCD535' } },
                label: { show: true, position: 'right', color: '#EAECEF', fontSize: 18, formatter: '¥{c}' } }]
        });
    }
    
    function buildMonthTabs() {
        const c = document.getElementById('month-tabs-container');
        c.innerHTML = '';
        months.forEach(m => {
            const btn = document.createElement('button');
            btn.className = 'btn-tab' + (m === currentMonth ? ' active' : '');
            btn.textContent = m;
            btn.onclick = () => selectMonth(m);
            c.appendChild(btn);
        });
    }
    
    function getCalColor(amt) {
        if (!amt || amt === 0) return 'var(--bg-highlight)';
        if (amt < 50) return '#203a27';
        if (amt < 200) return '#2e5c3b';
        if (amt < 500) return '#0ECB81';
        if (amt < 1000) return '#C99400';
        return '#F6465D';
    }
    
    function renderMonthlyData() {
        const d = monthlyData[currentMonth];
        if(!d) return;
        
        document.getElementById('m-total-val').textContent = d.total.toLocaleString(undefined, {minimumFractionDigits: 2});
        document.getElementById('m-count-val').textContent = d.count;
        document.getElementById('m-avg-val').textContent = d.avg.toLocaleString(undefined, {minimumFractionDigits: 2});
        document.getElementById('m-max-val').textContent = d.max.toLocaleString(undefined, {minimumFractionDigits: 2});
        
        // Render Calendar
        const grid = document.getElementById('calendar-grid');
        grid.innerHTML = '';
        ['一','二','三','四','五','六','日'].forEach(day => {
            const h = document.createElement('div'); h.className = 'cal-header'; h.textContent = day; grid.appendChild(h);
        });
        
        let firstDay = new Date(d.year, d.month - 1, 1).getDay();
        firstDay = firstDay === 0 ? 6 : firstDay - 1;
        const daysInMonth = new Date(d.year, d.month, 0).getDate();
        
        for (let i = 0; i < firstDay; i++) {
            const empty = document.createElement('div'); empty.className = 'cal-day empty'; grid.appendChild(empty);
        }
        for (let day = 1; day <= daysInMonth; day++) {
            const dateStr = d.year + '-' + String(d.month).padStart(2,'0') + '-' + String(day).padStart(2,'0');
            const cell = document.createElement('div');
            cell.className = 'cal-day';
            const dayData = d.calendar[dateStr];
            if (dayData) {
                cell.style.background = getCalColor(dayData.total);
                const txtColor = dayData.total >= 50 ? '#0B0E11' : '#EAECEF';
                if(dayData.total>=50) cell.style.color = '#0B0E11';
                cell.innerHTML = '<div class="d-num" style="'+(dayData.total>=50?'color:#0B0E11':'')+'">'+day+'</div><div class="d-amt">'+dayData.total.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 0})+'</div>';
                cell.onclick = () => openDrawer(`日期流水: ${dateStr}`, t => t.date === dateStr);
            } else {
                cell.innerHTML = '<div class="d-num">'+day+'</div>';
            }
            grid.appendChild(cell);
        }
        
        // Month Pie Chart
        if(!charts.mPie) {
            charts.mPie = echarts.init(document.getElementById('chart-m-pie'), 'binance-dark');
            charts.mPie.on('click', function(params) { openDrawer(`分类明细: ${params.name}`, t => t.category === params.name); });
        }
        charts.mPie.setOption({
            tooltip: { trigger: 'item' },
            legend: { type: 'scroll', orient: 'vertical', right: 10, top: 'middle', itemWidth: 12, itemHeight: 12, icon: 'roundRect', textStyle:{color:'#EAECEF', fontSize: 12}, itemGap: 8 },
            color: ['#0F4D92', '#B64342', '#E28E2C', '#8BCF8B', '#42949E', '#9A4D8E', '#3775BA'],
            series: [{ type: 'pie', radius: ['25%', '72%'], center: ['48%', '50%'],
                      itemStyle: { borderColor: '#181A20', borderWidth: 1 },
                      label: { show: true, color: '#EAECEF', fontSize: 11, formatter: '{b}: {d}%' }, emphasis: { label: { show: true, fontSize: 13, fontWeight: 'bold' } }, data: d.cat_data }]
        });
        
        // Month Daily Chart
        if(!charts.mDaily) {
            charts.mDaily = echarts.init(document.getElementById('chart-m-daily'), 'binance-dark');
            charts.mDaily.on('click', function(params) { openDrawer(`日期流水: ${params.name}`, t => t.date === params.name); });
        }
        charts.mDaily.setOption({
            tooltip: { trigger: 'axis' }, grid: { left: '2%', right: '4%', bottom: '2%', top: '10%', containLabel: true },
            xAxis: { type: 'category', data: d.daily_dates }, yAxis: { type: 'value' },
            series: [{ type: 'bar', data: d.daily_values, itemStyle: { color: '#5888FF' }, emphasis: { itemStyle: { color: '#FCD535' } } }]
        });
        
        // Month Top Merchants Chart
        if(!charts.mTop) {
            charts.mTop = echarts.init(document.getElementById('chart-m-top'), 'binance-dark');
            charts.mTop.on('click', function(params) { openDrawer(`商户明细: ${params.name}`, t => t.merchant === params.name); });
        }
        charts.mTop.setOption({
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } }, grid: { left: 80, right: 60, top: 20, bottom: 30, containLabel: true },
            xAxis: { type: 'value', axisLabel: { fontSize: 20 } }, yAxis: { type: 'category', data: d.top_names, axisLabel: { color: '#EAECEF', fontSize: 20 } },
            series: [{ type: 'bar', data: d.top_values, itemStyle: { color: '#0ECB81' }, emphasis: { itemStyle: { color: '#FCD535' } },
                label: { show: true, position: 'right', color: '#EAECEF', fontSize: 18, formatter: '¥{c}' } }]
        });
    }
    
    function selectMonth(m) {
        currentMonth = m;
        buildMonthTabs();
        renderMonthlyData();
    }
    
    // Drawer Logic for Transactions
    function openDrawer(title, filterFn) {
        document.getElementById('drawer-title').textContent = title;
        const d = monthlyData[currentMonth];
        const txns = d.all_txns.filter(filterFn);
        let html = '';
        if (txns.length === 0) {
            html = '<div style="padding:24px; text-align:center; color:var(--text-disabled);">无相关流水</div>';
        } else {
            txns.forEach(t => {
                html += `
                <div class="txn-row">
                    <div class="txn-info">
                        <span class="txn-merchant">${t.merchant}</span>
                        <div style="display:flex; gap:8px; align-items:center;">
                            <span class="badge">${t.category}</span>
                            <span class="txn-date">${t.date}</span>
                        </div>
                    </div>
                    <div class="col-amt">¥${t.amount.toLocaleString(undefined, {minimumFractionDigits: 2})}</div>
                </div>`;
            });
        }
        document.getElementById('drawer-content').innerHTML = html;
        document.getElementById('drawer-overlay').classList.add('show');
    }
    
    function closeDrawer(e) { 
        if (!e || e.target.id === 'drawer-overlay') {
            document.getElementById('drawer-overlay').classList.remove('show'); 
        }
    }
    
    const titles = { 'overview': 'OVERVIEW 总览', 'monthly': 'MONTHLY 月度数据与流水', 'large-exp': 'MONITOR 大额支出监控', 'suggestions': 'SUGGESTIONS 策略建议' };
    function switchTab(tabId, el) {
        document.querySelectorAll('.view-section').forEach(s => s.classList.remove('active'));
        document.getElementById('view-' + tabId).classList.add('active');
        document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
        if(el) el.classList.add('active');
        document.getElementById('page-title').textContent = titles[tabId];
        
        if(tabId === 'overview') { initOverviewCharts(); charts.trend.resize(); charts.pie.resize(); charts.bar.resize(); }
        if(tabId === 'monthly') { buildMonthTabs(); renderMonthlyData(); charts.mPie.resize(); charts.mDaily.resize(); charts.mTop.resize(); }
    }
    
    window.onload = () => { switchTab('overview', document.querySelector('.nav-item.active')); };
    window.onresize = () => { Object.values(charts).forEach(c => { if(c) c.resize(); }); };
    </script>
</body>
</html>"""
    
    html_content = (html_content
        .replace('__BUILD_TIME__', now_str)
        .replace('__DATE_RANGE_START__', date_range_start)
        .replace('__DATE_RANGE_END__', date_range_end)
        .replace('__TOTAL_EXPENSE__', '{:,.2f}'.format(total_expense))
        .replace('__TOTAL_TRANSACTIONS__', str(total_transactions))
        .replace('__AVG_MONTHLY__', '{:,.2f}'.format(avg_monthly))
        .replace('__MONTH_COUNT__', str(len(months)))
        .replace('__MAX_AMOUNT__', '{:,.2f}'.format(max_amount))
        .replace('__MAX_MERCHANT__', max_merchant)
        .replace('__THRESHOLD__', '{:.0f}'.format(threshold))
        .replace('__LARGE_EXPENSE_ROWS__', large_expense_rows)
        .replace('__SUGGESTIONS_HTML__', suggestions_html)
    )
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    return output_path


def main():
    import sys
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = r'C:\Users\Admin\Desktop\downfile.csv'
        if not os.path.exists(input_file):
            input_file = r'C:\Users\Admin\Downloads\downfile.txt'
    
    if not os.path.exists(input_file):
        print("Error: File not found - {}".format(input_file))
        return
    
    output_dir = os.path.dirname(input_file) or os.getcwd()
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    print("[1/5] Parsing: {}".format(input_file))
    df = parse_bill(input_file)
    print("[2/5] Parsed: {} records".format(len(df)))
    df = apply_classification(df)
    print("[3/5] Classification done")
    
    excel_path = os.path.join(output_dir, '信用卡账单汇总_{}.xlsx'.format(timestamp))
    generate_excel(df, excel_path)
    print("[4/5] Excel: {}".format(excel_path))
    
    html_path = os.path.join(output_dir, '信用卡账单可视化_{}.html'.format(timestamp))
    generate_html(df, html_path)
    print("[5/5] HTML: {}".format(html_path))
    
    expense_df = df[df['金额'] > 0]
    print("\n=== Summary ===")
    print("   Total: {:,.2f}".format(expense_df['金额'].sum()))
    print("   Transactions: {}".format(len(expense_df)))
    print("   Period: {} ~ {}".format(df['交易日期'].min().strftime('%Y-%m-%d'), df['交易日期'].max().strftime('%Y-%m-%d')))
    print("\n=== Categories ===")
    for cat, amt in expense_df.groupby('分类')['金额'].sum().sort_values(ascending=False).items():
        pct = amt / expense_df['金额'].sum() * 100
        print("   {}: {:,.2f} ({:.1f}%)".format(cat, amt, pct))


if __name__ == '__main__':
    main()