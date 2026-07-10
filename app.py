# app.py
import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(
    page_title="智谱AI账单核算系统",
    page_icon="📊",
    layout="wide"
)

# ---------- 复用原有分析逻辑 ----------
def find_column(df, possible_names):
    for name in possible_names:
        if name in df.columns:
            return name
    return None

def analyze_discount(df):
    try:
        product_name_col = find_column(df, ['模型产品名称', '产品名称', '模型名称'])
        catalog_price_col = find_column(df, ['目录价', '原价', '标准价'])
        unit_price_col = find_column(df, ['单价', '结算单价'])
        discount_ratio_col = find_column(df, ['折扣比', '折扣率', 'discount'])
        discount_type_col = find_column(df, ['折扣类型', '优惠类型'])

        if not all([product_name_col, catalog_price_col, unit_price_col]):
            return {'has_discount_data': False, 'message': '缺少目录价或单价数据'}

        analysis_df = df[[product_name_col, catalog_price_col, unit_price_col]].copy()
        if discount_ratio_col:
            analysis_df['折扣比'] = df[discount_ratio_col]
        if discount_type_col:
            analysis_df['折扣类型'] = df[discount_type_col]

        analysis_df['有折扣'] = False
        analysis_df['折扣说明'] = '无折扣'

        mask_price_diff = analysis_df[catalog_price_col] != analysis_df[unit_price_col]
        analysis_df.loc[mask_price_diff, '有折扣'] = True
        analysis_df.loc[mask_price_diff, '折扣说明'] = '目录价与单价不一致'

        if discount_ratio_col:
            mask_discount = analysis_df['折扣比'] != 1.0
            analysis_df.loc[mask_discount, '有折扣'] = True
            analysis_df.loc[mask_discount, '折扣说明'] = '有折扣比'

        product_summary = analysis_df.groupby(product_name_col).agg({
            '有折扣': lambda x: any(x),
            catalog_price_col: 'first',
            unit_price_col: 'first'
        }).reset_index()

        if discount_ratio_col:
            product_summary['折扣比'] = analysis_df.groupby(product_name_col)['折扣比'].first().values
        if discount_type_col:
            product_summary['折扣类型'] = analysis_df.groupby(product_name_col)['折扣类型'].first().values

        products_with_discount = product_summary[product_summary['有折扣'] == True]
        products_without_discount = product_summary[product_summary['有折扣'] == False]

        result = {
            'has_discount_data': True,
            'total_products': len(product_summary),
            'products_with_discount': len(products_with_discount),
            'products_without_discount': len(products_without_discount),
            'product_list': product_summary.to_dict('records'),
            'discount_details': []
        }

        for _, row in products_with_discount.iterrows():
            detail = {
                '产品名称': row[product_name_col],
                '目录价': row[catalog_price_col],
                '单价': row[unit_price_col],
                '折扣比': row.get('折扣比', 'N/A'),
                '折扣类型': row.get('折扣类型', 'N/A')
            }
            if row[catalog_price_col] and row[catalog_price_col] != 0:
                discount_rate = (row[catalog_price_col] - row[unit_price_col]) / row[catalog_price_col] * 100
                detail['折扣率'] = f"{discount_rate:.2f}%"
            else:
                detail['折扣率'] = 'N/A'
            result['discount_details'].append(detail)

        return result
    except Exception as e:
        return {'has_discount_data': False, 'message': f'折扣分析出错: {str(e)}'}

def analyze_single_file(file_content, file_name):
    try:
        df = pd.read_excel(io.BytesIO(file_content), sheet_name='Cost Details', header=0)
    except:
        try:
            xl = pd.ExcelFile(io.BytesIO(file_content))
            sheet_names = xl.sheet_names
            df = pd.read_excel(io.BytesIO(file_content), sheet_name=sheet_names[0], header=0)
        except Exception as e:
            return {'file_name': file_name, 'error': str(e), 'success': False}

    try:
        consumption_col = find_column(df, ['总消费金额（结算金额加总）', '总消费金额', '结算金额'])
        unpaid_col = find_column(df, ['未付款金额', '未付金额'])
        invoicable_col = find_column(df, ['可开票金额', '可开票'])
        paid_col = find_column(df, ['已付款金额', '已付款'])

        total_consumption = df[consumption_col].sum() if consumption_col else 0
        total_unpaid = df[unpaid_col].sum() if unpaid_col else 0
        total_invoicable = df[invoicable_col].sum() if invoicable_col else 0
        total_paid = df[paid_col].sum() if paid_col else (total_consumption - total_unpaid if total_unpaid is not None else 0)

        discount_analysis = analyze_discount(df)
        product_name_col = find_column(df, ['模型产品名称', '产品名称', '模型名称'])
        models = df[product_name_col].unique().tolist() if product_name_col else []

        return {
            'file_name': file_name,
            'total_consumption': total_consumption,
            'total_unpaid': total_unpaid,
            'total_paid': total_paid,
            'total_invoicable': total_invoicable,
            'models': models,
            'discount_analysis': discount_analysis,
            'success': True
        }
    except Exception as e:
        return {'file_name': file_name, 'error': str(e), 'success': False}

def calculate_bill_summary(results):
    records = []
    all_discount_products = {}
    all_normal_products = set()

    for r in results:
        if not r['success']:
            continue

        record = {
            '文件名': r['file_name'],
            '总消费金额': r['total_consumption'],
            '未付款金额': r['total_unpaid'],
            '可开票金额': r['total_invoicable'],
        }

        models = r.get('models', [])
        discount_analysis = r.get('discount_analysis', {})
        
        if discount_analysis and discount_analysis.get('has_discount_data'):
            for detail in discount_analysis.get('discount_details', []):
                product_name = detail['产品名称']
                if product_name not in all_discount_products:
                    all_discount_products[product_name] = detail
            
            for product in discount_analysis.get('product_list', []):
                if not product.get('有折扣', False):
                    product_name = product.get('产品名称', '')
                    if product_name and product_name not in ['未知', 'N/A']:
                        all_normal_products.add(product_name)
        
        records.append(record)

    if not records:
        return pd.DataFrame(), {}, set()

    df_summary = pd.DataFrame(records)
    df_summary = df_summary.sort_values('文件名').reset_index(drop=True)
    return df_summary, all_discount_products, all_normal_products

def generate_discount_config(all_discount_products, all_normal_products):
    rows = []
    
    for product_name, detail in sorted(all_discount_products.items()):
        discount_ratio = detail.get('折扣比', 'N/A')
        if discount_ratio == 'N/A' and '折扣率' in detail:
            rate_str = detail.get('折扣率', '')
            if '%' in rate_str:
                try:
                    discount_ratio = 1 - float(rate_str.replace('%', '')) / 100
                    discount_ratio = round(discount_ratio, 2)
                except:
                    discount_ratio = 'N/A'
        
        rows.append({
            '模型产品名称': product_name,
            '账单核算折扣': discount_ratio,
            '是否配置': '是',
            '备注': f"目录价: {detail.get('目录价', 'N/A')}, 单价: {detail.get('单价', 'N/A')}"
        })
    
    for product_name in sorted(all_normal_products):
        if product_name not in [p['模型产品名称'] for p in rows]:
            rows.append({
                '模型产品名称': product_name,
                '账单核算折扣': 1.0,
                '是否配置': '否',
                '备注': '无折扣'
            })
    
    return pd.DataFrame(rows)

# ---------- Streamlit UI ----------
st.title("📊 智谱AI开放平台 · 账单核算系统")
st.markdown("上传账单明细Excel文件，自动生成核算汇总表")

with st.sidebar:
    st.header("⚙️ 使用说明")
    st.markdown("""
    1. 点击下方上传按钮，选择多个Excel账单文件
    2. 系统自动分析每个文件的费用明细
    3. 生成核算汇总表格
    4. 支持下载CSV格式结果
    
    **支持识别的列名：**
    - 总消费金额 / 结算金额
    - 未付款金额 / 未付金额
    - 可开票金额 / 可开票
    - 模型产品名称 / 产品名称
    - 目录价 / 原价
    - 单价 / 结算单价
    """)
    st.divider()
    st.caption("v1.0 | 支持 .xlsx 格式")

uploaded_files = st.file_uploader(
    "📤 上传账单Excel文件",
    type=['xlsx', 'xls'],
    accept_multiple_files=True,
    help="支持同时上传多个文件，系统会自动合并分析"
)

if uploaded_files:
    st.info(f"已上传 {len(uploaded_files)} 个文件")
    
    if st.button("🚀 开始核算", type="primary"):
        with st.spinner("正在分析账单文件，请稍候..."):
            results = []
            progress_bar = st.progress(0)
            for i, uploaded_file in enumerate(uploaded_files):
                file_content = uploaded_file.read()
                result = analyze_single_file(file_content, uploaded_file.name)
                results.append(result)
                progress_bar.progress((i + 1) / len(uploaded_files))
            progress_bar.empty()
            
            df_summary, discount_products, normal_products = calculate_bill_summary(results)
            
            st.success("✅ 核算完成！")
            
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.subheader("📋 费用汇总")
                
                # 检查 df_summary 是否为空
                if df_summary.empty:
                    st.warning("⚠️ 没有成功解析任何文件，请检查上传的Excel格式是否正确")
                else:
                    st.dataframe(
                        df_summary,
                        width='stretch',  # 替换 use_container_width
                        column_config={
                            '总消费金额': st.column_config.NumberColumn(format="¥%.4f"),
                            '未付款金额': st.column_config.NumberColumn(format="¥%.4f"),
                            '可开票金额': st.column_config.NumberColumn(format="¥%.4f"),
                        }
                    )
                    
                    # ========== 汇总统计 ==========
                    total_consumption = df_summary['总消费金额'].sum()
                    total_unpaid = df_summary['未付款金额'].sum()
                    total_paid = total_consumption - total_unpaid

                    st.subheader("📊 汇总统计")
                    metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)

                    with metric_col1:
                        st.markdown(f"""
                        <div style="text-align: center; padding: 8px; background: #f0f2f6; border-radius: 8px;">
                            <div style="font-size: 13px; color: #666;">总消费金额</div>
                            <div style="font-size: 20px; font-weight: bold; color: #1f77b4;">¥{total_consumption:,.2f}</div>
                        </div>
                        """, unsafe_allow_html=True)

                    with metric_col2:
                        st.markdown(f"""
                        <div style="text-align: center; padding: 8px; background: #f0f2f6; border-radius: 8px;">
                            <div style="font-size: 13px; color: #666;">已付款金额</div>
                            <div style="font-size: 20px; font-weight: bold; color: #2ca02c;">¥{total_paid:,.2f}</div>
                        </div>
                        """, unsafe_allow_html=True)

                    with metric_col3:
                        st.markdown(f"""
                        <div style="text-align: center; padding: 8px; background: #f0f2f6; border-radius: 8px;">
                            <div style="font-size: 13px; color: #666;">未付款金额</div>
                            <div style="font-size: 20px; font-weight: bold; color: #d62728;">¥{total_unpaid:,.2f}</div>
                        </div>
                        """, unsafe_allow_html=True)

                    with metric_col4:
                        if total_consumption > 0:
                            st.markdown(f"""
                            <div style="text-align: center; padding: 8px; background: #f0f2f6; border-radius: 8px;">
                                <div style="font-size: 13px; color: #666;">付款率</div>
                                <div style="font-size: 20px; font-weight: bold; color: #9467bd;">{(total_paid/total_consumption*100):.1f}%</div>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.markdown("""
                            <div style="text-align: center; padding: 8px; background: #f0f2f6; border-radius: 8px;">
                                <div style="font-size: 13px; color: #666;">付款率</div>
                                <div style="font-size: 20px; font-weight: bold; color: #9467bd;">N/A</div>
                            </div>
                            """, unsafe_allow_html=True)
            
            with col2:
                st.subheader("🏷️ 折扣配置")
                if discount_products or normal_products:
                    df_discount = generate_discount_config(discount_products, normal_products)
                    st.dataframe(
                        df_discount,
                        width='stretch',  # 替换 use_container_width
                        column_config={
                            '账单核算折扣': st.column_config.NumberColumn(format="%.2f"),
                        }
                    )
                    st.caption(f"共 {len(df_discount)} 个模型配置")
                else:
                    st.info("未检测到折扣数据")
            
            st.divider()
            download_col1, download_col2 = st.columns(2)
            
            with download_col1:
                if not df_summary.empty:
                    csv_summary = df_summary.to_csv(index=False, encoding='utf-8-sig')
                    st.download_button(
                        label="📥 下载费用汇总 (CSV)",
                        data=csv_summary,
                        file_name="账单费用汇总.csv",
                        mime="text/csv"
                    )
            
            with download_col2:
                if (discount_products or normal_products) and 'df_discount' in locals():
                    csv_discount = df_discount.to_csv(index=False, encoding='utf-8-sig')
                    st.download_button(
                        label="📥 下载折扣配置 (CSV)",
                        data=csv_discount,
                        file_name="折扣配置.csv",
                        mime="text/csv"
                    )
            
            failed_files = [r for r in results if not r['success']]
            if failed_files:
                st.warning(f"⚠️ {len(failed_files)} 个文件解析失败")
                for f in failed_files:
                    st.text(f"  • {f['file_name']}: {f.get('error', '未知错误')}")

else:
    st.info("👆 请上传Excel账单文件开始核算")
    with st.expander("📖 查看支持的Excel格式示例"):
        st.markdown("""
        **需要包含以下列（至少包含其中一列）：**
        
        | 列名 | 说明 |
        |------|------|
        | 总消费金额 / 结算金额 | 必填 |
        | 未付款金额 / 未付金额 | 可选 |
        | 可开票金额 / 可开票 | 可选 |
        | 模型产品名称 / 产品名称 | 用于折扣分析 |
        | 目录价 / 原价 | 用于折扣分析 |
        | 单价 / 结算单价 | 用于折扣分析 |
        """)

st.divider()
st.caption("💡 提示：系统仅处理上传文件，不会存储任何数据")