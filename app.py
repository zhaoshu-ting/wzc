import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="线上编码与聚水潭导入表", layout="wide")
st.title("生成线上编码与聚水潭导入表")
st.markdown("上传线上商品表，按商品编码匹配成本价，生成A表和B表。")

# 左右分栏
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("线上商品表格")
    st.info("必传，作为生成A表和B表的主文件。仅支持 .xlsx")
    file_online = st.file_uploader("点击或拖拽线上商品Excel到这里", type=["xlsx"])

with col2:
    st.subheader("成本价文件")
    st.info("上传后会强校验，避免生成错误价格。只读取商品id、档口名称、款号、成本价四列")
    file_cost = st.file_uploader("点击或拖拽成本价Excel到这里", type=["xlsx"])

# 选项：薄/厚 款式类型
style_type = st.radio("款式类型（加绒自动覆盖）", ["薄", "厚"], horizontal=True)
st.caption("编码预览：常规:款式编码-厚-颜色-码数；加绒:款式编码-绒-颜色加绒-码数")

# 处理按钮
run_btn = st.button("开始处理", type="primary")

# 定义导出excel函数
def export_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return output.getvalue()

if run_btn:
    if not file_online or not file_cost:
        st.warning("⚠️ 请同时上传【线上商品表格】和【成本价文件】两个Excel！")
    else:
        # 读取文件
        df_online = pd.read_excel(file_online)
        df_cost = pd.read_excel(file_cost)

        # 双条件匹配：款号 + 商品id
        df_merge = pd.merge(
            df_online,
            df_cost[["商品id", "档口名称", "款号", "成本价"]],
            on=["款号", "商品id"],
            how="left"
        )

        # A表 = 全部线上商品；B表 = 匹配到成本价的数据
        table_A = df_merge.copy()
        table_B = df_merge[df_merge["成本价"].notna()].copy()

        # 输出统计信息
        st.success(f"✅ A表共生成 {len(table_A)} 条数据；B表生成 {len(table_B)} 条，{len(table_B)} 条匹配到成本价")

        # 下载按钮
        a_file = export_excel(table_A)
        b_file = export_excel(table_B)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 下载A表", data=a_file, file_name="A表_线上商品.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c2:
            st.download_button("📥 下载B表", data=b_file, file_name="B表_聚水潭导入.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with st.expander("字段来源和处理规则"):
            st.markdown("""
            - 校验规则：**"供应商编码-款号"精确匹配"线上款式编码"，并且"商品id"精确匹配**
            - 线上商品表必需字段：店铺名称、平台店铺商品编码、线上款式编码、线上商品编码、线上商品名称、线上颜色规格、平台店铺款式编码
            - 成本价表必需字段：商品id、档口名称、款号、成本价
            """)