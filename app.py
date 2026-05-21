import streamlit as st
import pandas as pd
from datetime import datetime
from transformer import transform, generate_excel, validate_input

st.set_page_config(page_title="I14 \u5831\u8868\u8a08\u7b97\u5de5\u5177", page_icon="\U0001f4ca", layout="wide")

st.title("\U0001f4ca I14 \u5831\u8868\u8a08\u7b97\u5de5\u5177")
st.markdown("\u4e0a\u50b3 **By SKUxShop** Excel \u6a94\u6848\uff0c\u81ea\u52d5\u7522\u751f **I14 by SKUxShop_To Buyer** \u7d50\u679c\u6a94\u3002")

st.divider()

uploaded_file = st.file_uploader(
    "\U0001f4c2 \u8acb\u4e0a\u50b3 Excel \u6a94\u6848 (.xlsx)",
    type=["xlsx"],
    help="\u652f\u63f4 .xlsx \u683c\u5f0f\uff0c\u6700\u5927 200MB\uff0c\u53ef\u8655\u7406\u9ad8\u9054 5 \u842c\u884c\u8cc7\u6599",
)

if uploaded_file is not None:
    with st.spinner("\u8b80\u53d6\u6a94\u6848\u4e2d..."):
        try:
            df_input = pd.read_excel(uploaded_file, engine="openpyxl")
        except Exception as e:
            st.error(f"\u274c \u8b80\u53d6\u6a94\u6848\u5931\u6557: {e}")
            st.stop()

    st.success(f"\u2705 \u6a94\u6848\u8b80\u53d6\u6210\u529f\uff01\u5171 **{len(df_input):,}** \u884c\uff0c**{len(df_input.columns)}** \u6b04")

    missing = validate_input(df_input)
    if missing:
        st.error(f"\u274c \u7f3a\u5c11\u5fc5\u8981\u6b04\u4f4d: {', '.join(missing)}")
        st.info("\u6a94\u6848\u5fc5\u9808\u5305\u542b\u4ee5\u4e0b\u6b04\u4f4d: Buyer, Brand, SKU, Shop Stock Status")
        st.stop()

    with st.expander("\U0001f50d \u8cc7\u6599\u9810\u89bd", expanded=False):
        st.dataframe(df_input.head(20), use_container_width=True)
        st.caption(f"\u986f\u793a\u524d 20 \u884c\uff08\u5171 {len(df_input):,} \u884c\uff09")

    st.divider()

    if st.button("\U0001f680 \u7522\u751f\u5831\u8868", type="primary", use_container_width=True):
        with st.spinner("\u8655\u7406\u4e2d..."):
            try:
                summary_df, raw_df, categories = transform(df_input)
                buf = generate_excel(summary_df, raw_df, categories)
            except Exception as e:
                st.error(f"\u274c \u8655\u7406\u5931\u6557: {e}")
                st.stop()

        st.success("\u2705 \u5831\u8868\u7522\u751f\u6210\u529f\uff01")

        tab1, tab2 = st.tabs(["\U0001f4ca \u7d71\u8a08\u532f\u7e3d", "\U0001f4c4 \u539f\u59cb\u8cc7\u6599"])

        with tab1:
            st.subheader("Shop Stock Status \u7d71\u8a08")
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
            total_overstock = summary_df[summary_df["Buyer"] == "\u7e3d\u8a08"]["\u7e3d\u8a08"].values[0]
            st.metric("\u7e3d Overstock \u7b46\u6578", f"{int(total_overstock):,}")

        with tab2:
            st.subheader("RAW \u539f\u59cb\u8cc7\u6599")
            st.dataframe(raw_df.head(50), use_container_width=True)
            st.caption(f"\u986f\u793a\u524d 50 \u884c\uff08\u5171 {len(raw_df):,} \u884c\uff09")

        today = datetime.now().strftime("%Y%m%d")
        download_name = f"I14 by SKUxShop_To Buyer_{today}.xlsx"

        st.download_button(
            label="\U0001f4e5 \u4e0b\u8f09\u5831\u8868 (.xlsx)",
            data=buf,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

st.divider()
st.caption("I14 \u5831\u8868\u8a08\u7b97\u5de5\u5177 \u00b7 Streamlit + Zeabur")
