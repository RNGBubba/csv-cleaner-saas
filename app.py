import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime

st.set_page_config(page_title="CSV Cleaner by Hermes", layout="wide")

st.title("🧹 CSV Cleaner by Hermes")
st.markdown("Upload your messy CSV and get a clean, organized version in seconds.")

# Sidebar
st.sidebar.header("Cleaning Options")
remove_duplicates = st.sidebar.checkbox("Remove duplicate rows", value=True)
standardize_dates = st.sidebar.checkbox("Standardize date formats", value=True)
fix_phone_numbers = st.sidebar.checkbox("Fix phone number formats", value=True)
remove_empty_rows = st.sidebar.checkbox("Remove empty rows", value=True)
strip_whitespace = st.sidebar.checkbox("Trim whitespace from all cells", value=True)
add_row_numbers = st.sidebar.checkbox("Add row numbers", value=False)

# File upload
uploaded_file = st.file_uploader("Upload your CSV file", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)
        else:
            df = pd.read_csv(uploaded_file)
        
        st.success(f"✅ Loaded {len(df)} rows and {len(df.columns)} columns")
        
        st.subheader("Original Data (Preview)")
        st.dataframe(df.head(10), use_container_width=True)
        
        issues = []
        
        if remove_duplicates and df.duplicated().any():
            issues.append(f"Found {df.duplicated().sum()} duplicate rows")
        
        if remove_empty_rows and df.isnull().all(axis=1).any():
            issues.append(f"Found {df.isnull().all(axis=1).sum()} completely empty rows")
        
        if strip_whitespace:
            whitespace_cols = []
            for col in df.select_dtypes(include='object').columns:
                if df[col].astype(str).str.match(r'^\s|\s$').any():
                    whitespace_cols.append(col)
            if whitespace_cols:
                issues.append(f"Found whitespace issues in columns: {', '.join(whitespace_cols[:5])}")
        
        if issues:
            st.warning("**Issues detected:**")
            for issue in issues:
                st.markdown(f"- {issue}")
        else:
            st.info("No major issues detected!")
        
        if st.button("🧹 Clean Data", type="primary", use_container_width=True):
            cleaned = df.copy()
            report = []
            
            if remove_duplicates:
                before = len(cleaned)
                cleaned = cleaned.drop_duplicates()
                removed = before - len(cleaned)
                if removed > 0:
                    report.append(f"Removed {removed} duplicate rows")
            
            if remove_empty_rows:
                before = len(cleaned)
                cleaned = cleaned.dropna(how='all')
                removed = before - len(cleaned)
                if removed > 0:
                    report.append(f"Removed {removed} empty rows")
            
            if standardize_dates:
                for col in cleaned.columns:
                    if cleaned[col].dtype == 'object':
                        try:
                            parsed = pd.to_datetime(cleaned[col], infer_datetime_format=True, errors='coerce')
                            if parsed.notna().sum() > len(cleaned) * 0.5:
                                cleaned[col] = parsed.dt.strftime('%Y-%m-%d')
                                report.append(f"Standardized dates in column '{col}'")
                        except:
                            pass
            
            if fix_phone_numbers:
                for col in cleaned.columns:
                    if 'phone' in col.lower() or 'tel' in col.lower() or 'mobile' in col.lower():
                        cleaned[col] = cleaned[col].astype(str).apply(
                            lambda x: re.sub(r'[^0-9]', '', x) if x != 'nan' else x
                        )
                        cleaned[col] = cleaned[col].apply(
                            lambda x: f"{x[:3]}-{x[3:6]}-{x[6:10]}" if len(x) == 10 else x
                        )
                        report.append(f"Fixed phone format in column '{col}'")
            
            if strip_whitespace:
                for col in cleaned.select_dtypes(include='object').columns:
                    cleaned[col] = cleaned[col].astype(str).str.strip()
                report.append("Trimmed whitespace from all text cells")
            
            if add_row_numbers:
                cleaned.insert(0, 'Row_ID', range(1, len(cleaned) + 1))
                report.append("Added Row_ID column")
            
            cleaned = cleaned.reset_index(drop=True)
            
            st.subheader("🧾 Cleaning Report")
            if report:
                for item in report:
                    st.markdown(f"✅ {item}")
            else:
                st.markdown("No changes needed — your data was already clean!")
            
            st.success(f"**Cleaned data: {len(cleaned)} rows × {len(cleaned.columns)} columns**")
            
            st.subheader("Cleaned Data (Preview)")
            st.dataframe(cleaned.head(10), use_container_width=True)
            
            csv = cleaned.to_csv(index=False)
            st.download_button(
                label="📥 Download Cleaned CSV",
                data=csv,
                file_name=f"cleaned_{uploaded_file.name}",
                mime="text/csv",
                type="primary",
                use_container_width=True
            )
            
            st.divider()
            st.subheader("💰 Need more than automated cleaning?")
            st.markdown("For custom transformations, API integration, or complex data work:")
            st.markdown("- **$75** — Single file, custom rules, 24hr delivery")
            st.markdown("- **$150** — Multiple files + automation script")
            st.markdown("- **$250** — Full data pipeline + documentation")
            st.markdown("[Pay via PayPal](https://www.paypal.com/checkoutnow?token=4GE57494VV040024T)")
            st.markdown("Or email: mrbubba@agentmail.to")
    
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.info("Please make sure your file is a valid CSV or Excel file.")

st.divider()
st.markdown("Built by **Hermes Data Services** | Contact: mrbubba@agentmail.to")
