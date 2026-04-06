import streamlit as st
import pandas as pd
from thefuzz import fuzz
from io import BytesIO

# עיצוב דף האפליקציה
st.set_page_config(page_title="מערכת התאמת רכש ל-ERP", layout="wide")
st.title("📦 מערכת הצלבת פריטי רכש חכמה")
st.markdown("העלה את הקבצים והמערכת תייצר עבורך קובץ מעובד עם קודי ERP.")

# --- שלב 1: העלאת קבצים ---
col1, col2 = st.columns(2)
with col1:
    erp_file = st.file_uploader("העלה קובץ ERP (בסיס נתונים)", type=['xlsx', 'csv'])
with col2:
    purchase_file = st.file_uploader("העלה קובץ בקשת רכש", type=['xlsx', 'csv', 'xlsm'])

if erp_file and purchase_file:
    # טעינת נתונים
    erp_df = pd.read_excel(erp_file)
    purchase_df = pd.read_excel(purchase_file)
    
    st.success("הקבצים נטענו בהצלחה!")
    
    # בחירת עמודות (דינמי)
    st.sidebar.header("הגדרות עמודות")
    erp_desc_col = st.sidebar.selectbox("עמודת תיאור ב-ERP", erp_df.columns)
    erp_code_col = st.sidebar.selectbox("עמודת קוד פריט ב-ERP", erp_df.columns)
    pr_desc_col = st.sidebar.selectbox("עמודת תיאור בבקשת הרכש", purchase_df.columns)

    if st.button("🚀 התחל התאמה וייצור קובץ"):
        results = []
        analysis_log = []
        
        progress_bar = st.progress(0)
        total_rows = len(purchase_df)

        for idx, row in purchase_df.iterrows():
            query = str(row[pr_desc_col])
            
            # חישוב דמיון
            matches = erp_df.apply(lambda x: fuzz.token_set_ratio(query, str(x[erp_desc_col])), axis=1)
            best_idx = matches.idxmax()
            score = matches.max()
            
            # שליפת נתוני ERP
            res_code = erp_df.loc[best_idx, erp_code_col]
            res_desc = erp_df.loc[best_idx, erp_desc_col]
            
            # בניית שורה ללשונית א'
            new_row = {
                'קוד פריט נבחר': res_code,
                '% התאמה': f"{score}%",
                'תיאור ERP מקורי': res_desc
            }
            new_row.update(row.to_dict())
            results.append(new_row)
            
            # בניית שורה ללשונית ב'
            status = "✅ גבוהה" if score > 80 else "⚠️ דורש בדיקה" if score > 50 else "❌ נמוכה"
            analysis_log.append({
                'שורה': idx + 1,
                'טקסט מקורי': query,
                'התאמה לתיאור': res_desc,
                'סטטוס': status,
                'ציון': score
            })
            
            progress_bar.progress((idx + 1) / total_rows)

        # יצירת התוצאה הסופית
        final_pr = pd.DataFrame(results)
        final_analysis = pd.DataFrame(analysis_log)

        # יצירת קובץ אקסל בזיכרון
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_pr.to_excel(writer, sheet_name='בקשת רכש מעודכנת', index=False)
            final_analysis.to_excel(writer, sheet_name='ניתוח נתונים', index=False)
        
        st.balloons()
        st.download_button(
            label="📥 הורד קובץ מעובד ל-Monday",
            data=output.getvalue(),
            file_name="Processed_Purchase_Request.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
