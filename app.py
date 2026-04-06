import streamlit as st
import pandas as pd
from thefuzz import fuzz, process
from io import BytesIO

st.set_page_config(page_title="מערכת רכש חכמה", layout="wide")
st.title("🎯 מערכת הצלבת רכש - שלב אבחון")

# פונקציה לטעינת ה-ERP (תומכת ב-CSV ו-XLSX)
@st.cache_data
def load_erp():
    for ext in ['xlsx', 'csv']:
        try:
            name = f"erp_master.{ext}"
            df = pd.read_excel(name) if ext == 'xlsx' else pd.read_csv(name)
            df.columns = df.columns.str.strip()
            return df
        except: continue
    return None

erp_df = load_erp()

if erp_df is None:
    st.error("⚠️ לא נמצא קובץ erp_master.xlsx או erp_master.csv ב-GitHub")
else:
    st.success(f"בסיס נתונים ERP נטען בהצלחה ({len(erp_df)} שורות)")
    
    purchase_file = st.file_uploader("העלה בקשת רכש", type=['xlsx', 'xlsm'])
    
    if purchase_file:
        # קריאת הקובץ והצגת העמודות למשתמש כדי למנוע טעויות
        skip = st.number_input("באיזו שורה נמצאות הכותרות? (למשל 10)", value=10) - 1
        full_df = pd.read_excel(purchase_file, header=skip)
        full_df.columns = full_df.columns.astype(str).str.strip()
        
        st.write("---")
        st.subheader("בדיקת זיהוי עמודות:")
        col1, col2, col3 = st.columns(3)
        
        # בחירה ידנית של העמודות למקרה שהשם השתנה מעט
        with col1: sel_group = st.selectbox("בחר עמודת 'קבוצה'", full_df.columns, index=0 if 'קבוצה' in full_df.columns else 0)
        with col2: sel_mat = st.selectbox("בחר עמודת 'ח\"ג'", full_df.columns, index=0 if 'ח"ג' in full_df.columns else 0)
        with col3: sel_desc = st.selectbox("בחר עמודת 'תאור/מידה'", full_df.columns, index=0 if 'תאור/מידה' in full_df.columns else 0)

        if st.button("הפעל עיבוד"):
            # סינון שורות ריקות בתיאור
            data_to_process = full_df.dropna(subset=[sel_desc]).copy()
            
            if data_to_process.empty:
                st.warning("לא נמצאו נתונים לעיבוד בשורות שנבחרו.")
            else:
                results = []
                erp_choices = erp_df['תיאור פריט'].astype(str).tolist()
                
                for idx, row in data_to_process.iterrows():
                    # בניית השאילתה
                    q = f"{row[sel_group]} {row[sel_mat]} {row[sel_desc]}".replace("nan", "").strip()
                    
                    # חיפוש חכם
                    match_res = process.extractOne(q, erp_choices, scorer=fuzz.token_set_ratio)
                    
                    if match_res and match_res[1] > 40:
                        matched_text, score = match_res[0], match_res[1]
                        erp_row = erp_df[erp_df['תיאור פריט'] == matched_text].iloc[0]
                        res_code = erp_row['קוד פריט']
                        res_desc = erp_row['תיאור פריט']
                    else:
                        res_code, res_desc, score = "לא נמצא", "", 0
                    
                    # בניית השורה החדשה (הוספת הקוד בהתחלה)
                    res_row = {'קוד פריט סאפ': res_code, 'תיאור סאפ': res_desc, '% התאמה': f"{score}%"}
                    res_row.update(row.to_dict())
                    results.append(res_row)

                # יצירת קובץ הפלט
                final_df = pd.DataFrame(results)
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, sheet_name='בקשת רכש מעודכנת', index=False)
                    # יצירת לשונית ניתוח
                    pd.DataFrame(results)[['קוד פריט סאפ', '% התאמה']].to_excel(writer, sheet_name='ניתוח נתונים', index=False)
                
                st.success(f"עיבוד של {len(results)} שורות הסתיים!")
                st.download_button("📥 הורד קובץ מעובד", output.getvalue(), "Purchase_Analysis.xlsx")
