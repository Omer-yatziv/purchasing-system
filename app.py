import streamlit as st
import pandas as pd
from thefuzz import fuzz, process
from io import BytesIO

st.set_page_config(page_title="מערכת רכש חכמה", layout="wide")
st.title("🎯 מערכת הצלבת רכש - גרסת טבלה נקייה")

@st.cache_data
def load_erp():
    for ext in ['xlsx', 'csv']:
        try:
            name = f"erp_master.{ext}"
            if ext == 'xlsx':
                df = pd.read_excel(name)
            else:
                df = pd.read_csv(name)
            df.columns = df.columns.str.strip()
            return df
        except: continue
    return None

erp_df = load_erp()

if erp_df is None:
    st.error("⚠️ קובץ ה-ERP (erp_master) חסר ב-GitHub")
else:
    st.success(f"בסיס נתונים ERP נטען בהצלחה")
    
    purchase_file = st.file_uploader("העלה בקשת רכש", type=['xlsx', 'xlsm'])
    
    if purchase_file:
        # המשתמש מגדיר מאיזו שורה מתחילה הטבלה באמת
        start_row = st.number_input("באיזו שורה מתחילה הטבלה (כותרות)?", value=10) - 1
        
        # טעינת הנתונים בלבד (מתעלם מכל מה שלמעלה)
        raw_df = pd.read_excel(purchase_file, header=start_row)
        raw_df.columns = raw_df.columns.astype(str).str.strip()
        
        st.write("תצוגה מקדימה של הנתונים שזוהו:")
        st.dataframe(raw_df.head(3))

        # בחירת עמודות לחיפוש (מוודא שמות)
        col_options = list(raw_df.columns)
        c_group = st.selectbox("בחר עמודת 'קבוצה'", col_options, index=col_options.index('קבוצה') if 'קבוצה' in col_options else 0)
        c_mat = st.selectbox("בחר עמודת 'ח\"ג'", col_options, index=col_options.index('ח"ג') if 'ח"ג' in col_options else 0)
        c_desc = st.selectbox("בחר עמודת 'תאור/מידה'", col_options, index=col_options.index('תאור/מידה') if 'תאור/מידה' in col_options else 0)

        if st.button("בצע הצלבה והפק קובץ חדש"):
            erp_choices = erp_df['תיאור פריט'].astype(str).tolist()
            final_results = []
            analysis_data = []

            for idx, row in raw_df.iterrows():
                # בניית שאילתה
                q_group = str(row.get(c_group, "")).replace("nan", "")
                q_mat = str(row.get(c_mat, "")).replace("nan", "")
                q_desc = str(row.get(c_desc, "")).replace("nan", "")
                
                if not q_desc.strip(): continue
                
                query = f"{q_group} {q_mat} {q_desc}".strip()
                
                # חיפוש חכם
                match = process.extractOne(query, erp_choices, scorer=fuzz.token_set_ratio)
                
                if match and match[1] > 40:
                    res_text, score = match[0], match[1]
                    erp_row = erp_df[erp_df['תיאור פריט'] == res_text].iloc[0]
                    res_code = erp_row['קוד פריט']
                    
                    # הצעות נוספות
                    others = process.extract(query, erp_choices, limit=3)
                    alt_text = ", ".join([f"{m[0]} ({m[1]}%)" for m in others[1:]])
                else:
                    res_code, res_text, score, alt_text = "לא נמצא", "", 0, ""

                # בניית שורה חדשה לאקסל: קוד ודירוג בהתחלה, ואז שאר נתוני המקור
                new_row = {
                    'קוד פריט ERP': res_code,
                    'תיאור ERP': res_text,
                    '% התאמה': f"{score}%"
                }
                new_row.update(row.to_dict())
                final_results.append(new_row)
                
                analysis_data.append({
                    'שורה': idx + start_row + 2,
                    'טקסט מקורי': query,
                    'זיהוי נבחר': res_text,
                    'ציון': score,
                    'חלופות': alt_text
                })

            # יצירת קובץ הפלט
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                pd.DataFrame(final_results).to_excel(writer, sheet_name='בקשת רכש מעודכנת', index=False)
                pd.DataFrame(analysis_data).to_excel(writer, sheet_name='ניתוח נתונים', index=False)
            
            st.success(f"הסתיים! עובדו {len(final_results)} שורות.")
            st.download_button("📥 הורד קובץ אקסל נקי", output.getvalue(), "Purchasing_Match_Results.xlsx")
