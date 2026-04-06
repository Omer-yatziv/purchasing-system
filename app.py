import streamlit as st
import pandas as pd
import re
from thefuzz import fuzz
from io import BytesIO

st.set_page_config(page_title="מערכת רכש חכמה", layout="wide")
st.title("🎯 מערכת התאמת רכש לסאפ - גרסה 2.1")

@st.cache_data
def load_erp():
    try:
        # טעינת ה-CSV שהעלית (לפי המידע שיש לי הוא בפורמט CSV כרגע)
        df = pd.read_csv("erp_master.csv") 
        return df
    except:
        try:
            df = pd.read_excel("erp_master.xlsx")
            return df
        except:
            return None

erp_df = load_erp()

if erp_df is None:
    st.error("⚠️ קובץ ה-ERP לא נמצא. וודא שהעלית erp_master.xlsx או erp_master.csv")
else:
    # ניקוי שמות עמודות ב-ERP (הסרת רווחים מיותרים)
    erp_df.columns = erp_df.columns.str.strip()
    st.success(f"בסיס הנתונים ERP נטען בהצלחה")

    purchase_file = st.file_uploader("העלה קובץ בקשת רכש", type=['xlsx', 'xlsm'])

    if purchase_file:
        start_row = st.number_input("השורה בה מתחילים הנתונים (למשל 10)", value=10) - 1
        
        # קריאת כל הקובץ
        df_all = pd.read_excel(purchase_file, header=None)
        
        # הגדרת הכותרות לפי השורה שבחרת
        header = df_all.iloc[start_row].fillna("Unnamed")
        data = df_all.iloc[start_row+1:].copy()
        data.columns = header
        data.columns = data.columns.str.strip() # ניקוי רווחים מהכותרות

        if st.button("הפעל חיפוש חכם"):
            results = []
            
            # זיהוי עמודות הרכש (לפי מה שכתבת לי)
            # אם השמות שונים אצלך באקסל, שנה אותם כאן בין הגרשיים
            col_group = 'קבוצה'
            col_material = 'ח"ג'
            col_desc = 'תאור/מידה'

            for idx, row in data.iterrows():
                # בניית השאילתה מהעמודות
                q_group = str(row.get(col_group, ""))
                q_mat = str(row.get(col_material, ""))
                q_desc = str(row.get(col_desc, ""))
                
                # אם השורה ריקה לגמרי, דלג
                if q_desc == "nan" or q_desc == "": continue
                
                full_query = f"{q_group} {q_mat} {q_desc}".replace("nan", "").strip()
                
                # חיפוש ב-ERP - אנחנו משווים לעמודה "תיאור פריט"
                # משתמשים ב-process.extractOne לדיוק מירבי
                from thefuzz import process
                
                # סינון ה-ERP רק לשורות שיש בהן תיאור
                choices = erp_df['תיאור פריט'].dropna().astype(str).tolist()
                best_match_tuple = process.extractOne(full_query, choices, scorer=fuzz.token_set_ratio)
                
                if best_match_tuple:
                    matched_text, score = best_match_tuple[0], best_match_tuple[1]
                    # מציאת השורה המלאה ב-ERP כדי לשלוף את הקוד
                    erp_row = erp_df[erp_df['תיאור פריט'] == matched_text].iloc[0]
                    
                    res_code = erp_row['קוד פריט']
                    res_desc = erp_row['תיאור פריט']
                else:
                    res_code, res_desc, score = "לא נמצא", "", 0

                # הוספת התוצאה
                new_row = {
                    'קוד פריט סאפ': res_code,
                    'תיאור סאפ': res_desc,
                    '% התאמה': f"{score}%"
                }
                new_row.update(row.to_dict())
                results.append(new_row)

            # יצירת קובץ להורדה
            final_df = pd.DataFrame(results)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name="בקשת רכש מעודכנת")
            
            st.success("הסתיים! הורד את הקובץ:")
            st.download_button("הורד אקסל", output.getvalue(), "result.xlsx")
