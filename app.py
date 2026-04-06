import streamlit as st
import pandas as pd
from thefuzz import fuzz, process
from io import BytesIO
import openpyxl

st.set_page_config(page_title="מערכת רכש חכמה", layout="wide")
st.title("🎯 מערכת הצלבת רכש - שמירת מבנה ממוזג")

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
    st.error("⚠️ קובץ ה-ERP חסר ב-GitHub")
else:
    purchase_file = st.file_uploader("העלה בקשת רכש מקורית", type=['xlsx', 'xlsm'])
    
    if purchase_file:
        # הגדרות משתמש
        start_row = st.number_input("באיזו שורה מתחילה הטבלה (כותרות)?", value=10)
        
        # טעינת הקובץ בצורה שתשמור על הכל
        wb = openpyxl.load_workbook(purchase_file, keep_vba=True)
        ws = wb.active

        if st.button("הפעל עיבוד חכם"):
            erp_choices = erp_df['תיאור פריט'].astype(str).tolist()
            results_analysis = []
            
            # אנחנו לא מזיזים עמודות (insert_cols) כדי לא להרוס מיזוגים.
            # במקום זה, נניח שעמודות A ו-B ו-C הן המקום לנתוני ה-ERP.
            # אם הן תפוסות, נכתוב בעמודות פנויות בסוף או שנגדיר מראש.
            
            st.info("מעבד נתונים...")
            
            # לוגיקה למציאת העמודות לפי כותרות בשורה שצוינה
            header_map = {}
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row=start_row, column=col).value
                if val:
                    header_map[str(val).strip()] = col

            # וידוא עמודות נדרשות
            needed = ['קבוצה', 'ח"ג', 'תאור/מידה']
            found_all = all(n in header_map for n in needed)

            if not found_all:
                st.error(f"לא נמצאו כל העמודות הנדרשות. נמצאו: {list(header_map.keys())}")
            else:
                # כתיבת כותרות לעמודות ה-ERP (נבחר ב-A, B, C אם הן פנויות או פשוט נכתוב בהן)
                # אם אתה רוצה עמודות ספציפיות, נשנה את ה-1,2,3 למספרים אחרים
                ws.cell(row=start_row, column=1).value = "קוד פריט סאפ"
                ws.cell(row=start_row, column=2).value = "תיאור סאפ"
                ws.cell(row=start_row+1, column=1).value = "" # ניקוי שורת משנה אם יש

                for r in range(start_row + 1, ws.max_row + 1):
                    # שליפת הנתונים לפי המיקום שנמצא ב-header_map
                    q_group = str(ws.cell(row=r, column=header_map['קבוצה']).value or "")
                    q_mat = str(ws.cell(row=r, column=header_map['ח"ג']).value or "")
                    q_desc = str(ws.cell(row=r, column=header_map['תאור/מידה']).value or "")
                    
                    if q_desc == "" or q_desc == "None": continue
                    
                    query = f"{q_group} {q_mat} {q_desc}".strip()
                    
                    # הצלבה
                    match = process.extractOne(query, erp_choices, scorer=fuzz.token_set_ratio)
                    
                    if match and match[1] > 45:
                        best_text, score = match[0], match[1]
                        erp_row = erp_df[erp_df['תיאור פריט'] == best_text].iloc[0]
                        
                        # כתיבה ישירה לתאים A ו-B בשורה הנוכחית
                        ws.cell(row=r, column=1).value = erp_row['קוד פריט']
                        ws.cell(row=r, column=2).value = erp_row['תיאור פריט']
                        
                        # הצעות חלופיות ללשונית השנייה
                        all_matches = process.extract(query, erp_choices, limit=3)
                        alt_text = ", ".join([f"{m[0]} ({m[1]}%)" for m in all_matches[1:]])
                    else:
                        ws.cell(row=r, column=1).value = "לא נמצא"
                        score = 0
                        alt_text = ""

                    results_analysis.append({
                        'שורה': r,
                        'תיאור מקורי': query,
                        'זיהוי': ws.cell(row=r, column=2).value,
                        'ציון': score,
                        'חלופות': alt_text
                    })

                # לשונית ניתוח נתונים
                if "ניתוח נתונים" in wb.sheetnames: del wb["ניתוח נתונים"]
                ws_an = wb.create_sheet("ניתוח נתונים")
                an_df = pd.DataFrame(results_analysis)
                from openpyxl.utils.dataframe import dataframe_to_rows
                for row_data in dataframe_to_rows(an_df, index=False, header=True):
                    ws_an.append(row_data)

                output = BytesIO()
                wb.save(output)
                st.success("הסתיים בהצלחה!")
                st.download_button("📥 הורד קובץ מעובד", output.getvalue(), "Final_Order.xlsx")
