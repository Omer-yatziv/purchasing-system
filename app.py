import streamlit as st
import pandas as pd
from thefuzz import fuzz, process
from io import BytesIO
import openpyxl

st.set_page_config(page_title="מערכת רכש חכמה", layout="wide")
st.title("🎯 מערכת הצלבת רכש - שמירה על מבנה מקורי")

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
    purchase_file = st.file_uploader("העלה בקשת רכש", type=['xlsx', 'xlsm'])
    
    if purchase_file:
        start_row = st.number_input("באיזו שורה נמצאות הכותרות (קבוצה, ח\"ג)?", value=10)
        
        # טעינת הקובץ בצורה מלאה כולל הכל
        wb = openpyxl.load_workbook(purchase_file, keep_vba=True)
        ws = wb.active

        if st.button("בצע הצלבה חכמה"):
            erp_choices = erp_df['תיאור פריט'].astype(str).tolist()
            results_analysis = []
            
            # 1. מציאת מיקומי העמודות בבקשת הרכש
            header_map = {}
            last_col = ws.max_column
            for col in range(1, last_col + 1):
                val = ws.cell(row=start_row, column=col).value
                if val:
                    header_map[str(val).strip()] = col

            needed = ['קבוצה', 'ח"ג', 'תאור/מידה']
            if not all(n in header_map for n in needed):
                st.error(f"לא מצאתי את העמודות: {needed}. וודא שהשמות תואמים בשורה {start_row}.")
            else:
                # 2. הגדרת מיקום הכתיבה - בסוף הטבלה הקיימת כדי לא להרוס מיזוגים בהתחלה
                target_col_code = last_col + 1
                target_col_desc = last_col + 2
                
                # כתיבת כותרות בסוף
                ws.cell(row=start_row, column=target_col_code).value = "קוד פריט סאפ"
                ws.cell(row=start_row, column=target_col_desc).value = "תיאור סאפ"

                # 3. מעבר על השורות
                for r in range(start_row + 1, ws.max_row + 1):
                    # שליפת הנתונים מהעמודות המקוריות
                    q_group = str(ws.cell(row=r, column=header_map['קבוצה']).value or "")
                    q_mat = str(ws.cell(row=r, column=header_map['ח"ג']).value or "")
                    q_desc = str(ws.cell(row=r, column=header_map['תאור/מידה']).value or "")
                    
                    if q_desc.strip() in ["", "None", "nan"]: continue
                    
                    query = f"{q_group} {q_mat} {q_desc}".strip()
                    
                    # חיפוש חכם
                    match = process.extractOne(query, erp_choices, scorer=fuzz.token_set_ratio)
                    
                    if match and match[1] > 45:
                        best_text, score = match[0], match[1]
                        erp_row = erp_df[erp_df['תיאור פריט'] == best_text].iloc[0]
                        
                        # כתיבה לעמודות החדשות בסוף השורה
                        ws.cell(row=r, column=target_col_code).value = erp_row['קוד פריט']
                        ws.cell(row=r, column=target_col_desc).value = erp_row['תיאור פריט']
                        
                        # לוג ללשונית השנייה
                        alt_matches = process.extract(query, erp_choices, limit=3)
                        alt_text = ", ".join([f"{m[0]} ({m[1]}%)" for m in alt_matches[1:]])
                    else:
                        ws.cell(row=r, column=target_col_code).value = "לא נמצא"
                        score = 0
                        alt_text = ""

                    results_analysis.append({
                        'שורה באקסל': r,
                        'תיאור בבקשה': query,
                        'התאמה': ws.cell(row=r, column=target_col_desc).value,
                        'ציון %': score,
                        'חלופות': alt_text
                    })

                # 4. יצירת לשונית ניתוח
                if "ניתוח נתונים" in wb.sheetnames: del wb["ניתוח נתונים"]
                ws_an = wb.create_sheet("ניתוח נתונים")
                an_df = pd.DataFrame(results_analysis)
                from openpyxl.utils.dataframe import dataframe_to_rows
                for row_data in dataframe_to_rows(an_df, index=False, header=True):
                    ws_an.append(row_data)

                output = BytesIO()
                wb.save(output)
                st.success("הצלחה! הנתונים נוספו בסוף הטבלה (מימין) לשמירה על המבנה.")
                st.download_button("📥 הורד קובץ סופי", output.getvalue(), "Purchase_Order_Success.xlsx")
                
