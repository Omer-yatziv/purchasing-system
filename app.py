import streamlit as st
import pandas as pd
from thefuzz import fuzz, process
from io import BytesIO
import openpyxl

st.set_page_config(page_title="מערכת רכש חכמה", layout="wide")
st.title("🎯 מערכת הצלבת רכש - גרסה יציבה")

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
        start_row = st.number_input("שורת כותרות (קבוצה, ח\"ג...)", value=10)
        
        # טעינה בפורמט בטוח
        wb = openpyxl.load_workbook(purchase_file, data_only=False, keep_vba=True)
        ws = wb.active

        if st.button("בצע הצלבה חכמה"):
            erp_choices = erp_df['תיאור פריט'].astype(str).tolist()
            results_analysis = []
            
            # מיפוי עמודות המקור
            header_map = {}
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row=start_row, column=col).value
                if val:
                    header_map[str(val).strip()] = col

            needed = ['קבוצה', 'ח"ג', 'תאור/מידה']
            if not all(n in header_map for n in needed):
                st.error(f"חסרות עמודות בשורה {start_row}. וודא שכתוב שם: {needed}")
            else:
                # כתיבת כותרות בעמודות A ו-B (החל משורת הכותרת בלבד)
                ws.cell(row=start_row, column=1).value = "קוד פריט סאפ"
                ws.cell(row=start_row, column=2).value = "תיאור סאפ"

                # מעבר על הנתונים
                for r in range(start_row + 1, ws.max_row + 1):
                    q_group = str(ws.cell(row=r, column=header_map['קבוצה']).value or "")
                    q_mat = str(ws.cell(row=r, column=header_map['ח"ג']).value or "")
                    q_desc = str(ws.cell(row=r, column=header_map['תאור/מידה']).value or "")
                    
                    if q_desc.strip() in ["", "None", "nan"]: continue
                    
                    query = f"{q_group} {q_mat} {q_desc}".strip()
                    match = process.extractOne(query, erp_choices, scorer=fuzz.token_set_ratio)
                    
                    if match and match[1] > 45:
                        best_text, score = match[0], match[1]
                        erp_row = erp_df[erp_df['תיאור פריט'] == best_text].iloc[0]
                        
                        # הזרקת הנתונים לעמודה A ו-B בשורה הנוכחית בלבד
                        ws.cell(row=r, column=1).value = erp_row['קוד פריט']
                        ws.cell(row=r, column=2).value = erp_row['תיאור פריט']
                        alt = process.extract(query, erp_choices, limit=3)
                        alt_text = ", ".join([f"{m[0]} ({m[1]}%)" for m in alt[1:]])
                    else:
                        ws.cell(row=r, column=1).value = "לא נמצא"
                        score, alt_text = 0, ""

                    results_analysis.append({'שורה': r, 'שאילתה': query, 'זיהוי': best_text if score > 0 else "---", 'ציון': score, 'חלופות': alt_text})

                # יצירת לשונית ניתוח נקייה
                if "ניתוח נתונים" in wb.sheetnames: del wb["ניתוח נתונים"]
                ws_an = wb.create_sheet("ניתוח נתונים")
                headers_an = ['שורה', 'שאילתה', 'זיהוי', 'ציון', 'חלופות']
                ws_an.append(headers_an)
                for item in results_analysis:
                    ws_an.append([item[h] for h in headers_an])

                output = BytesIO()
                wb.save(output)
                st.success("העיבוד הסתיים!")
                st.download_button("📥 הורד קובץ מתוקן", output.getvalue(), "Purchase_Order_Fixed.xlsx")
