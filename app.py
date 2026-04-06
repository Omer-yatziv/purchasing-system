import streamlit as st
import pandas as pd
import re
from thefuzz import fuzz
from io import BytesIO

# הגדרות עיצוב
st.set_page_config(page_title="מערכת רכש חכמה", layout="wide")
st.title("🎯 מערכת התאמת רכש לסאפ - גרסה 2.0")

# פונקציית עזר לניקוי וזיהוי מידות
def clean_text(text):
    if pd.isna(text): return ""
    text = str(text).lower()
    # הפיכת לוכסן או איקס למקף אחיד לצורך השוואה
    text = text.replace('/', '-').replace('x', '-').replace('*', '-')
    return text

def check_material_conflict(req, erp):
    # רשימת חומרים שאסור לבלבל ביניהם
    conflicts = [
        ("מגולוון", "שחור"),
        ("נירוסטה", "אלומיניום"),
        ("נירוסטה", "ברזל"),
        ("פח", "אום")
    ]
    for m1, m2 in conflicts:
        if (m1 in req and m2 in erp) or (m2 in req and m1 in erp):
            return True
    return False

# --- טעינת בסיס נתונים קבוע ---
@st.cache_data
def load_erp():
    try:
        # המערכת מחפשת את הקובץ שהעלית ל-GitHub
        df = pd.read_excel("erp_master.xlsx")
        return df
    except:
        return None

erp_df = load_erp()

if erp_df is None:
    st.error("⚠️ קובץ erp_master.xlsx לא נמצא ב-GitHub. אנא העלה אותו.")
else:
    st.success(f"בסיס הנתונים ERP נטען (נמצאו {len(erp_df)} פריטים)")

    # --- העלאת בקשת רכש ---
    purchase_file = st.file_uploader("העלה קובץ בקשת רכש (XLSX/XLSM)", type=['xlsx', 'xlsm'])

    if purchase_file:
        # קריאת הקובץ - מתחילים משורה 10 (index 9) כברירת מחדל
        start_row = st.number_input("שורה בה מתחילה הטבלה (10 כברירת מחדל)", value=10) - 1
        purchase_df_full = pd.read_excel(purchase_file, header=None)
        
        # הפרדה בין הכותרת לנתונים
        header_row = purchase_df_full.iloc[start_row]
        data_df = purchase_df_full.iloc[start_row+1:].copy()
        data_df.columns = header_row

        if st.button("בצע התאמה חכמה"):
            results_list = []
            analysis_list = []
            
            progress = st.progress(0)
            rows_to_process = data_df.dropna(subset=[data_df.columns[2]]) # מסנן שורות ריקות

            for idx, row in rows_to_process.iterrows():
                # בניית תיאור מלא מ-3 עמודות: קבוצה, ח"ג, תיאור/מידה
                # הנחה: עמודות אלו הן בסדר מסוים (למשל עמודות 0, 1, 2)
                group_req = str(row.get('קבוצה', ''))
                material_req = str(row.get('ח"ג', ''))
                desc_req = str(row.get('תאור/מידה', ''))
                
                full_query = f"{group_req} {material_req} {desc_req}".strip()
                clean_query = clean_text(full_query)

                # חיפוש ב-ERP
                best_score = 0
                best_match = None
                
                for _, erp_row in erp_df.iterrows():
                    erp_desc = str(erp_row['תיאור פריט'])
                    clean_erp = clean_text(erp_desc)
                    
                    # בדיקת סתירה בחומרים
                    if check_material_conflict(clean_query, clean_erp):
                        continue
                        
                    # חישוב דמיון
                    score = fuzz.token_set_ratio(clean_query, clean_erp)
                    
                    # בונוס על מידות מדויקות (למשל 6 מ"מ)
                    numbers_req = re.findall(r'\d+', clean_query)
                    for num in numbers_req:
                        if num in clean_erp:
                            score += 5
                    
                    if score > best_score:
                        best_score = min(score, 100)
                        best_match = erp_row

                # הכנת שורה לתוצאה
                match_code = best_match['קוד פריט'] if best_match is not None and best_score > 50 else "לא נמצא"
                match_desc = best_match['תיאור פריט'] if best_match is not None and best_score > 50 else ""
                
                # הוספת העמודות החדשות להתחלה
                new_row_data = {
                    'קוד פריט': match_code,
                    'תיאור סאפ': match_desc,
                    '% התאמה': f"{best_score}%"
                }
                # איחוד עם השורה המקורית
                full_new_row = {**new_row_data, **row.to_dict()}
                results_list.append(full_new_row)
                
                # לוג ניתוח
                analysis_list.append({
                    'שורה': idx + 1,
                    'תיאור בבקשה': full_query,
                    'התאמה שנבחרה': match_desc,
                    'ציון': best_score,
                    'סטטוס': "✅" if best_score > 85 else "⚠️" if best_score > 55 else "❌"
                })
                
            # יצירת קבצים
            final_df = pd.DataFrame(results_list)
            analysis_df = pd.DataFrame(analysis_list)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, sheet_name='בקשת רכש מעודכנת', index=False)
                analysis_df.to_excel(writer, sheet_name='ניתוח נתונים', index=False)
            
            st.success("העיבוד הושלם!")
            st.download_button("📥 הורד אקסל מוכן", output.getvalue(), "Purchase_Order_Final.xlsx")
