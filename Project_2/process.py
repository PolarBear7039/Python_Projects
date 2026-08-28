import os
import pandas as pd

def count_duplicates():
    input_file = 'A1.xlsx'
    output_file = 'A1_duplicates.xlsx'
 
    if not os.path.exists(input_file):
        print(f"خطأ: الملف '{input_file}' غير موجود في هذا المجلد!")
        return
    try:
        df = pd.read_excel(input_file)
        
        if df.empty:
            print("الملف فارغ!")
            return
            
        column_name = df.columns[0]
        
        df_counts = df[column_name].value_counts().reset_index()
        
        df_counts.columns = [column_name, 'Count (التكرار)']
        
        df_counts.to_excel(output_file, index=False)
        print(f"تم بنجاح! تم إنشاء ملف التكرارات باسم: {output_file}")
        
    except Exception as e:
        print(f"حدث خطأ أثناء معالجة الملف: {e}")

if __name__ == '__main__':
    count_duplicates()
