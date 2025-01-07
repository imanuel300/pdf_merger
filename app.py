import os
import json
import fitz  # PyMuPDF
from docx2pdf import convert
from PyPDF2 import PdfMerger
import tempfile
import shutil
from datetime import datetime

class PDFMergerTool:
    def __init__(self, config_path='config.json'):
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        # יצירת תיקיית temp אם לא קיימת
        os.makedirs(self.config['temp_folder'], exist_ok=True)

    def add_footer(self, input_path, output_path, original_filename, total_pages):
        doc = fitz.open(input_path)
        for page_num in range(len(doc)):
            page = doc[page_num]
            footer_text = f"{original_filename} - עמוד {page_num + 1} מתוך {total_pages}"
            
            # חישוב מיקום ה-footer
            rect = page.rect
            # חישוב מיקום אמצע העמוד
            x = rect.width/2
            y = rect.height - 20
            
            # הוספת הטקסט - הסרנו את הפרמטר align שגרם לשגיאה
            page.insert_text(
                point=fitz.Point(x, y),  # שימוש ב-Point במקום שני פרמטרים נפרדים
                text=footer_text,
                fontsize=10
            )
        
        doc.save(output_path)
        doc.close()

    def process_files(self):
        merger = PdfMerger()
        
        # המרת DOCX ל-PDF
        for filename in os.listdir(self.config['input_folder']):
            if filename.endswith('.docx'):
                docx_path = os.path.join(self.config['input_folder'], filename)
                pdf_path = os.path.join(self.config['temp_folder'], 
                                      filename.replace('.docx', '_temp.pdf'))
                convert(docx_path, pdf_path)
        
        # עיבוד כל קובצי ה-PDF והוספת footer
        processed_files = []
        for filename in sorted(os.listdir(self.config['input_folder']) + 
                             os.listdir(self.config['temp_folder'])):
            if filename.endswith('.pdf') or filename.endswith('_temp.pdf'):
                if filename.endswith('_temp.pdf'):
                    file_path = os.path.join(self.config['temp_folder'], filename)
                    original_filename = filename.replace('_temp.pdf', '.docx')
                else:
                    file_path = os.path.join(self.config['input_folder'], filename)
                    original_filename = filename
                
                # הוספת footer
                doc = fitz.open(file_path)
                total_pages = len(doc)
                doc.close()
                
                processed_path = os.path.join(self.config['temp_folder'], 
                                            f"processed_{filename}")
                self.add_footer(file_path, processed_path, original_filename, total_pages)
                processed_files.append(processed_path)
        
        # מיזוג כל הקבצים
        for file_path in processed_files:
            merger.append(file_path)
        
        # שמירת הקובץ הסופי
        merger.write(self.config['output_file'])
        merger.close()
        
        # ניקוי קבצים זמניים
        shutil.rmtree(self.config['temp_folder'])

def main():
    try:
        merger_tool = PDFMergerTool()
        merger_tool.process_files()
        print("המיזוג הושלם בהצלחה!")
    except Exception as e:
        print(f"אירעה שגיאה: {str(e)}")

if __name__ == "__main__":
    main()