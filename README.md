# PDF Merger Tool

כלי למיזוג קבצי PDF ו-DOCX עם הוספת כותרת תחתונה אוטומטית.

## תכונות עיקריות

* מיזוג קבצי PDF ו-DOCX לקובץ PDF אחד
* המרה אוטומטית של קבצי DOCX ל-PDF
* הוספת כותרת תחתונה לכל עמוד הכוללת שם הקובץ המקורי ומספר עמוד
* מיון הקבצים לפי סדר אלפביתי
* תמיכה מלאה בעברית

## דרישות מערכת

* Python 3.7 ומעלה
* הספריות הנדרשות מפורטות בקובץ `requirements.txt`

## התקנה

1. שכפל את המאגר:

    git clone https://github.com/your-username/pdf-merger.git
    cd pdf-merger

2. התקן את הספריות הנדרשות:

    pip install -r requirements.txt

3. התקן את הכלי:

    pip install .

## הגדרות

ערוך את קובץ `config.json` והגדר את הנתיבים הרצויים:

    {
        "input_folder": "C:\\Your\\Input\\Folder",
        "output_file": "C:\\Your\\Output\\Folder\\merged_output.pdf",
        "footer_text": "מסמך מאוחד",
        "temp_folder": "C:\\Your\\Temp\\Folder"
    }

## שימוש

1. העתק את הקבצים שברצונך למזג לתיקיית הקלט שהוגדרה ב-`config.json`
2. הרץ את התוכנית:

    python app.py

## יצירת קובץ הרצה עצמאי

ניתן ליצור קובץ הרצה עצמאי (exe) באמצעות PyInstaller:

1. התקן את PyInstaller:

    pip install pyinstaller

2. צור את קובץ ההרצה:

    pyinstaller --onefile --name pdf_merger app.py

הקובץ המוכן יימצא בתיקיית `dist`.

## הערות

* וודא שיש לך הרשאות קריאה וכתיבה בתיקיות שהוגדרו
* הקבצים ימוזגו לפי סדר אלפביתי
* קבצי DOCX יומרו אוטומטית ל-PDF לפני המיזוג
* נוצרת תיקייה זמנית לעיבוד הקבצים שנמחקת בסיום התהליך

## רישיון

MIT License

## תמיכה

אם נתקלת בבעיות או יש לך שאלות, אנא פתח issue במאגר הפרויקט.
