# 📧 Outlook Draft Creator# 📧 Outlook Draft Creator# 📧 Outlook Draft Creator



Application for creating email drafts in local Outlook from a Web interface.



---Application for creating email drafts in local Outlook from a Web interface.אפליקציה ליצירת טיוטות מייל ב-Outlook המקומי מתוך ממשק Web.



## 📂 Project Structure



```------

outlook-draft-automation/

├── frontend/              # Web App Interface

├── backend/               # Node.js Server

│   ├── server.js## 📂 Project Structure## 📂 מבנה הפרויקט

│   └── package.json

├── watcher/               # Python Watcher

│   ├── watcher.py

│   └── requirements.txt``````

├── drafts/                # Drafts folder (created automatically)

│   └── attachments/       # Attachment filesoutlook-draft-automation/outlook-draft-automation/

└── README.md

```├── frontend/              # React Web App (in artifact above)├── frontend/              # React Web App (בartifact למעלה)



---├── backend/               # Node.js Server├── backend/               # Node.js Server



## 🚀 Running the Project│   ├── server.js│   ├── server.js



### **Step 1: Install Backend**│   └── package.json│   └── package.json



```bash├── watcher/               # Python Watcher├── watcher/               # Python Watcher

cd backend

npm install│   ├── watcher.py│   ├── watcher.py

npm start

```│   └── requirements.txt│   └── requirements.txt



✅ Backend runs on: `http://localhost:3000`├── drafts/                # Drafts folder (created automatically)├── drafts/                # תיקיית טיוטות (נוצרת אוטומטית)



### **Step 2: Install Watcher**│   └── attachments/       # Attached files│   └── attachments/       # קבצים מצורפים



```bash└── README.md└── README.md

cd watcher

pip install -r requirements.txt``````

python watcher.py

```



✅ Watcher monitors the `drafts/` folder------



### **Step 3: Open Web App**



Open the `frontend/index.html` file in your browser.## 🚀 Running the Project## 🚀 הרצת הפרויקט



---



## 🎯 How it Works### **Step 1: Install Backend**### **שלב 1: התקנת Backend**



1. **User fills out form** in Web App

2. **Web App sends data** to Backend (`localhost:3000`)

3. **Backend creates JSON file** for each recipient in `drafts/` folder```bash```bash

4. **Watcher detects new file** within 1-2 seconds

5. **Watcher opens Outlook** with prepared draftcd backendcd backend

6. **Watcher deletes file** after opening

npm installnpm install

---

npm startnpm start

## 📋 System Requirements

``````

- **Windows** (due to Outlook COM API)

- **Outlook Desktop** installed

- **Node.js** (v14 and above)

- **Python** (3.7 and above)✅ Backend runs on: `http://localhost:3000`✅ Backend רץ על: `http://localhost:3000`

- **pywin32** (installed via requirements.txt)



---

------

## 🔧 Troubleshooting



### ❌ "Connection error: Backend not available"

- Ensure Backend is running on `localhost:3000`### **Step 2: Install Watcher**### **שלב 2: התקנת Watcher**

- Check with: `curl http://localhost:3000/health`



### ❌ "Watcher doesn't detect files"

- Ensure path to `drafts/` folder is correct```bash```bash

- Check write permissions for folder

cd watchercd watcher

### ❌ "Outlook doesn't open"

- Ensure Outlook is installed and configuredpip install -r requirements.txtpip install -r requirements.txt

- Ensure `pywin32` is installed: `pip install pywin32`

- Try running: `python -c "import win32com.client; print('OK')"`python watcher.pypython watcher.py



### ❌ "File not attached"``````

- Ensure file is smaller than 50MB

- Check that file path is correct in JSON



---✅ Watcher monitors the `drafts/` folder✅ Watcher מנטר את תיקיית `drafts/`



## 🎥 Demo



1. Fill form with multiple recipients------

2. Upload PDF file (resume)

3. Click "Create Drafts"

4. Within seconds, Outlook will open with separate drafts for each recipient!

### **Step 3: Open Web App**### **שלב 3: פתיחת Web App**

---



## 🔐 Security

Open the Frontend Artifact in browser (or copy code to a separate React project).פתח את ה-Artifact של ה-Frontend בדפדפן (או העתק את הקוד לפרויקט React נפרד).

- Project runs **locally only** (`localhost`)

- No cloud data storage

- Files are automatically deleted after processing

- Suitable for controlled work environment------



---



## 📝 License## 🎯 How it Works?## 🎯 איך זה עובד?



Free project for personal and business use.



---1. **User fills out form** in Web App1. **המשתמש ממלא טופס** ב-Web App



## 💡 Future Improvements2. **Web App sends data** to Backend (`localhost:3000`)2. **Web App שולח את הנתונים** ל-Backend (`localhost:3000`)



- Mac support (via AppleScript)3. **Backend creates JSON file** for each recipient in `drafts/` folder3. **Backend יוצר קובץ JSON** לכל נמען בתיקיית `drafts/`

- UI for managing draft queue

- Email template support4. **Watcher detects new file** within 1-2 seconds4. **Watcher מזהה את הקובץ החדש** תוך 1-2 שניות

- WebSocket for real-time updates

- History saving5. **Watcher opens Outlook** with prepared draft5. **Watcher פותח Outlook** עם הטיוטה מוכנה



---6. **Watcher deletes file** after opening6. **Watcher מוחק את הקובץ** לאחר הפתיחה



**Good luck! 🚀**

------



## 📋 System Requirements## 📋 דרישות מערכת



- **Windows** (due to Outlook COM API)- **Windows** (בגלל Outlook COM API)

- **Outlook Desktop** installed- **Outlook Desktop** מותקן

- **Node.js** (v14 and above)- **Node.js** (v14 ומעלה)

- **Python** (3.7 and above)- **Python** (3.7 ומעלה)

- **pywin32** (installed via requirements.txt)- **pywin32** (מותקן דרך requirements.txt)



------



## 🔧 Troubleshooting## 🔧 פתרון בעיות נפוצות



### ❌ "Connection error: Backend not available"### ❌ "שגיאת חיבור: Backend לא זמין"

- Ensure Backend is running on `localhost:3000`- ודא ש-Backend רץ על `localhost:3000`

- Check with: `curl http://localhost:3000/health`- בדוק עם: `curl http://localhost:3000/health`



### ❌ "Watcher doesn't detect files"### ❌ "Watcher לא מזהה קבצים"

- Ensure path to `drafts/` folder is correct- ודא שהנתיב לתיקיית `drafts/` נכון

- Check write permissions for folder- בדוק הרשאות כתיבה לתיקייה



### ❌ "Outlook doesn't open"### ❌ "Outlook לא נפתח"

- Ensure Outlook is installed and configured- ודא ש-Outlook מותקן ומוגדר

- Ensure `pywin32` is installed: `pip install pywin32`- ודא ש-`pywin32` מותקן: `pip install pywin32`

- Try running: `python -c "import win32com.client; print('OK')"`- נסה להריץ: `python -c "import win32com.client; print('OK')"`



### ❌ "File not attached"### ❌ "קובץ לא מצורף"

- Ensure file is smaller than 50MB- ודא שהקובץ קטן מ-50MB

- Check that file path is correct in JSON- בדוק שהנתיב לקובץ נכון ב-JSON



------



## 🎥 Demo## 🎥 הדגמה



1. Fill form with multiple recipients1. מלא טופס עם מספר נמענים

2. Upload PDF file (resume)2. העלה קובץ PDF (קורות חיים)

3. Click "Create Drafts"3. לחץ "צור טיוטות"

4. Within seconds, Outlook will open with separate drafts for each recipient!4. תוך שניות, Outlook יפתח עם טיוטות נפרדות לכל נמען!



------



## 🔐 Security## 🔐 אבטחה



- Project runs **locally only** (`localhost`)- הפרויקט פועל **רק מקומית** (`localhost`)

- No cloud data storage- אין שמירת נתונים בענן

- Files are automatically deleted after processing- קבצים נמחקים אוטומטית לאחר עיבוד

- Suitable for controlled work environment- מתאים לסביבת עבודה מבוקרת



------



## 📝 License## 📝 רישיון



Free project for personal and business use.פרויקט חינמי לשימוש אישי ועסקי.



------



## 💡 Future Improvements## 💡 שיפורים עתידיים



- Mac support (via AppleScript)- תמיכה ב-Mac (דרך AppleScript)

- UI for managing draft queue- UI לניהול תור טיוטות

- Email template support- תמיכה בתבניות מייל

- WebSocket for real-time updates- WebSocket לעדכונים בזמן אמת

- History saving- שמירת היסטוריה



------



**Good luck! 🚀****בהצלחה! 🚀**