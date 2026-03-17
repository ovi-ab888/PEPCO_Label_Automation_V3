PEPCO Automation App

🚀 PEPCO_SS26 Updated হলো একটি Streamlit-ভিত্তিক অটোমেশন অ্যাপ, যা PEPCO-এর ডেটা প্রসেসিং ও CSV রিপোর্ট জেনারেশনের কাজ সহজ করে।

✨ Features

📂 PDF Upload → এক বা একাধিক PEPCO Data PDF ফাইল আপলোড করা যায়

🏷️ Department & Product Selector → ডিপার্টমেন্ট ও প্রোডাক্ট টাইপ সিলেক্ট

🧵 Material Composition Auto-Logic

100% হলে আর কোনো ইনপুট বক্স আসে না

যদি % < 100 হয় → পরবর্তী material input নিজে থেকে আসে

মোট composition 100% না হওয়া পর্যন্ত চলবে

🧼 Washing Code Mapping → সঠিক washing symbol কোড জেনারেট হয়

💰 Price Ladder → PLN প্রাইস ইনপুট দিলে অন্যান্য currency অটো-জেনারেট হয়

📝 Editable Data → ফাইনাল CSV ডাউনলোড করার আগে ডেটা এডিট করা যায়

🔑 Password Protection (ঐচ্ছিক)

🎨 Custom UI/Theme

🛠️ Installation (Local)

Repo ক্লোন করো:

git clone https://github.com/ovi-ab888/PEPCO_SS26_updated.git
cd PEPCO_SS26_updated


ভার্চুয়াল এনভ তৈরি করে active করো:

python -m venv .venv
source .venv/bin/activate   # Linux/Mac
.venv\Scripts\activate      # Windows


ডিপেন্ডেন্সি ইন্সটল করো:

pip install -r requirements.txt


অ্যাপ চালু করো:

streamlit run app.py

☁️ Deployment (Streamlit Cloud)

requirements.txt (dependencies):

streamlit==1.39.0
pandas==2.2.2
pymupdf==1.24.10
openpyxl
reportlab
requests==2.32.3
gspread
oauth2client


runtime.txt (Python version lock):

3.12


Secrets (Streamlit Cloud → Settings → Secrets):

app_password = "YOUR_PASSWORD"

📂 Project Structure
PEPCO_SS26_updated/
│
├── app.py                 # Main Streamlit app
├── auth.py                # Authentication (password check)
├── theme.py               # Custom UI/UX
├── logo.svg               # Logo
├── pepco_ui_hide_github.py
│
├── requirements.txt       # Dependencies
├── runtime.txt            # Python version (for Streamlit Cloud)
├── secrets.toml           # Local secrets config
│
├── README.md
└── .gitignore

👨‍💻 Author

Developed by Ovi AbdulOhab
