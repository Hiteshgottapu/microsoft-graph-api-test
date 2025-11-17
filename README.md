# microsoft-graph-api-test
Testing Microsoft Graph API integration for Outlook, OneDrive, and SharePoint access using Azure AD OAuth

## 🌟 Features

### 1. **User Profile** - View your Microsoft account information
### 2. **Outlook Emails** - Browse and search your emails
### 3. **OneDrive Files** - Access your cloud storage files
### 4. **SharePoint Sites** - View your SharePoint sites and document libraries
### 5. **📄 Resume Parser** - Parse PDF/DOCX resumes from OneDrive/SharePoint

## 📄 Resume Parser Capabilities

The Resume Parser module allows you to:
- **Search & Parse All Resumes**: Automatically find and parse all resume files from OneDrive
- **Extract Structured Data**: Name, email, phone, skills, experience, education
- **Support Multiple Formats**: PDF and DOCX files
- **Bulk Processing**: Parse multiple resumes at once
- **Manual Selection**: Choose specific files to parse
- **Skills Detection**: Identifies 40+ technical skills including Python, SQL, Machine Learning, etc.
- **Experience Extraction**: Detects years of experience
- **Education Parsing**: Finds degrees and certifications

### Extracted Information:
- ✅ Candidate Name
- ✅ Email Address
- ✅ Phone Number
- ✅ Technical Skills (Python, SQL, ML, Cloud, etc.)
- ✅ Years of Experience

- - ✅ Education (B.Tech, M.Tech, MBA, PhD, etc.)
- ✅ Raw Text Preview

## 💻 Installation

### 1. Clone the repository:
```bash
git clone https://github.com/Hiteshgottapu/microsoft-graph-api-test.git
cd microsoft-graph-api-test
```

### 2. Install dependencies:
```bash
pip install -r requirements.txt
```

### 3. Set up Azure AD App Registration:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
3. Configure:
   - **Name**: Your app name
   - **Supported account types**: Accounts in any organizational directory
   - **Redirect URI**: `http://localhost:5000/callback`
4. After registration:
   - Copy **Application (client) ID**
   - Copy **Directory (tenant) ID**
   - Create a **Client Secret** in "Certificates & secrets"
5. Add API permissions:
   - Microsoft Graph: `User.Read`, `Mail.Read`, `Files.Read.All`, `Sites.Read.All`
   - Grant admin consent

### 4. Configure environment variables:

Create a `.env` file or set environment variables:
```bash
AZURE_CLIENT_ID=your_client_id
AZURE_CLIENT_SECRET=your_client_secret
AZURE_TENANT_ID=your_tenant_id
REDIRECT_URI=http://localhost:5000/callback
FLASK_SECRET=your_random_secret_key
```

### 5. Run the Flask app:
```bash
python app.py
```

### 6. Open your browser:
Navigate to `http://localhost:5000`

---

## 🔐 Authentication Flow (Fully Automated)

1. Click **"Connect Microsoft Account"** button
2. You'll be redirected to Microsoft login page
3. Grant permissions to the app
4. You'll be redirected back to the dashboard automatically
5. **No manual copy-paste needed!**

---

## ✨ Features Available:

- **Dashboard**: View your profile information
- **Emails**: Browse recent Outlook emails
- **Files**: Access OneDrive files
- **Logout**: Clear session and disconnect

---

## 🚀 Tech Stack

- **Flask**: Backend web framework
- **MSAL**: Microsoft Authentication Library
- **Microsoft Graph API**: Access Microsoft 365 services
- **Python 3.8+**: Core language

---

## 📝 Notes

- All OAuth tokens are stored securely in Flask sessions
- Tokens are automatically refreshed when needed
- No user credentials or secrets are exposed to the browser
- For production, use HTTPS and a proper secret management solution

---

## 🤝 Contributing

Contributions are welcome! Feel free to open issues or submit pull requests.

---

## 📄 License

MIT License - feel free to use this project for learning and development!
- ✅ Education (B.Tech, M.Tech, MBA, PhD, etc.)
- ✅ Raw Text Preview

## 💾 Installation
