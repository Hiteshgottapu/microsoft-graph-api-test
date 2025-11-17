"""Flask App for Automated Microsoft Graph API OAuth

Run with: python app.py
Then visit: http://localhost:5000/login
"""

from flask import Flask, session, redirect, request, url_for, jsonify
import msal
import requests
import os
from datetime import timedelta

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "CHANGE_ME_FOR_PRODUCTION")
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=24)

# Azure AD Configuration
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "YOUR_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "YOUR_CLIENT_SECRET")
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "YOUR_TENANT_ID")
REDIRECT_URI = os.environ.get("REDIRECT_URI", "http://localhost:5000/callback")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = [
    "User.Read",
    "Mail.Read",
    "Files.Read.All",
    "Sites.Read.All"
]

@app.route('/')
def index():
    if session.get('access_token'):
        return '''<html>
            <body style="font-family: Arial; padding: 50px; text-align: center;">
                <h1>✅ Connected to Microsoft Graph API</h1>
                <p>You are authenticated!</p>
                <a href="/dashboard" style="padding: 10px 20px; background: #0078d4; color: white; text-decoration: none; border-radius: 5px;">Go to Dashboard</a>
                <br><br>
                <a href="/logout" style="padding: 10px 20px; background: #d13438; color: white; text-decoration: none; border-radius: 5px;">Logout</a>
            </body>
        </html>'''
    return '''<html>
        <body style="font-family: Arial; padding: 50px; text-align: center;">
            <h1>🔐 Microsoft Graph API OAuth Test</h1>
            <p>Click below to connect your Microsoft account</p>
            <a href="/login" style="padding: 15px 30px; background: #0078d4; color: white; text-decoration: none; border-radius: 5px; font-size: 18px;">🔗 Connect Microsoft Account</a>
        </body>
    </html>'''

@app.route('/login')
def login():
    session.permanent = True
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, 
        authority=AUTHORITY, 
        client_credential=CLIENT_SECRET
    )
    flow = msal_app.initiate_auth_code_flow(SCOPE, redirect_uri=REDIRECT_URI)
    session['auth_flow'] = flow
    return redirect(flow['auth_uri'])

@app.route('/callback')
def callback():
    try:
        msal_app = msal.ConfidentialClientApplication(
            CLIENT_ID, 
            authority=AUTHORITY, 
            client_credential=CLIENT_SECRET
        )
        result = msal_app.acquire_token_by_auth_code_flow(
            session.get('auth_flow', {}), 
            request.args
        )
        
        if "access_token" in result:
            session['access_token'] = result['access_token']
            session['refresh_token'] = result.get('refresh_token', '')
            return redirect(url_for('dashboard'))
        else:
            error_msg = result.get('error', 'Unknown error')
            error_desc = result.get('error_description', '')
            return f'''<html>
                <body style="font-family: Arial; padding: 50px; text-align: center;">
                    <h1>❌ Authentication Error</h1>
                    <p><strong>Error:</strong> {error_msg}</p>
                    <p>{error_desc}</p>
                    <a href="/">Go back</a>
                </body>
            </html>'''
    except Exception as e:
        return f'''<html>
            <body style="font-family: Arial; padding: 50px; text-align: center;">
                <h1>❌ Exception Occurred</h1>
                <p>{str(e)}</p>
                <a href="/">Go back</a>
            </body>
        </html>'''

@app.route('/dashboard')
def dashboard():
    access_token = session.get('access_token')
    if not access_token:
        return redirect(url_for('login'))
    
    headers = {"Authorization": f"Bearer {access_token}"}
    
    # Fetch user profile
    try:
        user_resp = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers)
        user_data = user_resp.json()
        
        name = user_data.get('displayName', 'Unknown')
        email = user_data.get('mail') or user_data.get('userPrincipalName', 'N/A')
        
        return f'''<html>
            <body style="font-family: Arial; padding: 50px;">
                <h1>📊 Dashboard</h1>
                <div style="background: #f3f2f1; padding: 20px; border-radius: 8px; margin: 20px 0;">
                    <h2>👤 Profile</h2>
                    <p><strong>Name:</strong> {name}</p>
                    <p><strong>Email:</strong> {email}</p>
                </div>
                <div style="margin: 20px 0;">
                    <a href="/api/emails" style="padding: 10px 20px; background: #0078d4; color: white; text-decoration: none; border-radius: 5px; margin-right: 10px;">📧 View Emails</a>
                    <a href="/api/files" style="padding: 10px 20px; background: #0078d4; color: white; text-decoration: none; border-radius: 5px; margin-right: 10px;">📁 View Files</a>
                    <a href="/logout" style="padding: 10px 20px; background: #d13438; color: white; text-decoration: none; border-radius: 5px;">🚪 Logout</a>
                </div>
            </body>
        </html>'''
    except Exception as e:
        return f'''<html>
            <body style="font-family: Arial; padding: 50px;">
                <h1>❌ Error fetching profile</h1>
                <p>{str(e)}</p>
                <a href="/logout">Logout and retry</a>
            </body>
        </html>'''

@app.route('/api/emails')
def get_emails():
    access_token = session.get('access_token')
    if not access_token:
        return redirect(url_for('login'))
    
    headers = {"Authorization": f"Bearer {access_token}"}
    
    try:
        emails_resp = requests.get(
            "https://graph.microsoft.com/v1.0/me/messages?$top=10",
            headers=headers
        )
        emails_data = emails_resp.json()
        
        html = '''<html>
            <body style="font-family: Arial; padding: 50px;">
                <h1>📧 Recent Emails</h1>
                <a href="/dashboard" style="padding: 10px 20px; background: #0078d4; color: white; text-decoration: none; border-radius: 5px;">← Back to Dashboard</a>
                <div style="margin-top: 20px;">
        '''
        
        for email in emails_data.get('value', []):
            subject = email.get('subject', 'No Subject')
            sender = email.get('from', {}).get('emailAddress', {}).get('address', 'Unknown')
            received = email.get('receivedDateTime', '')
            
            html += f'''<div style="background: #f3f2f1; padding: 15px; margin: 10px 0; border-radius: 5px;">
                <p><strong>Subject:</strong> {subject}</p>
                <p><strong>From:</strong> {sender}</p>
                <p><strong>Received:</strong> {received}</p>
            </div>'''
        
        html += '</div></body></html>'
        return html
    except Exception as e:
        return f'''<html>
            <body style="font-family: Arial; padding: 50px;">
                <h1>❌ Error fetching emails</h1>
                <p>{str(e)}</p>
                <a href="/dashboard">Back to Dashboard</a>
            </body>
        </html>'''

@app.route('/api/files')
def get_files():
    access_token = session.get('access_token')
    if not access_token:
        return redirect(url_for('login'))
    
    headers = {"Authorization": f"Bearer {access_token}"}
    
    try:
        files_resp = requests.get(
            "https://graph.microsoft.com/v1.0/me/drive/root/children",
            headers=headers
        )
        files_data = files_resp.json()
        
        html = '''<html>
            <body style="font-family: Arial; padding: 50px;">
                <h1>📁 OneDrive Files</h1>
                <a href="/dashboard" style="padding: 10px 20px; background: #0078d4; color: white; text-decoration: none; border-radius: 5px;">← Back to Dashboard</a>
                <div style="margin-top: 20px;">
        '''
        
        for file in files_data.get('value', []):
            name = file.get('name', 'Unknown')
            size = file.get('size', 0)
            file_type = 'Folder' if 'folder' in file else 'File'
            
            html += f'''<div style="background: #f3f2f1; padding: 15px; margin: 10px 0; border-radius: 5px;">
                <p><strong>Name:</strong> {name}</p>
                <p><strong>Type:</strong> {file_type}</p>
                <p><strong>Size:</strong> {size} bytes</p>
            </div>'''
        
        html += '</div></body></html>'
        return html
    except Exception as e:
        return f'''<html>
            <body style="font-family: Arial; padding: 50px;">
                <h1>❌ Error fetching files</h1>
                <p>{str(e)}</p>
                <a href="/dashboard">Back to Dashboard</a>
            </body>
        </html>'''

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

if __name__ == '__main__':
    print("\n" + "="*60)
    print("🚀 Microsoft Graph API OAuth Test Server")
    print("="*60)
    print("\n📝 Setup Instructions:")
    print("1. Set environment variables:")
    print("   - AZURE_CLIENT_ID")
    print("   - AZURE_CLIENT_SECRET")
    print("   - AZURE_TENANT_ID")
    print("   - REDIRECT_URI (default: http://localhost:5000/callback)")
    print("\n2. Make sure Azure redirect URI matches: http://localhost:5000/callback")
    print("\n3. Visit: http://localhost:5000")
    print("\n" + "="*60 + "\n")
    app.run(debug=True, port=5000)
