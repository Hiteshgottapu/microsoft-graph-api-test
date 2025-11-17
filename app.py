"""Flask Resume Parser with Microsoft Graph API
Automated OAuth + Resume Parsing from OneDrive/SharePoint
"""

from flask import Flask, session, redirect, request, url_for, jsonify, render_template_string
import msal
import requests
import os
from datetime import timedelta
from microsoft_graph_api import MicrosoftGraphAPI
from resume_parser import ResumeParser

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "CHANGE_ME_FOR_PRODUCTION")
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=24)

# Azure AD Configuration
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "YOUR_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "YOUR_CLIENT_SECRET")
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "YOUR_TENANT_ID")
REDIRECT_URI = os.environ.get("REDIRECT_URI", "http://localhost:5000/callback")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["User.Read", "Mail.Read", "Files.Read.All", "Sites.Read.All"]

# Simple HTML templates (you can move these to separate template files)
INDEX_HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Resume Parser - Microsoft Graph API</title>
    <style>
        body { font-family: Arial; text-align: center; padding: 50px; background: #f5f5f5; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 40px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; }
        .btn { display: inline-block; padding: 15px 30px; background: #0078d4; color: white; text-decoration: none; border-radius: 5px; font-size: 18px; margin-top: 20px; }
        .btn:hover { background: #005a9e; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📄 Resume Parser</h1>
        <p>Parse resumes from OneDrive/SharePoint using Microsoft Graph API</p>
        <p>Extract: Name, Email, Phone, Skills, Experience, Education</p>
        <a href="/login" class="btn">🔗 Connect Microsoft Account</a>
    </div>
</body>
</html>
"""

DASHBOARD_HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Resume Parser Dashboard</title>
    <style>
        body { font-family: Arial; padding: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; }
        h1 { color: #0078d4; }
        .btn { padding: 10px 20px; background: #0078d4; color: white; border: none; border-radius: 5px; cursor: pointer; margin: 5px; }
        .btn:hover { background: #005a9e; }
        .btn-danger { background: #d13438; }
        #results { margin-top: 20px; }
        .resume-card { background: white; padding: 20px; margin: 10px 0; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .resume-card h3 { color: #0078d4; margin-top: 0; }
        .skills { display: flex; flex-wrap: wrap; gap: 5px; margin: 10px 0; }
        .skill-tag { background: #e1f3ff; padding: 5px 10px; border-radius: 3px; font-size: 12px; }
        .loading { text-align: center; padding: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Resume Parser Dashboard</h1>
        <div>
            <button class="btn" onclick="fetchResumes()">📁 Find Resumes</button>
            <button class="btn" onclick="parseAllResumes()">🔍 Parse All Resumes</button>
            <a href="/logout" class="btn btn-danger">🚪 Logout</a>
        </div>
        <div id="results"></div>
    </div>

    <script>
        async function fetchResumes() {
            document.getElementById('results').innerHTML = '<div class="loading">🔄 Searching for resumes...</div>';
            try {
                const response = await fetch('/api/resumes');
                const data = await response.json();
                if (data.success) {
                    displayResumeList(data.resumes);
                } else {
                    document.getElementById('results').innerHTML = '<div class="resume-card">Error: ' + data.error + '</div>';
                }
            } catch (error) {
                document.getElementById('results').innerHTML = '<div class="resume-card">Error: ' + error.message + '</div>';
            }
        }

        function displayResumeList(resumes) {
            let html = '<h2>📄 Found ' + resumes.length + ' Resume Files</h2>';
            resumes.forEach(resume => {
                html += `
                    <div class="resume-card">
                        <h3>${resume.name}</h3>
                        <p>Size: ${(resume.size / 1024).toFixed(2)} KB | Modified: ${new Date(resume.modified).toLocaleDateString()}</p>
                        <button class="btn" onclick="parseResume('${resume.id}', '${resume.name}')">🔍 Parse This Resume</button>
                    </div>
                `;
            });
            document.getElementById('results').innerHTML = html;
        }

        async function parseResume(fileId, fileName) {
            document.getElementById('results').innerHTML = '<div class="loading">🔄 Parsing ' + fileName + '...</div>';
            try {
                const response = await fetch('/api/parse/' + fileId);
                const data = await response.json();
                if (data.success) {
                    displayParsedResume(data.data);
                } else {
                    document.getElementById('results').innerHTML = '<div class="resume-card">Error: ' + data.error + '</div>';
                }
            } catch (error) {
                document.getElementById('results').innerHTML = '<div class="resume-card">Error: ' + error.message + '</div>';
            }
        }

        async function parseAllResumes() {
            document.getElementById('results').innerHTML = '<div class="loading">🔄 Parsing all resumes...</div>';
            try {
                const response = await fetch('/api/parse-all');
                const data = await response.json();
                if (data.success) {
                    displayAllParsedResumes(data.resumes);
                } else {
                    document.getElementById('results').innerHTML = '<div class="resume-card">Error: ' + data.error + '</div>';
                }
            } catch (error) {
                document.getElementById('results').innerHTML = '<div class="resume-card">Error: ' + error.message + '</div>';
            }
        }

        function displayParsedResume(resume) {
            let html = `
                <div class="resume-card">
                    <h3>📄 ${resume.filename}</h3>
                    <p><strong>👤 Name:</strong> ${resume.name || 'Not found'}</p>
                    <p><strong>📧 Email:</strong> ${resume.email || 'Not found'}</p>
                    <p><strong>📞 Phone:</strong> ${resume.phone || 'Not found'}</p>
                    <p><strong>⏱️ Experience:</strong> ${resume.experience_years || 'Not specified'}</p>
                    <p><strong>🎓 Education:</strong> ${resume.education ? resume.education.join(', ') : 'Not found'}</p>
                    <p><strong>💼 Skills (${resume.skills ? resume.skills.length : 0}):</strong></p>
                    <div class="skills">
                        ${resume.skills ? resume.skills.map(skill => '<span class="skill-tag">' + skill + '</span>').join('') : 'None'}
                    </div>
                </div>
            `;
            document.getElementById('results').innerHTML = html;
        }

        function displayAllParsedResumes(resumes) {
            let html = '<h2>✅ Parsed ' + resumes.length + ' Resumes</h2>';
            resumes.forEach(resume => {
                if (resume.error) {
                    html += `<div class="resume-card"><h3>${resume.filename}</h3><p style="color:red;">Error: ${resume.error}</p></div>`;
                } else {
                    html += `
                        <div class="resume-card">
                            <h3>📄 ${resume.filename}</h3>
                            <p><strong>👤</strong> ${resume.name || 'N/A'} | <strong>📧</strong> ${resume.email || 'N/A'} | <strong>📞</strong> ${resume.phone || 'N/A'}</p>
                            <p><strong>💼 Skills:</strong> ${resume.skills ? resume.skills.slice(0, 10).join(', ') : 'None'}</p>
                        </div>
                    `;
                }
            });
            document.getElementById('results').innerHTML = html;
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    if session.get('access_token'):
        return redirect(url_for('dashboard'))
    return render_template_string(INDEX_HTML)

@app.route('/login')
def login():
    session.permanent = True
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    flow = msal_app.initiate_auth_code_flow(SCOPE, redirect_uri=REDIRECT_URI)
    session['auth_flow'] = flow
    return redirect(flow['auth_uri'])

@app.route('/callback')
def callback():
    try:
        msal_app = msal.ConfidentialClientApplication(
            CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
        )
        result = msal_app.acquire_token_by_auth_code_flow(
            session.get('auth_flow', {}), request.args
        )
        if "access_token" in result:
            session['access_token'] = result['access_token']
            return redirect(url_for('dashboard'))
        else:
            return f"Auth Error: {result.get('error')}"
    except Exception as e:
        return f"Exception: {str(e)}"

@app.route('/dashboard')
def dashboard():
    if not session.get('access_token'):
        return redirect(url_for('login'))
    return render_template_string(DASHBOARD_HTML)

@app.route('/api/resumes')
def get_resumes():
    """Fetch resume files from OneDrive"""
    access_token = session.get('access_token')
    if not access_token:
        return jsonify({'error': 'Not authenticated'}), 401
    
    try:
        graph_api = MicrosoftGraphAPI(access_token)
        search_results = graph_api.search_onedrive("(pdf OR docx) AND (resume OR cv)")
        
        resume_files = []
        for file in search_results.get('value', []):
            if file['name'].lower().endswith(('.pdf', '.docx', '.doc')):
                resume_files.append({
                    'id': file['id'],
                    'name': file['name'],
                    'size': file.get('size', 0),
                    'modified': file.get('lastModifiedDateTime', '')
                })
        
        return jsonify({'success': True, 'count': len(resume_files), 'resumes': resume_files})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/parse/<file_id>')
def parse_resume(file_id):
    """Parse a single resume"""
    access_token = session.get('access_token')
    if not access_token:
        return jsonify({'error': 'Not authenticated'}), 401
    
    try:
        graph_api = MicrosoftGraphAPI(access_token)
        parser = ResumeParser(graph_api)
        
        # Get file metadata
        headers = {"Authorization": f"Bearer {access_token}"}
        metadata_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}"
        metadata_response = requests.get(metadata_url, headers=headers)
        file_info = metadata_response.json()
        filename = file_info['name']
        
        # Parse resume using ResumeParser's method
        parsed_data = parser.parse_resume_from_onedrive(file_id, filename)
        parsed_data['file_id'] = file_id
        
        return jsonify({'success': True, 'data': parsed_data})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/parse-all')
def parse_all_resumes():
    """Parse all resumes"""
    access_token = session.get('access_token')
    if not access_token:
        return jsonify({'error': 'Not authenticated'}), 401
    
    try:
        graph_api = MicrosoftGraphAPI(access_token)
        parser = ResumeParser(graph_api)
        
        # Use ResumeParser's bulk parsing method
        parsed_resumes = parser.search_and_parse_all_resumes("resume")
        
        return jsonify({'success': True, 'count': len(parsed_resumes), 'resumes': parsed_resumes})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

if __name__ == '__main__':
    print("\n" + "="*60)
    print("🚀 Resume Parser with Microsoft Graph API")
    print("="*60)
    print("\n📝 Features:")
    print("  ✅ Automated OAuth with Microsoft")
    print("  ✅ Search resumes in OneDrive/SharePoint")
    print("  ✅ Parse PDF and DOCX files")
    print("  ✅ Extract: Name, Email, Phone, Skills, Experience, Education")
    print("\n🔗 Visit: http://localhost:5000")
    print("\n" + "="*60 + "\n")
    app.run(debug=True, port=5000)
