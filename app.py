"""Streamlit App for Microsoft Graph API Testing

Run with: streamlit run app.py
"""

import streamlit as st
from microsoft_graph_api import MicrosoftGraphAPI
import os

# Page config
st.set_page_config(
    page_title="Microsoft Graph API Test",
    page_icon="📧",
    layout="wide"
)

# Title
st.title("📧 Microsoft Graph API - Outlook, OneDrive, SharePoint")
st.markdown("---")

# Initialize session state
if 'graph' not in st.session_state:
    st.session_state.graph = MicrosoftGraphAPI()

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

# Sidebar for authentication
with st.sidebar:
    st.header("🔐 Authentication")
    
    if not st.session_state.authenticated:
        st.write("**Step 1:** Click the button below to get the login URL")
        
        if st.button("🔗 Get Login URL"):
            auth_url = st.session_state.graph.get_auth_url()
            st.code(auth_url, language="text")
            st.info("👆 Copy and open this URL in your browser to login")
        
        st.write("---")
        st.write("**Step 2:** After login, paste the authorization code here:")
        auth_code = st.text_input("Authorization Code", type="password")
        
        if st.button("✅ Authenticate"):
            if auth_code:
                try:
                    with st.spinner("Authenticating..."):
                        result = st.session_state.graph.get_token_from_code(auth_code)
                        st.session_state.authenticated = True
                        st.success("✅ Authentication successful!")
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ Authentication failed: {str(e)}")
            else:
                st.warning("Please enter the authorization code")
    else:
        st.success("✅ Authenticated")
        if st.button("🚪 Logout"):
            st.session_state.authenticated = False
            st.session_state.graph.token = None
            st.rerun()

# Main content
if st.session_state.authenticated:
    # Create tabs
    tab1, tab2, tab3, tab4 = st.tabs(["👤 Profile", "📧 Emails", "📁 OneDrive", "🏢 SharePoint"])
    
    # Tab 1: User Profile
    with tab1:
        st.header("User Profile")
        try:
            profile = st.session_state.graph.get_user_profile()
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Name", profile.get('displayName', 'N/A'))
                st.metric("Email", profile.get('mail', 'N/A'))
            with col2:
                st.metric("Job Title", profile.get('jobTitle', 'N/A'))
                st.metric("Office Location", profile.get('officeLocation', 'N/A'))
                
            with st.expander("📋 Full Profile Data"):
                st.json(profile)
        except Exception as e:
            st.error(f"Error fetching profile: {str(e)}")
    
    # Tab 2: Emails
    with tab2:
        st.header("Outlook Emails")
        
        col1, col2 = st.columns([3, 1])
        with col1:
            email_count = st.slider("Number of emails to fetch", 5, 50, 10)
        with col2:
            if st.button("🔄 Refresh Emails"):
                st.rerun()
        
        try:
            emails = st.session_state.graph.get_emails(top=email_count)
            
            if emails and 'value' in emails:
                st.success(f"Found {len(emails['value'])} emails")
                
                for i, email in enumerate(emails['value'], 1):
                    with st.expander(f"📧 {i}. {email.get('subject', 'No Subject')}"):
                        col_a, col_b = st.columns([2, 1])
                        with col_a:
                            st.write(f"**From:** {email['from']['emailAddress']['address']}")
                        with col_b:
                            st.write(f"**Date:** {email.get('receivedDateTime', 'N/A')}")
                        
                        st.write("**Preview:**")
                        st.write(email.get('bodyPreview', 'No preview available'))
                        
                        if st.button(f"View Full Email #{i}", key=f"email_{i}"):
                            st.json(email)
            else:
                st.info("No emails found")
        except Exception as e:
            st.error(f"Error fetching emails: {str(e)}")
    
    # Tab 3: OneDrive
    with tab3:
        st.header("OneDrive Files")
        
        if st.button("🔄 Refresh Files"):
            st.rerun()
        
        try:
            files = st.session_state.graph.get_onedrive_root()
            
            if files and 'value' in files:
                st.success(f"Found {len(files['value'])} items")
                
                for file in files['value']:
                    file_type = "📁" if 'folder' in file else "📄"
                    file_name = file.get('name', 'Unnamed')
                    file_size = file.get('size', 0)
                    
                    with st.expander(f"{file_type} {file_name}"):
                        col_a, col_b = st.columns(2)
                        with col_a:
                            st.write(f"**Size:** {file_size:,} bytes")
                        with col_b:
                            st.write(f"**Modified:** {file.get('lastModifiedDateTime', 'N/A')}")
                        
                        if st.button(f"View Details", key=f"file_{file['id']}"):
                            st.json(file)
            else:
                st.info("No files found")
        except Exception as e:
            st.error(f"Error fetching files: {str(e)}")
    
    # Tab 4: SharePoint
    with tab4:
        st.header("SharePoint Sites")
        
        if st.button("🔄 Refresh Sites"):
            st.rerun()
        
        try:
            sites = st.session_state.graph.get_sharepoint_sites()
            
            if sites and 'value' in sites:
                st.success(f"Found {len(sites['value'])} sites")
                
                for site in sites['value']:
                    site_name = site.get('displayName', 'Unnamed Site')
                    site_url = site.get('webUrl', 'N/A')
                    
                    with st.expander(f"🏢 {site_name}"):
                        st.write(f"**URL:** {site_url}")
                        st.write(f"**Description:** {site.get('description', 'No description')}")
                        
                        if st.button(f"View Site Details", key=f"site_{site['id']}"):
                            st.json(site)
            else:
                st.info("No SharePoint sites found")
        except Exception as e:
            st.error(f"Error fetching sites: {str(e)}")

else:
    # Show instructions when not authenticated
    st.info("👈 Please authenticate using the sidebar to access Microsoft services")
    
    st.markdown("""    ### 📝 Setup Instructions:
    
    1. **Get your Azure credentials:**
       - Client ID: `92149bb7-b052-4d4b-9d39-6cb1440716db`
       - Tenant ID: `2224abd7-7085-434d-b50a-add325728a07`
    
    2. **Create a `.env` file** (optional):
       ```
       MICROSOFT_CLIENT_ID=your_client_id
       MICROSOFT_CLIENT_SECRET=your_client_secret
       MICROSOFT_TENANT_ID=your_tenant_id
       MICROSOFT_REDIRECT_URI=http://localhost:8000/oauth/callback/microsoft
       ```
    
    3. **Install dependencies:**
       ```bash
       pip install -r requirements.txt
       ```
    
    4. **Run the app:**
       ```bash
       streamlit run app.py
       ```
    
    ### 🔐 Authentication Flow:
    
    1. Click "Get Login URL" in the sidebar
    2. Copy and open the URL in your browser
    3. Login with your Microsoft account
    4. Copy the authorization code from the redirect URL
    5. Paste it in the sidebar and click "Authenticate"
    
    ### ✨ Features:
    
    - **Profile**: View your Microsoft account details
    - **Emails**: Browse your Outlook inbox
    - **OneDrive**: View your files and folders
    - **SharePoint**: Access your SharePoint sites
    """)

# Footer
st.markdown("---")
st.caption("Built with Streamlit | Microsoft Graph API Integration")
