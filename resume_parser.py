"""Resume Parser Module for Microsoft Graph API

This module extracts structured information from resume files stored in 
OneDrive or SharePoint using Microsoft Graph API integration.

Supports: PDF, DOCX formats
Extracts: Name, Email, Phone, Skills, Experience, Education
"""

import PyPDF2
import docx
import io
import re
import os
from typing import Dict, List, Optional


class ResumeParser:
    """Parse resume files from OneDrive/SharePoint and extract structured data"""
    
    def __init__(self, graph_api):
        """
        Initialize Resume Parser with Microsoft Graph API client
        
        Args:
            graph_api: MicrosoftGraphAPI instance with valid token
        """
        self.graph_api = graph_api
        self.common_skills = [
            'python', 'java', 'javascript', 'c++', 'sql', 'r',
            'machine learning', 'deep learning', 'data science', 
            'data analysis', 'statistics', 'nlp', 'computer vision',
            'tensorflow', 'pytorch', 'keras', 'scikit-learn',
            'pandas', 'numpy', 'matplotlib', 'seaborn',
            'streamlit', 'flask', 'django', 'fastapi',
            'docker', 'kubernetes', 'aws', 'azure', 'gcp',
            'powerbi', 'tableau', 'excel', 'git', 'github',
            'html', 'css', 'react', 'node.js', 'mongodb',
            'postgresql', 'mysql', 'spark', 'hadoop'
        ]
    
    def extract_text_from_pdf(self, file_content: bytes) -> str:
        """
        Extract text from PDF resume file
        
        Args:
            file_content: Binary content of PDF file
            
        Returns:
            str: Extracted text from PDF
        """
        try:
            pdf_reader = PyPDF2.PdfReader(io.BytesIO(file_content))
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            return text
        except Exception as e:
            raise Exception(f"Error extracting PDF text: {str(e)}")
    
    def extract_text_from_docx(self, file_content: bytes) -> str:
        """
        Extract text from DOCX resume file
        
        Args:
            file_content: Binary content of DOCX file
            
        Returns:
            str: Extracted text from DOCX
        """
        try:
            doc = docx.Document(io.BytesIO(file_content))
            text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
            return text
        except Exception as e:
            raise Exception(f"Error extracting DOCX text: {str(e)}")
    
    def extract_email(self, text: str) -> Optional[str]:
        """
        Extract email address from resume text
        
        Args:
            text: Resume text content
            
        Returns:
            str or None: First email address found
        """
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        emails = re.findall(email_pattern, text)
        return emails[0] if emails else None
    
    def extract_phone(self, text: str) -> Optional[str]:
        """
        Extract phone number from resume text
        
        Args:
            text: Resume text content
            
        Returns:
            str or None: First phone number found
        """
        # Pattern for various phone formats
        phone_patterns = [
            r'[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}',  # US/International
            r'[\+]?[0-9]{1,3}[-\s]?[0-9]{10}',  # Indian format
            r'\d{3}-\d{3}-\d{4}',  # XXX-XXX-XXXX
            r'\(\d{3}\)\s*\d{3}-\d{4}'  # (XXX) XXX-XXXX
        ]
        
        for pattern in phone_patterns:
            phones = re.findall(pattern, text)
            if phones:
                return phones[0]
        return None
    
    def extract_name(self, text: str) -> Optional[str]:
        """
        Extract candidate name from resume (usually first line)
        
        Args:
            text: Resume text content
            
        Returns:
            str or None: Candidate name
        """
        lines = text.strip().split('\n')
        for line in lines[:5]:  # Check first 5 lines
            line = line.strip()
            if line and len(line.split()) <= 4 and len(line) > 3:
                # Likely a name if it's 1-4 words and not too long
                if not re.search(r'\d', line):  # No digits in name
                    return line
        return None
    
    def extract_skills(self, text: str) -> List[str]:
        """
        Extract technical skills from resume text
        
        Args:
            text: Resume text content
            
        Returns:
            list: List of matching skills found
        """
        text_lower = text.lower()
        found_skills = []
        
        for skill in self.common_skills:
            if skill.lower() in text_lower:
                found_skills.append(skill)
        
        return list(set(found_skills))  # Remove duplicates
    
    def extract_experience_years(self, text: str) -> Optional[str]:
        """
        Extract years of experience from resume
        
        Args:
            text: Resume text content
            
        Returns:
            str or None: Years of experience
        """
        # Patterns for experience
        exp_patterns = [
            r'(\d+)\+?\s*(?:years?|yrs?)\s*(?:of)?\s*experience',
            r'experience\s*:?\s*(\d+)\+?\s*(?:years?|yrs?)',
            r'(\d+)\+?\s*(?:years?|yrs?)\s*in'
        ]
        
        for pattern in exp_patterns:
            matches = re.findall(pattern, text.lower())
            if matches:
                return f"{matches[0]}+ years"
        return None
    
    def extract_education(self, text: str) -> List[str]:
        """
        Extract education qualifications from resume
        
        Args:
            text: Resume text content
            
        Returns:
            list: List of found degrees/certifications
        """
        degrees = [
            'b.tech', 'btech', 'b.e', 'be', 'bachelor',
            'm.tech', 'mtech', 'm.e', 'me', 'master', 'mba', 'mca',
            'phd', 'ph.d', 'doctorate',
            'diploma', 'certification', 'certified',
            'b.sc', 'bsc', 'm.sc', 'msc', 'b.a', 'ba', 'm.a', 'ma'
        ]
        
        text_lower = text.lower()
        found_degrees = []
        
        for degree in degrees:
            if degree in text_lower:
                found_degrees.append(degree.upper())
        
        return list(set(found_degrees))  # Remove duplicates
    
    def extract_resume_info(self, text: str) -> Dict:
        """
        Extract all structured information from resume text
        
        Args:
            text: Resume text content
            
        Returns:
            dict: Structured resume data
        """
        return {
            'name': self.extract_name(text),
            'email': self.extract_email(text),
            'phone': self.extract_phone(text),
            'skills': self.extract_skills(text),
            'experience': self.extract_experience_years(text),
            'education': self.extract_education(text),
            'raw_text_preview': text[:500] + '...' if len(text) > 500 else text
        }
    
    def parse_resume_from_onedrive(self, file_id: str, file_name: str) -> Dict:
        """
        Download and parse a resume file from OneDrive
        
        Args:
            file_id: OneDrive file ID
            file_name: Name of the file (to determine extension)
            
        Returns:
            dict: Parsed resume data
        """
        file_ext = os.path.splitext(file_name)[1].lower()
        
        if file_ext not in ['.pdf', '.docx', '.doc']:
            raise ValueError(f"Unsupported file type: {file_ext}. Only PDF and DOCX supported.")
        
        # Download file content from OneDrive
        print(f"Downloading {file_name} from OneDrive...")
        file_content = self.graph_api.get_file_content(file_id)
        
        # Extract text based on file type
        print(f"Extracting text from {file_ext} file...")
        if file_ext == '.pdf':
            text = self.extract_text_from_pdf(file_content)
        else:  # .docx or .doc
            text = self.extract_text_from_docx(file_content)
        
        # Parse structured information
        print("Parsing resume information...")
        resume_data = self.extract_resume_info(text)
        resume_data['file_name'] = file_name
        resume_data['file_type'] = file_ext
        
        return resume_data
    
    def search_and_parse_all_resumes(self, search_query: str = "resume") -> List[Dict]:
        """
        Search OneDrive for resume files and parse all of them
        
        Args:
            search_query: Search term for finding resumes (default: "resume")
            
        Returns:
            list: List of parsed resume data dictionaries
        """
        print(f"Searching OneDrive for '{search_query}'...")
        search_results = self.graph_api.search_onedrive(search_query)
        
        parsed_resumes = []
        files_found = search_results.get('value', [])
        
        print(f"Found {len(files_found)} files matching '{search_query}'")
        
        for file_item in files_found:
            file_name = file_item.get('name', '')
            file_id = file_item.get('id')
            file_ext = os.path.splitext(file_name)[1].lower()
            
            # Only process PDF and DOCX files
            if file_ext in ['.pdf', '.docx', '.doc']:
                try:
                    print(f"\nProcessing: {file_name}")
                    resume_data = self.parse_resume_from_onedrive(file_id, file_name)
                    parsed_resumes.append(resume_data)
                    print(f"✓ Successfully parsed {file_name}")
                except Exception as e:
                    print(f"✗ Error parsing {file_name}: {str(e)}")
                    parsed_resumes.append({
                        'file_name': file_name,
                        'error': str(e)
                    })
        
        print(f"\nCompleted! Successfully parsed {len([r for r in parsed_resumes if 'error' not in r])} resumes.")
        return parsed_resumes
    
    def get_resume_summary(self, resume_data: Dict) -> str:
        """
        Generate a human-readable summary of parsed resume
        
        Args:
            resume_data: Dictionary with parsed resume information
            
        Returns:
            str: Formatted summary string
        """
        summary = f"""\n=== RESUME SUMMARY ===
File: {resume_data.get('file_name', 'Unknown')}
Name: {resume_data.get('name', 'Not found')}
Email: {resume_data.get('email', 'Not found')}
Phone: {resume_data.get('phone', 'Not found')}
Experience: {resume_data.get('experience', 'Not specified')}
Education: {', '.join(resume_data.get('education', [])) if resume_data.get('education') else 'Not found'}
Skills ({len(resume_data.get('skills', []))} found): {', '.join(resume_data.get('skills', [])) if resume_data.get('skills') else 'None'}
=====================\n"""
        return summary
