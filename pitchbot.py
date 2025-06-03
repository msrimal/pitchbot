import os
import json
import re
from openai import OpenAI
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from dotenv import load_dotenv
from google.auth.transport.requests import Request  

# Load environment variables
load_dotenv()

# OpenAI client setup
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

SCOPES = ['https://www.googleapis.com/auth/presentations']
CREDENTIALS_FILE = 'creds.json'

def authenticate_google_slides():
    creds = None
    token_path = 'token.json'
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
    return build('slides', 'v1', credentials=creds)

def generate_pitch(user_data):
    prompt = build_prompt(user_data)
    print("Calling OpenAI API...")
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7,
        max_tokens=1000
    )
    return response.choices[0].message.content.strip()

def load_user_input(json_path="user_input.json"):
    with open(json_path, "r") as f:
        data = json.load(f)
    required_fields = ["problem", "solution-idea", "tone", "audience"]
    for field in required_fields:
        if field not in data or not data[field].strip():
            raise ValueError(f"Missing or empty field: {field}")
    return data

def build_prompt(user_data):
    return f"""
You are a startup pitch assistant.
Create a full startup pitch deck from these details:
- Problem: {user_data['problem']}
- Solution: {user_data['solution-idea']}
- Tone: {user_data['tone']}
- Audience: {user_data['audience']}
Output the following sections clearly labeled:
1. Problem Statement
2. Solution Overview
3. User Persona
4. Value Proposition
5. Business Model Summary
6. Tagline/Slogan
7. Elevator Pitch (1-minute speech)
"""

def create_presentation(service, title="Startup Pitch Deck"):
    presentation = service.presentations().create(body={"title": title}).execute()
    print(f"Created presentation with ID: {presentation['presentationId']}")
    return presentation['presentationId']

def create_section_slide(service, presentation_id, section_title, content):
    requests = [{
        "createSlide": {
            "slideLayoutReference": {
                "predefinedLayout": "TITLE_AND_BODY"
            }
        }
    }]
    response = service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": requests}
    ).execute()
    slide_id = response['replies'][0]['createSlide']['objectId']
    slide = service.presentations().pages().get(
        presentationId=presentation_id,
        pageObjectId=slide_id
    ).execute()
    title_id = None
    body_id = None
    for element in slide.get('pageElements', []):
        shape = element.get('shape')
        if not shape:
            continue
        placeholder = shape.get('placeholder')
        if not placeholder:
            continue
        if placeholder.get('type') == 'TITLE':
            title_id = element['objectId']
        elif placeholder.get('type') == 'BODY':
            body_id = element['objectId']
    requests = []
    if title_id:
        requests.append({
            "insertText": {
                "objectId": title_id,
                "insertionIndex": 0,
                "text": section_title
            }
        })
    if body_id:
        requests.append({
            "insertText": {
                "objectId": body_id,
                "insertionIndex": 0,
                "text": content
            }
        })
    service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": requests}
    ).execute()

def extract_sections(pitch_text):
    sections = {}
    pattern = r"(\d+)\.\s*(.+)"
    current_section = None
    lines = pitch_text.split('\n')
    buffer = []
    def save_buffer():
        if current_section:
            sections[current_section] = '\n'.join(buffer).strip()
    for line in lines:
        match = re.match(pattern, line)
        if match:
            save_buffer()
            buffer = []
            current_section = match.group(2).strip()
        else:
            if line.strip():
                buffer.append(line.strip())
    save_buffer()
    return sections

def main():
    print("\n--- PitchBot: AI Startup Pitch Generator ---\n")
    try:
        user_data = load_user_input()
        print("User input loaded.")
        print("\nGenerating pitch...")
        pitch = generate_pitch(user_data)
        print("\nGenerated pitch:\n")
        print(pitch)
        slides_service = authenticate_google_slides()
        presentation_id = create_presentation(slides_service, title="PitchBot Deck")
        sections = extract_sections(pitch)
        for section_title, content in sections.items():
            create_section_slide(slides_service, presentation_id, section_title, content)
        print(f"\nPresentation: https://docs.google.com/presentation/d/{presentation_id}/edit")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
