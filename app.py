# sheets_slides_mcp_server.py
# MCP Server for Google Sheets + Slides + Gemini AI Automation
# Run with: python sheets_slides_mcp_server.py

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import pickle
import os.path
import json
import google.generativeai as genai
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("Google Workspace + Gemini Automation")

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/presentations"
]

CREDENTIALS_FILE = "cred.json"
TOKEN_FILE = "token.pickle"

# Configure Gemini API (FREE)
# Get your free API key from: https://makersuite.google.com/app/apikey
GEMINI_API_KEY = "AIzaSyAV4qG5LKlgxqGElMGCnNSPvKpSCoqdkkc"  # Replace with your key
genai.configure(api_key=GEMINI_API_KEY)

def get_services():
    """Authenticate and return Google API services"""
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, "wb") as token:
            pickle.dump(creds, token)
    
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    slides_service = build("slides", "v1", credentials=creds)
    return drive_service, sheets_service, slides_service

drive_service, sheets_service, slides_service = get_services()

# ============== HELPER FUNCTIONS ==============

def find_or_create_folder(folder_name: str) -> str:
    """Find existing folder or create new one, return folder_id"""
    results = drive_service.files().list(
        q=f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id)"
    ).execute()
    
    if results.get('files'):
        return results['files'][0]['id']
    
    folder = drive_service.files().create(body={
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder"
    }, fields="id").execute()
    return folder["id"]

# ============== GEMINI AI INTEGRATION ==============

@mcp.tool()
def generate_slides_with_gemini(
    topic: str,
    num_slides: int = 10,
    folder_name: str = "AI Generated Presentations",
    presentation_name: str = None
) -> str:
    """
    Uses Gemini AI (FREE) to generate presentation slides on any topic.
    
    Args:
        topic: Topic for the presentation (e.g., "artificial intelligence and data science")
        num_slides: Number of slides to generate (default: 10)
        folder_name: Folder to save presentation
        presentation_name: Name of presentation (auto-generated if not provided)
    
    Returns:
        Presentation ID and URL
    
    Example:
        generate_slides_with_gemini("artificial intelligence and data science", 10)
    """
    try:
        # Initialize Gemini model (gemini-pro is FREE)
        model = genai.GenerativeModel('gemini-pro')
        
        # Create prompt for Gemini
        prompt = f"""Create a professional presentation on "{topic}" with exactly {num_slides} slides.

For each slide, provide:
1. A clear, concise title (max 10 words)
2. Body content with 3-5 bullet points (each point should be informative but brief)

Format your response as a JSON array like this:
[
  {{
    "title": "Introduction to AI",
    "body": "• Definition of Artificial Intelligence\\n• Historical development\\n• Current applications\\n• Future prospects"
  }},
  {{
    "title": "Machine Learning Fundamentals",
    "body": "• Supervised learning\\n• Unsupervised learning\\n• Reinforcement learning\\n• Deep learning basics"
  }}
]

Make the content educational, well-structured, and suitable for a professional presentation.
Return ONLY the JSON array, no additional text."""

        # Generate content using Gemini
        print("🤖 Generating slides using Gemini AI...")
        response = model.generate_content(prompt)
        
        # Parse the response
        response_text = response.text.strip()
        
        # Extract JSON from response (handle markdown code blocks)
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0].strip()
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0].strip()
        
        slides_data = json.loads(response_text)
        
        # Ensure we have the right number of slides
        slides_data = slides_data[:num_slides]
        
        # Auto-generate presentation name if not provided
        if not presentation_name:
            presentation_name = f"{topic.title()} - AI Generated"
        
        # Create the presentation
        print(f"📊 Creating presentation with {len(slides_data)} slides...")
        result = create_google_slides_from_data(folder_name, presentation_name, slides_data)
        
        return result
    
    except json.JSONDecodeError as e:
        return f"❌ Error parsing Gemini response: {str(e)}\n\nRaw response:\n{response_text[:500]}"
    except Exception as e:
        return f"❌ Error generating slides: {str(e)}"

@mcp.tool()
def ai_summarize_to_slides(
    text_content: str,
    num_slides: int = 5,
    folder_name: str = "AI Summaries",
    presentation_name: str = "AI Summary"
) -> str:
    """
    Uses Gemini AI to summarize long text into presentation slides.
    
    Args:
        text_content: Long text to summarize
        num_slides: Number of slides to create
        folder_name: Folder to save presentation
        presentation_name: Name of presentation
    
    Returns:
        Presentation ID and URL
    """
    try:
        model = genai.GenerativeModel('gemini-pro')
        
        prompt = f"""Summarize the following text into {num_slides} presentation slides.

Text to summarize:
{text_content[:5000]}

Create slides with:
- Clear titles
- 3-5 bullet points per slide
- Key insights and takeaways

Format as JSON array: [{{"title": "...", "body": "..."}}]
Return ONLY the JSON array."""

        response = model.generate_content(prompt)
        response_text = response.text.strip()
        
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0].strip()
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0].strip()
        
        slides_data = json.loads(response_text)
        
        return create_google_slides_from_data(folder_name, presentation_name, slides_data)
    
    except Exception as e:
        return f"❌ Error summarizing text: {str(e)}"

# ============== SHEETS TOOLS ==============

@mcp.tool()
def create_google_sheet(
    folder_name: str,
    sheet_name: str,
    headers: list[str],
    rows: list[list[str]]
) -> str:
    """
    Creates a Google Sheet with headers and data rows in a specified folder.
    
    Args:
        folder_name: Name of folder to create/use
        sheet_name: Name of the spreadsheet
        headers: List of column headers
        rows: List of rows, each row is a list of values
    
    Returns:
        URL of the created spreadsheet
    """
    try:
        folder_id = find_or_create_folder(folder_name)
        
        sheet = drive_service.files().create(body={
            "name": sheet_name,
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents": [folder_id]
        }, fields="id, webViewLink").execute()
        
        sheet_id = sheet["id"]
        sheet_url = sheet["webViewLink"]
        
        all_data = [headers] + rows
        sheets_service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range="Sheet1!A1",
            valueInputOption="RAW",
            body={"values": all_data}
        ).execute()
        
        return f"✅ Sheet created successfully!\n📊 URL: {sheet_url}\n🆔 ID: {sheet_id}"
    
    except Exception as e:
        return f"❌ Error creating sheet: {str(e)}"

@mcp.tool()
def read_google_sheet(spreadsheet_id: str, range_name: str = "Sheet1") -> dict:
    """
    Reads data from a Google Sheet.
    
    Args:
        spreadsheet_id: The ID of the spreadsheet (from URL)
        range_name: Range to read (default: "Sheet1")
    
    Returns:
        Dictionary with headers and rows
    """
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            return {"error": "No data found"}
        
        return {
            "headers": values[0] if values else [],
            "rows": values[1:] if len(values) > 1 else [],
            "total_rows": len(values) - 1
        }
    
    except Exception as e:
        return {"error": f"❌ Error reading sheet: {str(e)}"}

@mcp.tool()
def append_to_sheet(
    spreadsheet_id: str,
    rows: list[list[str]],
    range_name: str = "Sheet1"
) -> str:
    """
    Appends new rows to an existing Google Sheet.
    
    Args:
        spreadsheet_id: The ID of the spreadsheet
        rows: List of rows to append
        range_name: Sheet name to append to
    
    Returns:
        Success message
    """
    try:
        sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption="RAW",
            body={"values": rows}
        ).execute()
        
        return f"✅ Successfully appended {len(rows)} rows to sheet"
    
    except Exception as e:
        return f"❌ Error appending to sheet: {str(e)}"

# ============== SLIDES TOOLS ==============

def create_google_slides_from_data(
    folder_name: str,
    presentation_name: str,
    slides: list[dict]
) -> str:
    """
    Internal function to create Google Slides from slide data.
    
    Args:
        folder_name: Folder name
        presentation_name: Presentation name
        slides: List of dicts with 'title' and 'body'
    
    Returns:
        Success message with URL and ID
    """
    try:
        folder_id = find_or_create_folder(folder_name)
        
        presentation = slides_service.presentations().create(body={
            "title": presentation_name
        }).execute()
        
        presentation_id = presentation["presentationId"]
        
        # Move to folder
        drive_service.files().update(
            fileId=presentation_id,
            addParents=folder_id,
            fields="id, webViewLink"
        ).execute()
        
        # Get URL
        pres_file = drive_service.files().get(
            fileId=presentation_id,
            fields="webViewLink"
        ).execute()
        pres_url = pres_file["webViewLink"]
        
        # Get the first slide to delete later
        pres = slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()
        first_slide_id = pres['slides'][0]['objectId']
        
        # Create slides with content
        requests = []
        
        for idx, slide_data in enumerate(slides):
            slide_id = f"slide_{idx}"
            
            requests.append({
                "createSlide": {
                    "objectId": slide_id,
                    "slideLayoutReference": {"predefinedLayout": "TITLE_AND_BODY"}
                }
            })
        
        # Execute slide creation
        if requests:
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={"requests": requests}
            ).execute()
        
        # Add text to each slide
        pres = slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()
        
        text_requests = []
        for idx, slide_data in enumerate(slides):
            slide = pres['slides'][idx + 1]
            
            title_id = None
            body_id = None
            
            for element in slide.get('pageElements', []):
                shape = element.get('shape', {})
                placeholder = shape.get('placeholder', {})
                p_type = placeholder.get('type')
                
                if p_type == 'TITLE' or p_type == 'CENTERED_TITLE':
                    title_id = element['objectId']
                elif p_type == 'BODY' or p_type == 'SUBTITLE':
                    body_id = element['objectId']
            
            if title_id:
                text_requests.append({
                    "insertText": {
                        "objectId": title_id,
                        "text": slide_data.get("title", ""),
                        "insertionIndex": 0
                    }
                })
            
            if body_id:
                text_requests.append({
                    "insertText": {
                        "objectId": body_id,
                        "text": slide_data.get("body", ""),
                        "insertionIndex": 0
                    }
                })
        
        # Delete blank first slide
        text_requests.append({
            "deleteObject": {
                "objectId": first_slide_id
            }
        })
        
        if text_requests:
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={"requests": text_requests}
            ).execute()
        
        return f"""✅ Presentation created successfully!

📊 Title: {presentation_name}
🆔 Presentation ID: {presentation_id}
📄 Total Slides: {len(slides)}
🔗 URL: {pres_url}

You can now edit, share, or present this slide deck!"""
    
    except Exception as e:
        return f"❌ Error creating presentation: {str(e)}"

@mcp.tool()
def create_google_slides(
    folder_name: str,
    presentation_name: str,
    slides: list[dict]
) -> str:
    """
    Creates a Google Slides presentation with multiple slides.
    
    Args:
        folder_name: Name of folder to create/use
        presentation_name: Name of the presentation
        slides: List of slide dicts with 'title' and 'body' keys
                Example: [{"title": "Slide 1", "body": "Content here"}]
    
    Returns:
        URL and ID of the created presentation
    """
    return create_google_slides_from_data(folder_name, presentation_name, slides)

@mcp.tool()
def add_slide(
    presentation_id: str,
    title: str,
    body: str,
    position: int = -1
) -> str:
    """
    Adds a new slide to an existing presentation.
    
    Args:
        presentation_id: The ID of the presentation
        title: Title text for the slide
        body: Body text for the slide
        position: Position to insert (default: -1 for end)
    
    Returns:
        Success message
    """
    try:
        import time
        slide_id = f"new_slide_{int(time.time())}"
        
        requests = [{
            "createSlide": {
                "objectId": slide_id,
                "slideLayoutReference": {"predefinedLayout": "TITLE_AND_BODY"},
                "placeholderIdMappings": []
            }
        }]
        
        if position >= 0:
            requests[0]["createSlide"]["insertionIndex"] = position
        
        slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": requests}
        ).execute()
        
        pres = slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()
        
        created_slide = None
        for slide in pres['slides']:
            if slide['objectId'] == slide_id:
                created_slide = slide
                break
        
        if not created_slide:
            return "Slide created but couldn't add text"
        
        text_requests = []
        for element in created_slide.get('pageElements', []):
            shape = element.get('shape', {})
            placeholder = shape.get('placeholder', {})
            p_type = placeholder.get('type')
            
            if p_type in ['TITLE', 'CENTERED_TITLE']:
                text_requests.append({
                    "insertText": {
                        "objectId": element['objectId'],
                        "text": title,
                        "insertionIndex": 0
                    }
                })
            elif p_type in ['BODY', 'SUBTITLE']:
                text_requests.append({
                    "insertText": {
                        "objectId": element['objectId'],
                        "text": body,
                        "insertionIndex": 0
                    }
                })
        
        if text_requests:
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={"requests": text_requests}
            ).execute()
        
        return f"✅ Slide '{title}' added successfully"
    
    except Exception as e:
        return f"❌ Error adding slide: {str(e)}"

@mcp.tool()
def get_presentation_info(presentation_id: str) -> dict:
    """
    Gets information about a presentation including all slides.
    
    Args:
        presentation_id: The ID of the presentation
    
    Returns:
        Dictionary with presentation details
    """
    try:
        pres = slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()
        
        slides_info = []
        for slide in pres.get('slides', []):
            slide_data = {
                "slide_id": slide['objectId'],
                "elements": []
            }
            
            for element in slide.get('pageElements', []):
                if 'shape' in element:
                    shape = element['shape']
                    if 'text' in shape:
                        text_content = ""
                        for text_elem in shape['text'].get('textElements', []):
                            if 'textRun' in text_elem:
                                text_content += text_elem['textRun'].get('content', '')
                        
                        slide_data['elements'].append({
                            "type": shape.get('shapeType', 'UNKNOWN'),
                            "text": text_content.strip()
                        })
            
            slides_info.append(slide_data)
        
        return {
            "title": pres.get('title'),
            "presentation_id": pres.get('presentationId'),
            "total_slides": len(slides_info),
            "slides": slides_info
        }
    
    except Exception as e:
        return {"error": f"❌ Error getting presentation info: {str(e)}"}

if __name__ == "__main__":
    mcp.run(transport="sse")