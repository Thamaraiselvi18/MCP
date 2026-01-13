# Google Workspace – MCP Server

This project demonstrates how to build an **MCP (Model Context Protocol) server** that allows AI agents to interact with **Google Sheets** programmatically. It shows authentication, reading/writing data, and basic automation using the Google Sheets API.

The repository is intended for learning and experimentation with MCP servers and Google APIs.

---

## 📁 Project Structure

```
sheet_api_demo/
│
├── MCP server/
│   ├── app.py                 # MCP server entry / API logic
│   ├── main.py                # Main execution file
│   ├── cred.json              # Service account or OAuth credentials (example)      
│   ├── token.pickle           # Generated access token after authentication
│   ├── temp_test.txt          # Temporary test file
│   ├── temp_meeting_notes.txt # Sample input text
│   └── prac.ipynb              # Practice / experimentation notebook
│
├── evn/
│   └── .env                   # Environment variables
│
├── __pycache__/               # Python cache files
├── .vscode/                   # VS Code configuration
└── README.md
```

---

## 🚀 Features

* MCP server implementation in Python
* Google Sheets API integration
* OAuth / credential-based authentication
* Read and write spreadsheet data using AI agents
* Modular and extensible server structure

---

## 🛠️ Requirements

* Python **3.10+**
* Google Cloud Project with **Sheets API enabled**
* MCP-compatible client / agent

### Python Libraries

Install required packages:

```bash
pip install google-api-python-client google-auth google-auth-oauthlib python-dotenv
```

---

## 🔐 Google Sheets API Setup

1. Go to **Google Cloud Console**
2. Create a new project
3. Enable **Google Sheets API**
4. Create OAuth credentials
5. Download `credentials.json`
6. Place it inside:

   ```
   sheet_api_demo/MCP server/
   ```

On first run, `token.pickle` will be generated automatically after authentication.

---

## ⚙️ Environment Variables

Update the `.env` file inside `evn/`:

```
GOOGLE_APPLICATION_CREDENTIALS=credentials.json
```

(Modify if your setup differs.)

---

## ▶️ How to Run

Navigate to the MCP server directory:

```bash
cd sheet_api_demo/MCP\ server
```

Run the server:

```bash
python main.py
```

or

```bash
python app.py
```

(depending on your implementation entry point)

---

## 🧪 Testing

* Use `prac.ipynb` for testing Sheets API calls
* Use `temp_test.txt` or `temp_meeting_notes.txt` as sample input data
* Verify spreadsheet updates in your Google Sheets account

---

## 📌 Notes

* Do **not** commit real credentials to public repositories
* Add `credentials.json` and `token.pickle` to `.gitignore`
* This project is for **educational/demo purposes**

---

## 📚 Use Cases

* AI agents updating spreadsheets automatically
* Meeting notes → structured Google Sheets
* MCP-based automation workflows
* Learning Google API + MCP integration

---

## 👩‍💻 Author

Developed as an MCP + Google Sheets API demo project.

---

## 📄 License

This project is open for learning and experimentation.
You may adapt it for personal or academic use.
