# Quick start

## Prerequisites

- Docker + Compose on a LAN host  
- OmniRoute (or any OpenAI-compatible `/v1`) reachable from the WebUI container  
- Enough RAM for WebUI embeddings + your router (8GB+ free recommended)

## 1) Configure env

```bash
cp deploy/.env.example deploy/.env
# set OPENAI_API_BASE_URL and OPENAI_API_KEY (OmniRoute client key)
```

## 2) Start Open WebUI

```bash
cd deploy
docker compose -f docker-compose.open-webui.example.yml up -d
```

Bind only to your LAN IP in compose (example uses a placeholder).

## 3) First admin

Open the UI → first signup becomes admin → create named users → **disable signup**.

## 4) Install tools

In Admin / Workspace → **Tools**, add each file from `deploy/tools/`:

- `generate_forecast_pack.py`
- `generate_pdf.py`
- `generate_docx.py`
- `generate_excel.py`
- (optional) slides tool if you maintain `generate_slides.py` on the host

Install filter from `deploy/functions/follow_up_gate.py` as a global **Filter**.

Inside the container (after image recreate):

```bash
docker exec -it open-webui pip install python-pptx pillow reportlab python-docx openpyxl ddgs
```

## 5) Grant tool access to users

Non-admin users need **read** `access_grant` on each tool (UI share, or DB grant).  
Without this they only see built-ins (Web Search / Code Interpreter).

## 6) Web search

Confirm compose has `ENABLE_WEB_SEARCH=true` and `WEB_SEARCH_ENGINE=duckduckgo`.  
In chat: Tools grid → Web Search ON.

## 7) Feel-test

1. New Chat → enable needed tools  
2. Ask a web question **or** request a Generate* deliverable  
3. Download via `/api/v1/files/{id}/content` style link  

Keep **Interface → Follow-Up Auto-Generation** OFF.
