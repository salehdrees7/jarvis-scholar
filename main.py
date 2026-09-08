from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse
from pypdf import PdfReader
import httpx
import io
import os

app = FastAPI(
    title="AI Document Platform",
    description="Upload PDF documents and chat with JARVIS about their contents.",
    version="3.0.0",
)

JARVIS_API_URL = os.getenv(
    "JARVIS_API_URL",
    "http://127.0.0.1:8000",
).rstrip("/")

uploaded_document_text = ""
uploaded_document_name = ""
chat_history = []

HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>JARVIS Document AI</title>
    <style>
        :root {
            --bg: #07111d;
            --panel: #0d1b2a;
            --panel2: #0a1724;
            --border: #1f3b52;
            --text: #f3f7fb;
            --muted: #98adbf;
            --accent: #57c7ff;
            --accent2: #199bd7;
            --success: #62d8a0;
            --user: #15354b;
            --assistant: #0c2232;
        }

        * { box-sizing: border-box; }

        body {
            margin: 0;
            min-height: 100vh;
            background:
                radial-gradient(circle at top right, rgba(87,199,255,.12), transparent 32%),
                linear-gradient(180deg, #06101a 0%, #08131f 100%);
            color: var(--text);
            font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        }

        .shell {
            width: min(1180px, calc(100% - 30px));
            margin: 0 auto;
            padding: 28px 0 50px;
        }

        .topbar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 20px;
            margin-bottom: 22px;
        }

        .brand {
            display: flex;
            gap: 13px;
            align-items: center;
        }

        .orb {
            width: 44px;
            height: 44px;
            border-radius: 50%;
            background: radial-gradient(circle at 38% 34%, #e1f8ff 0 6%, #5ad2ff 12%, #087db8 42%, #031d2f 74%);
            box-shadow: 0 0 24px rgba(80,202,255,.45);
        }

        h1 {
            margin: 0;
            font-size: 22px;
            letter-spacing: .15em;
        }

        .subtitle {
            color: var(--muted);
            font-size: 13px;
            margin-top: 4px;
        }

        .status {
            color: var(--muted);
            font-size: 13px;
            display: flex;
            gap: 8px;
            align-items: center;
        }

        .dot {
            width: 9px;
            height: 9px;
            border-radius: 50%;
            background: #667583;
        }

        .dot.online {
            background: var(--success);
            box-shadow: 0 0 10px rgba(98,216,160,.7);
        }

        .layout {
            display: grid;
            grid-template-columns: 330px 1fr;
            gap: 20px;
            align-items: stretch;
        }

        .card {
            background: rgba(13,27,42,.92);
            border: 1px solid var(--border);
            border-radius: 18px;
            box-shadow: 0 18px 45px rgba(0,0,0,.18);
        }

        .sidebar {
            padding: 20px;
            min-height: 680px;
        }

        .sidebar h2, .chat-head h2 {
            margin: 0 0 8px;
            font-size: 17px;
        }

        .sidebar p, .chat-head p {
            margin: 0;
            color: var(--muted);
            font-size: 13px;
            line-height: 1.5;
        }

        .dropzone {
            margin-top: 18px;
            background: var(--panel2);
            border: 1.5px dashed #31566f;
            border-radius: 14px;
            padding: 22px 14px;
            text-align: center;
        }

        .dropzone strong {
            display: block;
            margin-bottom: 8px;
        }

        input[type="file"] {
            width: 100%;
            margin-top: 14px;
            color: var(--muted);
        }

        button {
            border: 0;
            border-radius: 11px;
            padding: 11px 15px;
            font: inherit;
            font-weight: 750;
            cursor: pointer;
        }

        button:disabled {
            opacity: .5;
            cursor: not-allowed;
        }

        .primary {
            background: linear-gradient(135deg, var(--accent), var(--accent2));
            color: #03131e;
        }

        .secondary {
            background: #14283a;
            color: var(--text);
            border: 1px solid var(--border);
        }

        .fullwidth {
            width: 100%;
            margin-top: 12px;
        }

        .docinfo {
            margin-top: 14px;
            padding: 12px;
            border-radius: 11px;
            border: 1px solid #17334a;
            background: #091724;
            color: var(--muted);
            font-size: 13px;
            white-space: pre-wrap;
        }

        .quick {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
            margin-top: 18px;
        }

        .quick button {
            background: #10263a;
            color: #cdefff;
            border: 1px solid #224761;
            font-size: 12px;
            padding: 9px 10px;
        }

        .chat {
            display: flex;
            flex-direction: column;
            min-height: 680px;
            overflow: hidden;
        }

        .chat-head {
            padding: 18px 20px;
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
        }

        .chat-actions {
            display: flex;
            gap: 8px;
        }

        .messages {
            flex: 1;
            padding: 22px;
            overflow-y: auto;
            max-height: 540px;
        }

        .empty {
            color: var(--muted);
            text-align: center;
            padding: 90px 20px;
            line-height: 1.6;
        }

        .msg {
            max-width: 82%;
            margin-bottom: 14px;
            padding: 13px 15px;
            border-radius: 14px;
            line-height: 1.6;
            white-space: pre-wrap;
        }

        .msg.user {
            margin-left: auto;
            background: var(--user);
            border: 1px solid #24516d;
        }

        .msg.assistant {
            margin-right: auto;
            background: var(--assistant);
            border: 1px solid #18384e;
        }

        .role {
            display: block;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: .08em;
            color: var(--muted);
            margin-bottom: 5px;
        }

        .composer {
            padding: 16px;
            border-top: 1px solid var(--border);
            background: #0a1724;
        }

        textarea {
            width: 100%;
            min-height: 84px;
            max-height: 180px;
            resize: vertical;
            border-radius: 13px;
            border: 1px solid var(--border);
            background: #07131f;
            color: var(--text);
            padding: 13px 14px;
            font: inherit;
            outline: none;
        }

        textarea:focus {
            border-color: var(--accent2);
            box-shadow: 0 0 0 3px rgba(87,199,255,.08);
        }

        .composer-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 10px;
            margin-top: 10px;
        }

        .hint {
            color: var(--muted);
            font-size: 12px;
        }

        @media (max-width: 850px) {
            .layout {
                grid-template-columns: 1fr;
            }

            .sidebar {
                min-height: auto;
            }
        }
    </style>
</head>
<body>
    <main class="shell">
        <div class="topbar">
            <div class="brand">
                <div class="orb"></div>
                <div>
                    <h1>JARVIS</h1>
                    <div class="subtitle">Document Intelligence Platform</div>
                </div>
            </div>

            <div class="status">
                <span id="statusDot" class="dot"></span>
                <span id="statusText">Checking JARVIS...</span>
            </div>
        </div>

        <div class="layout">
            <aside class="card sidebar">
                <h2>Document</h2>
                <p>Upload one PDF, then chat with JARVIS about it.</p>

                <div class="dropzone">
                    <strong>Select a PDF</strong>
                    <span style="color:var(--muted);font-size:12px;">PDF files only</span>
                    <input id="pdfFile" type="file" accept=".pdf,application/pdf">
                </div>

                <button id="uploadBtn" class="primary fullwidth" onclick="uploadPDF()">Upload document</button>

                <div id="docInfo" class="docinfo">No document uploaded.</div>

                <div class="quick">
                    <button onclick="quickAsk('Summarise this document clearly.')">Summarise</button>
                    <button onclick="quickAsk('Explain the main points in simple terms.')">Explain simply</button>
                    <button onclick="quickAsk('What are the most important points in this document?')">Key points</button>
                    <button onclick="quickAsk('Create a concise study guide from this document.')">Study guide</button>
                </div>
            </aside>

            <section class="card chat">
                <div class="chat-head">
                    <div>
                        <h2>Chat with JARVIS</h2>
                        <p>Ask follow-up questions without re-uploading the PDF.</p>
                    </div>
                    <div class="chat-actions">
                        <button class="secondary" onclick="clearChat()">Clear chat</button>
                    </div>
                </div>

                <div id="messages" class="messages">
                    <div class="empty" id="emptyState">
                        Upload a PDF on the left, then ask JARVIS anything about it.
                    </div>
                </div>

                <div class="composer">
                    <textarea id="question" placeholder="Ask something about the document..."></textarea>
                    <div class="composer-row">
                        <span class="hint">Follow-up questions use the same uploaded document.</span>
                        <button id="askBtn" class="primary" onclick="askJarvis()">Ask JARVIS</button>
                    </div>
                </div>
            </section>
        </div>
    </main>

    <script>
        async function checkJarvis() {
            const dot = document.getElementById("statusDot");
            const text = document.getElementById("statusText");

            try {
                const res = await fetch("/jarvis-status");
                const data = await res.json();

                if (data.connected) {
                    dot.classList.add("online");
                    const brain = data.jarvis?.brain || "Local AI";
                    text.textContent = `JARVIS online • ${brain}`;
                } else {
                    dot.classList.remove("online");
                    text.textContent = "JARVIS offline";
                }
            } catch {
                dot.classList.remove("online");
                text.textContent = "JARVIS offline";
            }
        }

        function addMessage(role, text) {
            const messages = document.getElementById("messages");
            const empty = document.getElementById("emptyState");

            if (empty) empty.remove();

            const div = document.createElement("div");
            div.className = `msg ${role}`;

            const label = document.createElement("span");
            label.className = "role";
            label.textContent = role === "user" ? "You" : "JARVIS";

            const content = document.createElement("div");
            content.textContent = text;

            div.appendChild(label);
            div.appendChild(content);
            messages.appendChild(div);
            messages.scrollTop = messages.scrollHeight;
        }

        async function uploadPDF() {
            const fileInput = document.getElementById("pdfFile");
            const info = document.getElementById("docInfo");
            const button = document.getElementById("uploadBtn");

            if (!fileInput.files.length) {
                info.textContent = "Choose a PDF first.";
                return;
            }

            const formData = new FormData();
            formData.append("file", fileInput.files[0]);

            button.disabled = true;
            info.textContent = "Uploading and extracting text...";

            try {
                const res = await fetch("/upload", {
                    method: "POST",
                    body: formData
                });

                const data = await res.json();

                if (data.error) {
                    info.textContent = `Error: ${data.error}`;
                } else {
                    info.textContent =
                        `Ready: ${data.filename}\n${data.pages} pages • ${data.characters_extracted.toLocaleString()} characters`;
                    await clearChat(false);
                }
            } catch (e) {
                info.textContent = `Upload failed: ${e.message}`;
            } finally {
                button.disabled = false;
            }
        }

        async function quickAsk(text) {
            document.getElementById("question").value = text;
            await askJarvis();
        }

        async function askJarvis() {
            const questionBox = document.getElementById("question");
            const question = questionBox.value.trim();
            const button = document.getElementById("askBtn");

            if (!question) return;

            addMessage("user", question);
            questionBox.value = "";
            button.disabled = true;

            const thinkingId = "thinking-" + Date.now();
            const messages = document.getElementById("messages");

            const thinking = document.createElement("div");
            thinking.className = "msg assistant";
            thinking.id = thinkingId;
            thinking.innerHTML = '<span class="role">JARVIS</span><div>Reading the document...</div>';
            messages.appendChild(thinking);
            messages.scrollTop = messages.scrollHeight;

            const formData = new FormData();
            formData.append("question", question);

            try {
                const res = await fetch("/ask", {
                    method: "POST",
                    body: formData
                });

                const data = await res.json();

                document.getElementById(thinkingId)?.remove();

                if (data.error) {
                    addMessage("assistant", `Error: ${data.error}`);
                } else {
                    addMessage("assistant", data.answer || "JARVIS returned an empty response.");
                }
            } catch (e) {
                document.getElementById(thinkingId)?.remove();
                addMessage("assistant", `Request failed: ${e.message}`);
            } finally {
                button.disabled = false;
            }
        }

        async function clearChat(showMessage = true) {
            try {
                await fetch("/clear-chat", { method: "POST" });
            } catch {}

            const messages = document.getElementById("messages");
            messages.innerHTML = '<div class="empty" id="emptyState">Upload a PDF on the left, then ask JARVIS anything about it.</div>';

            if (showMessage) {
                document.getElementById("question").value = "";
            }
        }

        document.getElementById("question").addEventListener("keydown", function(event) {
            if (event.key === "Enter" && !event.shiftKey) {
                event.preventDefault();
                askJarvis();
            }
        });

        checkJarvis();
        setInterval(checkJarvis, 10000);
    </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def website():
    return HTMLResponse(HTML)


@app.get("/jarvis-status")
async def jarvis_status():
    try:
        async with httpx.AsyncClient(timeout=10.0) as client:
            response = await client.get(f"{JARVIS_API_URL}/health")
            response.raise_for_status()

        return {
            "connected": True,
            "jarvis": response.json(),
        }

    except Exception as error:
        return {
            "connected": False,
            "error": str(error),
        }


@app.post("/upload")
async def upload_pdf(file: UploadFile = File(...)):
    global uploaded_document_text
    global uploaded_document_name
    global chat_history

    if not file.filename or not file.filename.lower().endswith(".pdf"):
        return {"error": "Please upload a PDF file."}

    contents = await file.read()

    try:
        reader = PdfReader(io.BytesIO(contents))
    except Exception:
        return {"error": "The PDF could not be read."}

    text_parts = []

    for page_number, page in enumerate(reader.pages, start=1):
        page_text = page.extract_text()
        if page_text:
            text_parts.append(f"\n--- PAGE {page_number} ---\n{page_text}")

    uploaded_document_text = "\n".join(text_parts)
    uploaded_document_name = file.filename
    chat_history = []

    if not uploaded_document_text.strip():
        return {"error": "No readable text could be extracted from this PDF."}

    return {
        "filename": file.filename,
        "pages": len(reader.pages),
        "characters_extracted": len(uploaded_document_text),
        "message": "PDF uploaded successfully and is ready for JARVIS.",
    }


@app.post("/ask")
async def ask_question(question: str = Form(...)):
    global chat_history

    if not uploaded_document_text:
        return {"error": "Upload a PDF before asking a question."}

    if not question.strip():
        return {"error": "Please enter a question."}

    recent_history = chat_history[-8:]

    history_text = "\n".join(
        f"{item['role'].upper()}: {item['content']}"
        for item in recent_history
    )

    combined_question = question.strip()

    if history_text:
        combined_question = (
            "Use the previous document conversation below only to understand follow-up references.\n\n"
            f"{history_text}\n\n"
            f"NEW USER QUESTION:\n{question.strip()}"
        )

    payload = {
        "document_name": uploaded_document_name,
        "context": uploaded_document_text,
        "question": combined_question,
    }

    try:
        async with httpx.AsyncClient(timeout=120.0) as client:
            response = await client.post(
                f"{JARVIS_API_URL}/api/document-query",
                json=payload,
            )

        if response.status_code != 200:
            try:
                detail = response.json()
            except Exception:
                detail = response.text

            return {
                "error": "JARVIS returned an error.",
                "status_code": response.status_code,
                "detail": detail,
            }

        result = response.json()
        answer = result.get("answer", "")

        chat_history.append({
            "role": "user",
            "content": question.strip(),
        })
        chat_history.append({
            "role": "assistant",
            "content": answer,
        })

        return {
            "document": uploaded_document_name,
            "question": question.strip(),
            "answer": answer,
            "ai_backend": "JARVIS",
            "brain": result.get("brain"),
            "context_truncated": result.get("context_truncated", False),
        }

    except httpx.ConnectError:
        return {
            "error": "Could not connect to JARVIS. Make sure JARVIS is running on port 8000."
        }

    except Exception as error:
        return {"error": f"JARVIS request failed: {str(error)}"}


@app.post("/clear-chat")
def clear_chat():
    global chat_history
    chat_history = []

    return {
        "message": "Chat history cleared."
    }


@app.get("/document")
def document_info():
    if not uploaded_document_text:
        return {"message": "No document has been uploaded yet."}

    return {
        "filename": uploaded_document_name,
        "characters": len(uploaded_document_text),
        "chat_messages": len(chat_history),
        "status": "Document loaded and ready for JARVIS questions.",
    }
