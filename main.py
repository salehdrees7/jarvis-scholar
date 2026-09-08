from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse
from pypdf import PdfReader
from docx import Document
from pptx import Presentation
import httpx, io, os

app = FastAPI(title="JARVIS Scholar", description="Cloud document intelligence by JARVIS Scholar.", version="5.0.0")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip()
uploaded_document_text = ""
uploaded_document_name = ""
uploaded_document_type = ""
chat_history = []

HTML = """<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>JARVIS Scholar</title>
<style>
:root{--bg:#07111d;--panel:#0d1b2a;--border:#1f3b52;--text:#f3f7fb;--muted:#98adbf;--accent:#57c7ff;--success:#62d8a0}
*{box-sizing:border-box}body{margin:0;min-height:100vh;background:#07111d;color:var(--text);font-family:system-ui,sans-serif}
.shell{width:min(1180px,calc(100% - 30px));margin:auto;padding:28px 0}.top{display:flex;justify-content:space-between;align-items:center;margin-bottom:22px}
.brand{display:flex;gap:13px;align-items:center}.orb{width:44px;height:44px;border-radius:50%;background:radial-gradient(circle at 38% 34%,#e1f8ff,#5ad2ff 15%,#087db8 45%,#031d2f 75%);box-shadow:0 0 24px #50caff66}
h1{margin:0;font-size:22px;letter-spacing:.15em}.muted,.status{color:var(--muted);font-size:13px}.dot{display:inline-block;width:9px;height:9px;border-radius:50%;background:#667583;margin-right:7px}.dot.online{background:var(--success)}
.layout{display:grid;grid-template-columns:330px 1fr;gap:20px}.card{background:var(--panel);border:1px solid var(--border);border-radius:18px}.side{padding:20px;min-height:680px}
.drop{margin-top:18px;padding:22px 14px;text-align:center;border:1.5px dashed #31566f;border-radius:14px;background:#0a1724}.drop strong{display:block;margin-bottom:8px}input[type=file]{width:100%;margin-top:14px;color:var(--muted)}
button{border:0;border-radius:11px;padding:11px 15px;font:inherit;font-weight:700;cursor:pointer}.primary{background:#57c7ff;color:#03131e}.secondary{background:#14283a;color:var(--text);border:1px solid var(--border)}.full{width:100%;margin-top:12px}
.info{margin-top:14px;padding:12px;border:1px solid #17334a;border-radius:11px;background:#091724;color:var(--muted);font-size:13px;white-space:pre-wrap}.quick{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:18px}.quick button{background:#10263a;color:#cdefff;border:1px solid #224761;font-size:12px}
.chat{display:flex;flex-direction:column;min-height:680px;overflow:hidden}.head{padding:18px 20px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between}.messages{flex:1;padding:22px;overflow:auto;max-height:540px}.empty{text-align:center;color:var(--muted);padding:90px 20px}
.msg{max-width:82%;margin-bottom:14px;padding:13px 15px;border-radius:14px;white-space:pre-wrap;line-height:1.6}.user{margin-left:auto;background:#15354b}.assistant{margin-right:auto;background:#0c2232}.role{display:block;font-size:11px;color:var(--muted);margin-bottom:5px}
.composer{padding:16px;border-top:1px solid var(--border);background:#0a1724}textarea{width:100%;min-height:84px;border-radius:13px;border:1px solid var(--border);background:#07131f;color:var(--text);padding:13px;font:inherit}.row{display:flex;justify-content:space-between;align-items:center;margin-top:10px;gap:10px}
@media(max-width:850px){.layout{grid-template-columns:1fr}.side{min-height:auto}}
</style></head><body><main class="shell">
<div class="top"><div class="brand"><div class="orb"></div><div><h1>JARVIS</h1><div class="muted">Scholar • Document Intelligence</div></div></div><div class="status"><span id="statusDot" class="dot"></span><span id="statusText">Checking JARVIS...</span></div></div>
<div class="layout"><aside class="card side"><h2>Document</h2><p class="muted">Upload a document, then chat with JARVIS about it.</p>
<div class="drop"><strong>Select a document</strong><span class="muted">PDF • Word • PowerPoint • TXT</span><input id="documentFile" type="file" accept=".pdf,.docx,.pptx,.txt"></div>
<button id="uploadBtn" class="primary full" onclick="uploadDocument()">Upload document</button><div id="docInfo" class="info">No document uploaded.</div>
<div class="quick"><button onclick="quickAsk('Summarise this document clearly.')">Summarise</button><button onclick="quickAsk('Explain the main points in simple terms.')">Explain simply</button><button onclick="quickAsk('What are the most important points in this document?')">Key points</button><button onclick="quickAsk('Create a concise study guide from this document.')">Study guide</button></div></aside>
<section class="card chat"><div class="head"><div><h2>Chat with JARVIS</h2><div class="muted">Ask follow-up questions without re-uploading the document.</div></div><button class="secondary" onclick="clearChat()">Clear chat</button></div>
<div id="messages" class="messages"><div class="empty" id="emptyState">Upload a document on the left, then ask JARVIS anything about it.</div></div>
<div style="padding:10px 16px;border-top:1px solid var(--border);display:flex;justify-content:space-between;align-items:center;gap:10px">
<span id="voiceStatus" class="muted">Voice ready</span>
<div style="display:flex;gap:7px">
<button class="secondary" onclick="pauseSpeech()">⏸ Pause</button>
<button class="secondary" onclick="resumeSpeech()">▶ Resume</button>
<button class="secondary" onclick="stopSpeech()">⏹ Stop</button>
</div>
</div>
<div class="composer"><textarea id="question" placeholder="Ask something about the document..."></textarea><div class="row"><span class="muted">Follow-up questions use the same uploaded document.</span><button id="askBtn" class="primary" onclick="askJarvis()">Ask JARVIS</button></div></div></section></div>
<footer style="text-align:center;color:#98adbf;font-size:12px;padding:24px 10px 4px">
<strong style="color:#f3f7fb">JARVIS Scholar</strong> — Built by <strong style="color:#cdefff">Saleh Pour</strong><br>
<span>Robotics &amp; Artificial Intelligence</span>
</footer></main>
<script>
let currentSpeech=null;
function chooseVoice(){
 const voices=window.speechSynthesis.getVoices();
 return voices.find(v=>v.lang.toLowerCase().startsWith("en-gb") &&
   ["ryan","george","arthur","daniel","oliver"].some(n=>v.name.toLowerCase().includes(n)))
   || voices.find(v=>v.lang.toLowerCase().startsWith("en-gb"))
   || voices.find(v=>v.lang.toLowerCase().startsWith("en"))
   || voices[0] || null;
}
function speakText(text){
 if(!("speechSynthesis" in window)){document.getElementById("voiceStatus").textContent="Speech unavailable";return;}
 window.speechSynthesis.cancel();
 currentSpeech=new SpeechSynthesisUtterance(text);
 const voice=chooseVoice();
 if(voice) currentSpeech.voice=voice;
 currentSpeech.lang=voice?.lang||"en-GB";
 currentSpeech.rate=.96; currentSpeech.pitch=.92; currentSpeech.volume=1;
 currentSpeech.onstart=()=>document.getElementById("voiceStatus").textContent="JARVIS is speaking...";
 currentSpeech.onend=()=>document.getElementById("voiceStatus").textContent="Voice ready";
 currentSpeech.onerror=()=>document.getElementById("voiceStatus").textContent="Voice stopped";
 window.speechSynthesis.speak(currentSpeech);
}
function pauseSpeech(){if(window.speechSynthesis.speaking&&!window.speechSynthesis.paused){window.speechSynthesis.pause();document.getElementById("voiceStatus").textContent="Voice paused";}}
function resumeSpeech(){if(window.speechSynthesis.paused){window.speechSynthesis.resume();document.getElementById("voiceStatus").textContent="JARVIS is speaking...";}}
function stopSpeech(){if("speechSynthesis" in window)window.speechSynthesis.cancel();currentSpeech=null;const s=document.getElementById("voiceStatus");if(s)s.textContent="Voice ready";}

async function checkJarvis(){let d=document.getElementById("statusDot"),t=document.getElementById("statusText");try{let r=await fetch("/jarvis-status"),x=await r.json();if(x.connected){d.classList.add("online");t.textContent="JARVIS online • "+(x.jarvis?.brain||"Local AI")}else{d.classList.remove("online");t.textContent="JARVIS offline"}}catch{d.classList.remove("online");t.textContent="JARVIS offline"}}
function addMessage(role,text){let m=document.getElementById("messages"),e=document.getElementById("emptyState");if(e)e.remove();let d=document.createElement("div");d.className="msg "+role;let l=document.createElement("span");l.className="role";l.textContent=role==="user"?"You":"JARVIS";let c=document.createElement("div");c.textContent=text;d.append(l,c);
if(role==="assistant"){
 let tools=document.createElement("div");
 tools.style.marginTop="10px";
 let read=document.createElement("button");
 read.textContent="🔊 Read aloud";
 read.className="secondary";
 read.style.padding="6px 9px";
 read.style.fontSize="11px";
 read.onclick=()=>speakText(text);
 tools.appendChild(read);
 d.appendChild(tools);
}
m.appendChild(d);m.scrollTop=m.scrollHeight}
async function uploadDocument(){let f=document.getElementById("documentFile"),i=document.getElementById("docInfo"),b=document.getElementById("uploadBtn");if(!f.files.length){i.textContent="Choose a document first.";return}let fd=new FormData();fd.append("file",f.files[0]);b.disabled=true;i.textContent="Uploading and extracting text...";try{let r=await fetch("/upload",{method:"POST",body:fd}),x=await r.json();i.textContent=x.error?"Error: "+x.error:`Ready: ${x.filename}\n${x.document_type} • ${x.structure} • ${x.characters_extracted.toLocaleString()} characters`;if(!x.error)await clearChat(false)}catch(e){i.textContent="Upload failed: "+e.message}finally{b.disabled=false}}
async function quickAsk(q){document.getElementById("question").value=q;await askJarvis()}
async function askJarvis(){let q=document.getElementById("question"),text=q.value.trim(),b=document.getElementById("askBtn");if(!text)return;addMessage("user",text);q.value="";b.disabled=true;let fd=new FormData();fd.append("question",text);try{let r=await fetch("/ask",{method:"POST",body:fd}),x=await r.json();addMessage("assistant",x.error?"Error: "+x.error:(x.answer||"JARVIS returned an empty response."))}catch(e){addMessage("assistant","Request failed: "+e.message)}finally{b.disabled=false}}
async function clearChat(clear=true){try{await fetch("/clear-chat",{method:"POST"})}catch{}document.getElementById("messages").innerHTML='<div class="empty" id="emptyState">Upload a document on the left, then ask JARVIS anything about it.</div>';if(clear)document.getElementById("question").value=""}
document.getElementById("question").addEventListener("keydown",e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();askJarvis()}});checkJarvis();setInterval(checkJarvis,10000);
</script></body></html>"""

def extract_pdf(data):
    reader = PdfReader(io.BytesIO(data))
    parts = []
    for n, page in enumerate(reader.pages, 1):
        text = page.extract_text() or ""
        if text.strip():
            parts.append(f"\n--- PAGE {n} ---\n{text}")
    return "\n".join(parts), f"{len(reader.pages)} pages"

def extract_docx(data):
    doc = Document(io.BytesIO(data))
    parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                parts.append(" | ".join(cells))
    return "\n".join(parts), f"{len(doc.paragraphs)} paragraphs"

def extract_pptx(data):
    pres = Presentation(io.BytesIO(data))
    parts = []
    for n, slide in enumerate(pres.slides, 1):
        texts = [s.text.strip() for s in slide.shapes if hasattr(s, "text") and s.text.strip()]
        if texts:
            parts.append(f"\n--- SLIDE {n} ---\n" + "\n".join(texts))
    return "\n".join(parts), f"{len(pres.slides)} slides"

def extract_txt(data):
    text = data.decode("utf-8", errors="replace")
    return text, f"{len(text.splitlines())} lines"

@app.get("/", response_class=HTMLResponse)
async def website():
    return HTMLResponse(HTML)

@app.get("/jarvis-status")
async def jarvis_status():
    if not GEMINI_API_KEY:
        return {"connected": False, "error": "GEMINI_API_KEY is not configured."}
    return {"connected": True, "jarvis": {"brain": f"Cloud AI • {GEMINI_MODEL}"}}

@app.post("/upload")
async def upload_document(file: UploadFile = File(...)):
    global uploaded_document_text, uploaded_document_name, uploaded_document_type, chat_history
    if not file.filename:
        return {"error": "Please choose a document."}
    ext = os.path.splitext(file.filename.lower())[1]
    if ext not in {".pdf", ".docx", ".pptx", ".txt"}:
        return {"error": "Unsupported file type. Please upload PDF, DOCX, PPTX or TXT."}
    data = await file.read()
    if not data:
        return {"error": "The uploaded document is empty."}
    try:
        if ext == ".pdf":
            text, structure, kind = *extract_pdf(data), "PDF"
        elif ext == ".docx":
            text, structure, kind = *extract_docx(data), "Word document"
        elif ext == ".pptx":
            text, structure, kind = *extract_pptx(data), "PowerPoint"
        else:
            text, structure, kind = *extract_txt(data), "Text document"
    except Exception as error:
        return {"error": f"The document could not be read: {error}"}
    if not text.strip():
        return {"error": "No readable text could be extracted from this document."}
    uploaded_document_text, uploaded_document_name, uploaded_document_type = text, file.filename, kind
    chat_history = []
    return {"filename": file.filename, "document_type": kind, "structure": structure, "characters_extracted": len(text)}

@app.post("/ask")
async def ask_question(question: str = Form(...)):
    global chat_history
    if not uploaded_document_text:
        return {"error": "Upload a document before asking a question."}
    if not question.strip():
        return {"error": "Please enter a question."}
    if not GEMINI_API_KEY:
        return {"error": "Cloud AI is not configured on the server."}

    context_limit = 120000
    document_context = uploaded_document_text[:context_limit]
    context_truncated = len(uploaded_document_text) > context_limit
    history = "\n".join(f"{x['role'].upper()}: {x['content']}" for x in chat_history[-8:])
    prompt = f"""You are JARVIS Scholar, an AI document-analysis assistant created by Saleh Pour.
Answer using the uploaded document as the primary source. Do not invent facts that are not supported by the document. If the answer is not in the document, say so clearly. Be concise but useful. Use previous conversation only to understand follow-up references.

DOCUMENT NAME:
{uploaded_document_name}

DOCUMENT CONTENT:
{document_context}

PREVIOUS DOCUMENT CONVERSATION:
{history or '(none)'}

USER QUESTION:
{question.strip()}
"""
    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.2, "maxOutputTokens": 2048},
    }
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    try:
        async with httpx.AsyncClient(timeout=120.0) as client:
            response = await client.post(
                url,
                headers={"x-goog-api-key": GEMINI_API_KEY, "Content-Type": "application/json"},
                json=payload,
            )
        if response.status_code != 200:
            try:
                detail = response.json().get("error", {}).get("message", "")
            except Exception:
                detail = response.text[:300]
            return {"error": f"Cloud AI returned HTTP {response.status_code}: {detail or 'request failed'}"}
        result = response.json()
        candidates = result.get("candidates") or []
        if not candidates:
            return {"error": "Cloud AI returned no answer."}
        parts = candidates[0].get("content", {}).get("parts", [])
        answer = "\n".join(part.get("text", "") for part in parts if part.get("text")).strip()
        if not answer:
            return {"error": "Cloud AI returned an empty answer."}
        chat_history += [
            {"role": "user", "content": question.strip()},
            {"role": "assistant", "content": answer},
        ]
        return {
            "document": uploaded_document_name,
            "document_type": uploaded_document_type,
            "answer": answer,
            "brain": GEMINI_MODEL,
            "context_truncated": context_truncated,
        }
    except Exception as error:
        return {"error": f"Cloud AI request failed: {error}"}

@app.post("/clear-chat")
def clear_chat():
    global chat_history
    chat_history = []
    return {"message": "Chat history cleared."}

@app.get("/document")
def document_info():
    if not uploaded_document_text:
        return {"message": "No document has been uploaded yet."}
    return {"filename": uploaded_document_name, "document_type": uploaded_document_type, "characters": len(uploaded_document_text), "chat_messages": len(chat_history)}
