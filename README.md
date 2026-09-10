# JARVIS Scholar

JARVIS Scholar is a web app I built as an extension of my JARVIS project.

The idea was pretty simple: upload a PDF and then be able to actually talk to the document instead of constantly searching through pages manually.

You can ask questions about the PDF, get summaries, pull out the main points and generate study material from it.

## What it does

- Upload and read PDF documents
- Ask questions about the uploaded document
- Generate summaries
- Explain sections in simpler terms
- Pull out key points
- Create study guides
- Ask follow-up questions about the same document

## How I built it

The web application uses **FastAPI** for the backend.

The document is processed by the Scholar backend and questions can then be passed through to the JARVIS AI system through an API.

I kept Scholar separate from the main JARVIS program because I wanted the document interface to work as its own application while still being able to use the AI backend I was already building.

## Why I built it

This actually came from working on JARVIS.

I wanted to see if I could take the assistant and give it a more specialised use case instead of only using it as a general voice assistant.

Studying seemed like a good test because I regularly work with lecture notes, reports and PDFs at university.

The project also gave me more experience with building APIs and connecting a web application to an AI backend rather than having everything inside one Python script.

## Tech

- Python
- FastAPI
- HTML/CSS
- REST APIs
- PDF text extraction
- JARVIS AI backend

## Current state

The main document features are working, but I'm still treating Scholar as an ongoing project.

I'd like to improve document retrieval, handling of larger PDFs and the reliability of answers when documents contain incomplete or conflicting information.

---

Built by **Saleh Pour**  
Robotics & Artificial Intelligence, University of Hertfordshire
