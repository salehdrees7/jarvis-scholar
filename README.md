# JARVIS Scholar

JARVIS Scholar is an AI-powered document intelligence platform built as an extension of my JARVIS AI assistant.

It allows users to upload PDF documents and interact with them through natural-language questions, summaries and study tools.

## Features

- Upload and extract text from PDF documents
- Ask natural-language questions about uploaded documents
- Generate document summaries
- Explain document content in simpler terms
- Extract key points
- Generate study guides
- Continue follow-up conversations about the same document
- Live connection monitoring between the web application and JARVIS
- Local AI inference through the JARVIS backend

## Architecture

JARVIS Scholar runs as a separate FastAPI application and communicates with the JARVIS AI backend through an API.

```text
User
  ↓
JARVIS Scholar Web Interface
  ↓
FastAPI Document Backend
  ↓
JARVIS API
  ↓
Local AI Model
