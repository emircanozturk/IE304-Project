# 🔴 METU-IE Summer Practice Intelligent Assistant

An interactive, AI-powered virtual consultant designed to help Middle East Technical University (METU) Industrial Engineering students navigate the official requirements, deadlines, and paperwork for their IE 300 and IE 400 Summer Practices.

This application is built using **Streamlit** and the **OpenAI API** (gpt-4o-mini), utilizing a local Retrieval-Augmented Generation (RAG) approach to parse official department documents and provide strictly accurate, hallucination-free answers.

## ✨ Key Features

* **Intelligent Document Parsing:** Automatically ingests and reads official guidelines from `.pdf`, `.docx`, `.pptx`, and `.txt` files placed in the local knowledge base.
* **Strict Guardrails:** Programmed to only answer based on official METU guidelines, explicitly preventing the AI from making up false deadlines or applying external assumptions.
* **Native Bilingual Support:** Instantly switches the entire User Interface and the AI's response language between English (ENG) and Turkish (TR).
* **Interactive FAQ Shortcuts:** One-click prompt buttons for the most common student questions regarding company constraints, paperwork, and prerequisites.
* **Responsive UI:** Custom-styled Streamlit interface matching METU's official branding.

## 🏗️ System Architecture

1. **Frontend:** Streamlit handles the web-based conversational interface, sidebar navigation, and dynamic language toggling.
2. **Knowledge Base:** The `load_knowledge_base()` function pre-processes official files from the `/knowledge_base` directory using `pypdf`, `python-docx`, and `python-pptx`.
3. **LLM Engine:** The extracted text and user prompts are passed securely to OpenAI's `gpt-4o-mini` model, constrained by a strict System Prompt to ensure domain-specific accuracy.

## 📂 Project Structure

```text
ie304-chatbot-project/
│
├── app.py                 # Main Streamlit application script
├── requirements.txt       # Python dependencies for deployment
├── metu_logo.png          # Official METU logo for the UI
├── README.md              # Project documentation
│
├── knowledge_base/        # Folder containing all official documentation
│   ├── rules.pdf
│   ├── ie400_presentation.pptx
│   └── faq.docx
│
└── .streamlit/            # Hidden folder for secure environment variables
    └── secrets.toml       # Local API Key storage (IGNORED BY GIT)
