# WordConsolidation

A secure, lightweight web application to consolidate revision authors and comments in Microsoft Word (`.docx`) documents.

## Features

- **Privacy First**: All processing occurs in-memory. No files are stored on the server.
- **Anonymization**: Replaces all `w:author` and `w:initials` attributes in the document's internal XML.
- **Easy UI**: Simple drag-and-drop interface powered by Streamlit.
- **Containerized**: Ready to deploy with Docker.

## Quick Start (Docker)

### Prerequisites

- Docker installed on your machine.

### 1. Build the Image

```bash
docker build -t wordconsolidation .
```

### 2. Run the Container

```bash
docker run -p 8501:8501 wordconsolidation
```

### 3. Access the App

Open your browser and navigate to: http://localhost:8501

## Manual Installation (Local)

If you prefer running without Docker:

1.  Ensure you have Python 3.9+ installed.
2.  Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```
3.  Run the app:
    ```bash
    python3 -m streamlit run app.py
    ```
