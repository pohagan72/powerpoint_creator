# AI Document to Presentation Converter

The **AI Document to Presentation Converter** is a web application that allows users to convert their documents (`.docx` and `.pdf`) into professionally styled PowerPoint presentations. Leveraging Artificial Intelligence, the app designs a detailed slide outline complete with speaker notes, visual suggestions, and actionable tips, specifically tailored to your chosen tone, audience, and visual style.

---

## Table of Contents

- [Features](#features)
- [How It Works](#how-it-works)
- [Installation & Running Locally](#installation--running-locally)
- [Dockerized Deployment](#dockerized-deployment)
- [Configuration](#configuration)
- [Project Structure](#project-structure)
- [License](#license)

---

## Features

- **Multi-file Support:** Upload documents in `.docx` or `.pdf` format (up to 32MB).
- **AI-Powered Presentation Generation:** Transforms document content into a robust PowerPoint outline with:
  - Slide Title, Content Type, Key Message
  - Bulleted points, Visual Suggestions, Design Notes
  - Detailed Elaboration, Enhancement Suggestions, and Best Practice Tips
- **Customizable Visual Themes:** Choose from Professional, Creative, or Minimalist design templates.
- **Content Personalization:** Specify the target audience and tone to tailor the presentation content.
- **User-friendly Interface:** Supports drag-and-drop file uploads with real-time status updates.
- **Dockerized Application:** Easily deployable using Docker.

---

## How It Works

1. **Upload and Configuration:**  
   - The user selects a visual theme and optionally specifies details like target audience and desired tone.
   - A `.docx` or `.pdf` document is uploaded via a simple drag-and-drop interface or by browsing files.

2. **Document Processing:**  
   - The system extracts text content using libraries such as `python-docx` for DOCX files and `PyPDF2` for PDFs.
   - The extracted text is truncated (if needed) to meet token limits and then formatted into a detailed prompt.

3. **AI Driven Slide Generation:**  
   - A detailed prompt is sent to an Azure-hosted OpenAI service, which processes it to generate a structured slide outline.
   - The returned outline contains all mandatory fields for each slide including title, bullets, elaboration, and actionable tips.

4. **Download Presentation:**  
   - Once processing is complete, the generated presentation in PowerPoint format (`.pptx`) is made available for download.

---

## Installation & Running Locally

### Prerequisites

- Python 3.8 or later
- `pip` package manager

### Setup

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/yourusername/powerpoint_creator.git
   cd powerpoint_creator
   ```

2. **Create and Activate a Virtual Environment (Optional but Recommended):**

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install Dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

4. **Environment Variables:**

   Create a `.env` file in the root directory with the following content (adjust values as needed):

   ```env
   AZURE_OPENAI_ENDPOINT=https://<your-azure-openai-endpoint>
   AZURE_OPENAI_API_KEY=<your-api-key>
   ```

5. **Run the Application:**

   ```bash
   python run.py
   ```

   The application will be available at [http://localhost:5000](http://localhost:5000).

---

## Dockerized Deployment

This application has been dockerized, allowing you to run it inside a container easily. The Docker image is available on Docker Hub:

[**paulohagan/ppt_creator**](https://hub.docker.com/repository/docker/paulohagan/ppt_creator/general)

### Running with Docker

1. **Pull the Docker Image:**

   ```bash
   docker pull paulohagan/ppt_creator
   ```

2. **Run the Container:**

   ```bash
   docker run -d -p 5000:5000 --name ppt_creator paulohagan/ppt_creator
   ```

   The app should now be accessible at [http://localhost:5000](http://localhost:5000).

3. **Setting Environment Variables:**

   If you need to configure the Azure OpenAI settings or other environment variables, pass them during container startup:

   ```bash
   docker run -d -p 5000:5000 \
     -e AZURE_OPENAI_ENDPOINT="https://<your-azure-endpoint>" \
     -e AZURE_OPENAI_API_KEY="<your-api-key>" \
     --name ppt_creator paulohagan/ppt_creator
   ```

---

## Configuration

- **Upload & Generated Files:**  
  - Uploaded files are stored in the `uploads` directory.
  - Generated PowerPoint presentations are saved in the `generated` directory.

- **Templates:**  
  The app supports different visual themes defined in `app.py`:
  - Professional
  - Creative
  - Minimalist

- **AI Service:**  
  The AI prompt is sent to an Azure OpenAI endpoint specified by the environment variables:
  - `AZURE_OPENAI_ENDPOINT`
  - `AZURE_OPENAI_API_KEY`

---

## Project Structure

```plaintext
powerpoint_creator/
├── app.py                   # Main application logic for file processing and API calls
├── run.py                   # Entry point using Waitress for production
├── requirements.txt         # Python dependencies
├── Dockerfile               # Docker configuration (if available)
├── uploads/                 # Directory for uploaded files
├── generated/               # Directory for generated presentations
└── templates/
    └── index.html           # Frontend UI for the web app
```

---

## License

This project is licensed under the [MIT License](LICENSE).

---
