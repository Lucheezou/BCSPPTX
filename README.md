# Document to PowerPoint Converter

A web application that converts DOCX documents into professional PowerPoint presentations using AI.

## Features

- Upload DOCX files via drag-and-drop or file selection
- AI-powered content extraction and formatting
- Automatic generation of title, agenda, and content slides
- Professional template with branding
- Direct PowerPoint download

## Environment Variables

Create a `.env` file with:

```
GROQ_API_KEY=your_groq_api_key_here
```

## Local Development

1. Install dependencies:
```bash
npm install
```

2. Start the server:
```bash
npm start
```

3. Open http://localhost:3000

## Deployment

This app is configured for deployment on Render.com. Set the `GROQ_API_KEY` environment variable in your Render dashboard.

## Tech Stack

- Node.js + Express
- Groq AI API (openai/gpt-oss-120b model)
- PptxGenJS for PowerPoint generation
- Mammoth.js for DOCX parsing
- Multer for file uploads