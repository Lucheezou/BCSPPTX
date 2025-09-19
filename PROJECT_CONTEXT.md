# Document to PowerPoint Converter WebApp

## Project Overview
A web application that converts documents into PowerPoint presentations by leveraging AI-powered HTML generation and style matching.

## Workflow
1. **Document Input**: User uploads a document (likely PDF, DOCX, or text format)
2. **Style Reference**: System uses a reference HTML file as a template for styling and structure
3. **AI Processing**: OpenAI API analyzes the input document and generates HTML content that matches the reference style
4. **Style Matching**: The generated HTML adopts:
   - Fonts from reference HTML
   - Structural components and layout
   - Visual design patterns
   - Color schemes and styling
5. **PPT Conversion**: Convert the styled HTML into a PowerPoint (.pptx) file
6. **Download**: User downloads the generated presentation

## Key Components
- Document upload/processing system
- OpenAI API integration for content analysis and HTML generation
- HTML style extraction and matching from reference templates
- HTML to PowerPoint conversion engine
- File download functionality

## Technical Stack
- Frontend: Web interface for document upload and download
- Backend: Document processing, OpenAI API calls, HTML generation
- AI: OpenAI API for intelligent content structuring and style matching
- Conversion: HTML to PowerPoint conversion library/service