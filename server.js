const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs-extra');
const mammoth = require('mammoth');
const OpenAI = require('openai');
const { encoding_for_model } = require('tiktoken');
const PptxGenJS = require('pptxgenjs');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Initialize OpenAI client
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

// Initialize tiktoken encoder for token counting
const encoder = encoding_for_model('gpt-4');

// Helper function to count tokens
function countTokens(text) {
  return encoder.encode(text).length;
}

// Helper function to chunk text smartly
function chunkText(text, maxTokens = 2000) {
  const sentences = text.split(/[.!?]+/).filter(s => s.trim().length > 0);
  const chunks = [];
  let currentChunk = '';

  for (const sentence of sentences) {
    const testChunk = currentChunk + sentence + '.';
    if (countTokens(testChunk) > maxTokens && currentChunk.length > 0) {
      chunks.push(currentChunk.trim());
      currentChunk = sentence + '.';
    } else {
      currentChunk = testChunk;
    }
  }

  if (currentChunk.trim().length > 0) {
    chunks.push(currentChunk.trim());
  }

  return chunks;
}

// Helper function to create slides from content chunks
async function processChunk(chunk, chunkIndex, totalChunks, templateStyles, isFirstChunk = false) {
  let prompt;

  if (isFirstChunk) {
    // First chunk gets component-based presentation
    prompt = `You are an expert at converting document content into comprehensive presentation slides using a component-based template system.

Document content:
"""
${chunk}
"""

Template styles to use:
"""
${templateStyles}
"""

AVAILABLE SLIDE COMPONENTS (choose appropriate ones based on content):

1. TITLE PAGE: div.slide - Main presentation title
2. AGENDA PAGE: div.agenda-slide - Agenda with checkmarks
3. CONTENT SLIDE: div.content-slide - Standard content with paragraphs
4. TABLE LAYOUT: div.table-slide - For tabular data and structured information
5. TRANSITION SLIDES: div.transition-slide OR div.transition-alt-slide - Section breaks
6. GRAY TEXT BOXES: div.textbox-slide with textbox-content-gray - Highlighted content boxes
7. TEAL TEXT BOXES: div.textbox-slide with textbox-content-teal - Emphasized content boxes
8. STATISTICS: div.statistics-slide - Charts, metrics, and data visualization
9. CHECKLIST: div.checklist-slide - Action items and task tracking
10. THANK YOU: div.thankyou-slide - Closing slide with Q&A

INTELLIGENT COMPONENT SELECTION:
- Analyze document content to determine appropriate slide types
- Use TABLE LAYOUT for data, comparisons, or structured lists
- Use STATISTICS for numbers, percentages, growth metrics
- Use CHECKLIST for requirements, action items, or compliance steps
- Use TEXT BOXES for important quotes, highlights, or callouts
- Use TRANSITION slides between major sections
- Use THANK YOU as final slide for presentations
- Use standard CONTENT slides for regular text content

CREATE PRESENTATION:
1. Start with TITLE PAGE (extract compelling title from document)
2. Add AGENDA PAGE (based on document structure)
3. Intelligently select appropriate components for document content:
   - Tables/data → TABLE LAYOUT
   - Statistics/numbers → STATISTICS
   - Requirements/tasks → CHECKLIST
   - Important highlights → TEXT BOXES
   - Regular content → CONTENT SLIDE
   - Section breaks → TRANSITION
4. End with THANK YOU slide (MUST copy this EXACT structure from template):
   <div class="thankyou-slide">
     <div class="thankyou-background"></div>
     <div class="thankyou-overlay"></div>
     <div class="thankyou-content">
       <h1 class="thankyou-title">THANK YOU!</h1>
       <h2 class="thankyou-subtitle">Q&A</h2>
     </div>
     <div class="thankyou-disclaimer">The material presented here is for general educational purposes only and is subject to change and law, rules, and regulations. It does not provide legal or tax opinions or advice.</div>
     <div class="thankyou-logo"><img src="/assets/image8.png" alt="BCS Logo"></div>
     <div class="thankyou-page-number">[PAGE_NUMBER]</div>
   </div>

CRITICAL FORMATTING REQUIREMENTS:
- Return ONLY raw HTML slide elements - no markdown, no explanations
- Do NOT use code blocks or backticks anywhere
- Start immediately with <div class="slide"> for title page
- Use exact CSS classes from template components
- Use "/assets/image8.png" for ALL logo references
- Number pages sequentially starting from 1
- Choose components that best match content type and purpose
- CRITICAL: The Thank You Q&A slide MUST be copied EXACTLY from template including ALL elements: thankyou-background, thankyou-overlay, thankyou-content with "THANK YOU!" and "Q&A", thankyou-disclaimer with exact legal text, thankyou-logo, and thankyou-page-number - do NOT modify ANY part of this slide
- MANDATORY: Include the thankyou-disclaimer div with EXACT text: "The material presented here is for general educational purposes only and is subject to change and law, rules, and regulations. It does not provide legal or tax opinions or advice."
- DISCLAIMER REQUIREMENT: Never omit, paraphrase, or modify the disclaimer text - copy it character-for-character

Return ONLY the complete slide HTML elements without html/head/body tags.`;
  } else {
    // Subsequent chunks only get content slides
    prompt = `You are an expert at converting document content into presentation slides.

Document content (chunk ${chunkIndex + 1} of ${totalChunks}):
"""
${chunk}
"""

Template styles to use:
"""
${templateStyles}
"""

Create 1-3 content slides from this content chunk using the provided template styles. Each slide should:
1. Use the exact CSS classes (div.content-slide structure)
2. Have content-header with title and content-logo
3. Have content-body with content-paragraph elements
4. Include content-page-number
5. Be comprehensive but not overcrowded
6. Use: <img src="assets/image8.png" alt="BCS Logo"> for all logos

CRITICAL FORMATTING REQUIREMENTS:
- Return ONLY raw HTML elements - no markdown, no code blocks, no explanations
- Do NOT use code blocks or backticks anywhere in your response
- Do NOT number slides or add section headers
- Start immediately with <div class="content-slide">
- Use exact CSS classes from template
- Use "/assets/image8.png" for ALL logo references

Return ONLY the slide HTML elements (div.content-slide) without the full HTML document structure.`;
  }

  const completion = await openai.responses.create({
    model: "gpt-5",
    input: prompt,
    reasoning: { "effort": "medium" },
    text: { "verbosity": "medium" }
  });

  return completion.output_text;
}

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Use memory storage for cloud deployment
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB limit
  }
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/upload', upload.single('document'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  res.json({
    message: 'File uploaded successfully',
    filename: req.file.filename,
    path: req.file.path
  });
});

app.post('/process-document', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    if (!req.file.originalname.endsWith('.docx')) {
      return res.status(400).json({ error: 'Only DOCX files are supported' });
    }

    // Extract text from DOCX file (using buffer from memory storage)
    const result = await mammoth.extractRawText({ buffer: req.file.buffer });
    const documentText = result.value;

    // Read template HTML and extract styles
    const templateHtml = await fs.readFile('template.html', 'utf8');
    const styleMatch = templateHtml.match(/<style>([\s\S]*?)<\/style>/);
    let templateStyles = styleMatch ? styleMatch[1] : '';

    // Update image paths in styles to use absolute paths from server root
    templateStyles = templateStyles.replace(/url\('desiredresults\/assets\/ppt\/media\//g, "url('/assets/");

    // Process entire document at once with GPT-5
    console.log(`Processing entire document with GPT-5`);

    const prompt = `You are an expert at converting document content into comprehensive presentation slides using a component-based template system.

🚨 CRITICAL TEMPLATE PRESERVATION RULE: The Thank You slide MUST include the exact disclaimer text: "The material presented here is for general educational purposes only and is subject to change and law, rules, and regulations. It does not provide legal or tax opinions or advice." - Do NOT omit or modify this text in any way.

Document content:
"""
${documentText}
"""

Template styles to use:
"""
${templateStyles}
"""

AVAILABLE SLIDE COMPONENTS (choose appropriate ones based on content):

1. TITLE PAGE: div.slide - Main presentation title
2. AGENDA PAGE: div.agenda-slide - Agenda with checkmarks
3. CONTENT SLIDE: div.content-slide - Standard content with paragraphs
4. TABLE LAYOUT: div.table-slide - For tabular data and structured information
5. TRANSITION SLIDES: div.transition-slide OR div.transition-alt-slide - Section breaks
6. GRAY TEXT BOXES: div.textbox-slide with textbox-content-gray - Highlighted content boxes
7. TEAL TEXT BOXES: div.textbox-slide with textbox-content-teal - Emphasized content boxes
8. STATISTICS: div.statistics-slide - Charts, metrics, and data visualization
9. CHECKLIST: div.checklist-slide - Action items and task tracking
10. THANK YOU: div.thankyou-slide - Closing slide with Q&A

INTELLIGENT COMPONENT SELECTION:
- Analyze document content to determine appropriate slide types
- Use TABLE LAYOUT for data, comparisons, or structured lists
- Use STATISTICS for numbers, percentages, growth metrics
- Use CHECKLIST for requirements, action items, or compliance steps
- Use TEXT BOXES for important quotes, highlights, or callouts
- Use TRANSITION slides between major sections
- Use THANK YOU as final slide for presentations
- Use standard CONTENT slides for regular text content

CREATE PRESENTATION:
1. Start with TITLE PAGE (extract compelling title from document)
2. Add AGENDA PAGE (based on document structure)
3. Intelligently select appropriate components for document content:
   - Tables/data → TABLE LAYOUT
   - Statistics/numbers → STATISTICS
   - Requirements/tasks → CHECKLIST
   - Important highlights → TEXT BOXES
   - Regular content → CONTENT SLIDE
   - Section breaks → TRANSITION
4. End with THANK YOU slide (MUST copy this EXACT structure from template):
   <div class="thankyou-slide">
     <div class="thankyou-background"></div>
     <div class="thankyou-overlay"></div>
     <div class="thankyou-content">
       <h1 class="thankyou-title">THANK YOU!</h1>
       <h2 class="thankyou-subtitle">Q&A</h2>
     </div>
     <div class="thankyou-disclaimer">The material presented here is for general educational purposes only and is subject to change and law, rules, and regulations. It does not provide legal or tax opinions or advice.</div>
     <div class="thankyou-logo"><img src="/assets/image8.png" alt="BCS Logo"></div>
     <div class="thankyou-page-number">[PAGE_NUMBER]</div>
   </div>

CRITICAL FORMATTING REQUIREMENTS:
- Return ONLY raw HTML slide elements - no markdown, no explanations
- Do NOT use code blocks or backticks anywhere
- Start immediately with <div class="slide"> for title page
- Use exact CSS classes from template components
- Use "/assets/image8.png" for ALL logo references
- Number pages sequentially starting from 1
- Choose components that best match content type and purpose
- CRITICAL: The Thank You Q&A slide MUST be copied EXACTLY from template including ALL elements: thankyou-background, thankyou-overlay, thankyou-content with "THANK YOU!" and "Q&A", thankyou-disclaimer with exact legal text, thankyou-logo, and thankyou-page-number - do NOT modify ANY part of this slide
- MANDATORY: Include the thankyou-disclaimer div with EXACT text: "The material presented here is for general educational purposes only and is subject to change and law, rules, and regulations. It does not provide legal or tax opinions or advice."
- DISCLAIMER REQUIREMENT: Never omit, paraphrase, or modify the disclaimer text - copy it character-for-character

Return ONLY the complete slide HTML elements without html/head/body tags.`;

    const completion = await openai.responses.create({
      model: "gpt-5",
      input: prompt,
      reasoning: { "effort": "medium" },
      text: { "verbosity": "medium" }
    });

    const allSlides = [completion.output_text];

    // Clean up generated slides and remove markdown artifacts
    const cleanedSlides = allSlides.map(slide => {
      return slide
        .replace(/```html/g, '') // Remove opening markdown
        .replace(/```/g, '') // Remove closing markdown
        .replace(/^\d+\.\s*[^:]*:\s*/gm, '') // Remove numbered headers like "1. Title Page:"
        .replace(/^[A-Z][^:]*:\s*/gm, '') // Remove section headers
        .trim();
    });

    // Combine all slides into complete HTML
    const htmlHead = templateHtml.substring(0, templateHtml.indexOf('</head>'));
    // Update CSS paths in the head section to use absolute paths
    const updatedHtmlHead = htmlHead.replace(/url\('desiredresults\/assets\/ppt\/media\//g, "url('/assets/");
    const combinedSlides = cleanedSlides.join('\n\n');

    const finalHtml = `${updatedHtmlHead}
</head>
<body>
${combinedSlides}
</body>
</html>`;

    // No file cleanup needed - using memory storage

    // Save generated HTML for preview
    const timestamp = Date.now();
    const previewFilename = `presentation_${timestamp}.html`;
    const previewPath = path.join('public', 'previews', previewFilename);

    // Ensure previews directory exists
    await fs.ensureDir(path.join('public', 'previews'));
    await fs.writeFile(previewPath, finalHtml);

    res.json({
      success: true,
      html: finalHtml,
      originalContent: documentText,
      chunksProcessed: allSlides.length,
      previewUrl: `/previews/${previewFilename}`,
      presentationId: timestamp
    });

  } catch (error) {
    console.error('Error processing document:', error);

    // No file cleanup needed - using memory storage

    res.status(500).json({
      error: 'Failed to process document: ' + error.message
    });
  }
});

// Helper function to validate and fix color values for PptxGenJS
function validateColor(color, defaultColor = '000000') {
  // Handle any data type - convert to string first for safety
  if (color === null || color === undefined) {
    console.warn('Null/undefined color detected! Using default:', defaultColor);
    return defaultColor;
  }

  // Convert to string to handle numbers or other types
  const colorStr = String(color);

  if (colorStr === '' || colorStr === 'null' || colorStr === 'undefined') {
    console.warn('Empty/invalid string detected as color value!', color, 'Using default:', defaultColor);
    console.trace(); // This will show us where the empty string is coming from
    return defaultColor;
  }

  // Remove # if present and ensure it's a valid hex color
  const cleanColor = colorStr.replace('#', '').toUpperCase();

  // Check if it's a valid 6-digit hex color
  if (/^[0-9A-F]{6}$/.test(cleanColor)) {
    return cleanColor;
  }

  // Check if it's a valid 3-digit hex color and expand it
  if (/^[0-9A-F]{3}$/.test(cleanColor)) {
    return cleanColor.split('').map(c => c + c).join('');
  }

  console.warn('Invalid color format detected:', color, 'Using default:', defaultColor);
  return defaultColor;
}

// Helper function to clean text content and fix encoding issues
function cleanTextContent(text) {
  if (!text) return '';

  return text
    // Fix HTML entities first
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&#8211;/g, '-')     // en dash entity
    .replace(/&#8212;/g, '-')     // em dash entity
    .replace(/&#8216;/g, "'")     // left single quote entity
    .replace(/&#8217;/g, "'")     // right single quote entity
    .replace(/&#8220;/g, '"')     // left double quote entity
    .replace(/&#8221;/g, '"')     // right double quote entity

    // Fix common special characters that show as ? or �
    .replace(/–/g, '-')           // en dash (U+2013)
    .replace(/—/g, '-')           // em dash (U+2014)
    .replace(/'/g, "'")           // left single quote (U+2018)
    .replace(/'/g, "'")           // right single quote (U+2019)
    .replace(/"/g, '"')           // left double quote (U+201C)
    .replace(/"/g, '"')           // right double quote (U+201D)
    .replace(/…/g, '...')         // ellipsis (U+2026)
    .replace(/§/g, 'Section')     // section symbol (U+00A7)
    .replace(/®/g, '(R)')         // registered trademark (U+00AE)
    .replace(/©/g, '(C)')         // copyright (U+00A9)
    .replace(/™/g, '(TM)')        // trademark (U+2122)

    // Fix specific problematic characters showing as �
    .replace(/\uFFFD/g, ' ')      // replacement character
    .replace(/[\u2000-\u206F]/g, ' ')  // general punctuation block
    .replace(/[\u2E00-\u2E7F]/g, ' ')  // supplemental punctuation

    // Fix zero-width and non-breaking spaces
    .replace(/[\u200B-\u200D\uFEFF]/g, '')  // zero-width spaces
    .replace(/\u00A0/g, ' ')      // non-breaking space

    // Remove any remaining non-printable characters and replacement chars
    .replace(/[\u0000-\u001F\u007F-\u009F\uFFFD]/g, '')

    // Clean up extra whitespace
    .replace(/\s+/g, ' ')
    .trim();
}

// AI-powered HTML content extraction
async function extractContentWithAI(html) {
  try {
    console.log('Using GPT-5 to extract content from HTML...');

    // Clean the HTML input first to remove problematic characters
    const cleanedHtml = cleanTextContent(html);

    const prompt = `Analyze this HTML presentation content and extract the structured data for PowerPoint conversion. The HTML contains various component types that the AI dynamically selected.

HTML Content:
"""
${cleanedHtml}
"""

Extract and return ONLY a JSON object with this structure, analyzing ALL slide components:
{
  "slides": [
    {
      "type": "title",
      "briefing_header": "BCS Monthly Briefing: [Date]",
      "title": "extracted title text",
      "subtitle": "extracted subtitle text"
    },
    {
      "type": "agenda",
      "title": "Agenda",
      "items": ["item 1", "item 2", "item 3"]
    },
    {
      "type": "content",
      "title": "slide title",
      "content": ["paragraph 1", "paragraph 2"]
    },
    {
      "type": "table",
      "title": "table slide title",
      "headers": ["col1", "col2", "col3"],
      "rows": [["data1", "data2", "data3"], ["data4", "data5", "data6"]]
    },
    {
      "type": "statistics",
      "title": "statistics title",
      "description": "stats description",
      "chart_data": [{"year": "2020", "value": 30}],
      "metrics": [{"label": "Revenue Growth", "value": "80%"}]
    },
    {
      "type": "checklist",
      "title": "checklist title",
      "content": ["item 1", "item 2"],
      "checklist_items": [{"text": "item", "checked": true}]
    },
    {
      "type": "textbox",
      "title": "textbox title",
      "boxes": [{"header": "header text", "content": "box content", "color": "gray"}]
    },
    {
      "type": "transition",
      "title": "transition text"
    },
    {
      "type": "thankyou",
      "title": "THANK YOU!",
      "subtitle": "Q&A"
    }
  ]
}

Component Detection Rules:
1. div.slide = title slide
2. div.agenda-slide = agenda slide
3. div.content-slide = content slide
4. div.table-slide = table slide (extract table headers and data)
5. div.statistics-slide = statistics slide (extract charts and metrics)
6. div.checklist-slide = checklist slide (extract checklist items and status)
7. div.textbox-slide = textbox slide (extract textbox content and colors)
8. div.transition-slide OR div.transition-alt-slide = transition slide
9. div.thankyou-slide = thank you slide

Extraction Rules:
- Extract actual content, NOT placeholder text
- For briefing_header: Always format as "BCS Monthly Briefing: [Month Day, Year]" with current date
- For tables: extract th elements as headers, td elements as row data
- For statistics: extract chart data and key metrics
- For checklists: detect checked/unchecked status from classes
- For textboxes: detect gray/teal colors and extract header/content pairs
- Convert HTML entities to proper text
- Return ONLY valid JSON, no explanations or markdown`;

    const completion = await openai.responses.create({
      model: "gpt-5",
      input: prompt,
      reasoning: { "effort": "low" },
      text: { "verbosity": "low" }
    });

    const response = completion.output_text.trim();
    console.log('AI response:', response.substring(0, 300));

    // Clean and parse JSON response
    const cleanResponse = response.replace(/```json\s*/g, '').replace(/```\s*$/g, '').trim();
    const extractedData = JSON.parse(cleanResponse);

    // Clean up text content to fix encoding issues and validate colors
    const cleanedSlides = extractedData.slides.map(slide => {
      const cleanSlide = { ...slide };

      // Clean title and subtitle
      if (cleanSlide.title) {
        cleanSlide.title = cleanTextContent(cleanSlide.title);
      }
      if (cleanSlide.subtitle) {
        cleanSlide.subtitle = cleanTextContent(cleanSlide.subtitle);
      }

      // Clean agenda items
      if (cleanSlide.items) {
        cleanSlide.items = cleanSlide.items.map(item => cleanTextContent(item));
      }

      // Clean content paragraphs
      if (cleanSlide.content) {
        cleanSlide.content = cleanSlide.content.map(paragraph => cleanTextContent(paragraph));
      }

      // Validate and clean color properties in textbox data
      if (cleanSlide.boxes && Array.isArray(cleanSlide.boxes)) {
        cleanSlide.boxes = cleanSlide.boxes.map(box => {
          const cleanBox = { ...box };
          if (cleanBox.color) {
            // Validate color property - only allow 'gray' or 'teal'
            cleanBox.color = (cleanBox.color === 'teal') ? 'teal' : 'gray';
          }
          return cleanBox;
        });
      }

      // Validate checklist items color properties
      if (cleanSlide.checklist_items && Array.isArray(cleanSlide.checklist_items)) {
        cleanSlide.checklist_items = cleanSlide.checklist_items.map(item => {
          const cleanItem = { ...item };
          // Ensure checked property is boolean
          cleanItem.checked = Boolean(cleanItem.checked);
          return cleanItem;
        });
      }

      return cleanSlide;
    });

    console.log('Extracted slides:', cleanedSlides.length);


    return cleanedSlides;

  } catch (error) {
    console.error('AI extraction failed:', error.message);
    console.log('Falling back to regex parsing...');
    return parseHtmlSlidesRegex(html);
  }
}

// Helper function to extract content from HTML slides (fallback)
function parseHtmlSlidesRegex(html) {
  const slides = [];

  console.log('Parsing HTML content for PPT conversion with regex...');
  console.log('HTML length:', html.length);
  console.log('HTML preview (first 500 chars):', html.substring(0, 500));

  // Extract title slide - account for nested content structure
  const titleMatch = html.match(/<div class="slide">([\s\S]*?)<\/div>/);
  console.log('Title match found:', !!titleMatch);
  if (titleMatch) {
    const titleContent = titleMatch[1];
    console.log('Title content:', titleContent.substring(0, 300));

    // Look for title and subtitle within the content div
    const contentDiv = titleContent.match(/<div class="content">([\s\S]*?)<\/div>/);
    if (contentDiv) {
      const contentInner = contentDiv[1];
      console.log('Content inner:', contentInner.substring(0, 200));

      const titleText = contentInner.match(/<div class="title"[^>]*>(.*?)<\/div>/)?.[1] || 'Presentation Title';
      const subtitleText = contentInner.match(/<div class="subtitle"[^>]*>(.*?)<\/div>/)?.[1] || '';

      console.log('Extracted title:', titleText);
      console.log('Extracted subtitle:', subtitleText);

      slides.push({
        type: 'title',
        title: titleText.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&'),
        subtitle: subtitleText.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&')
      });

      console.log('Found title slide:', titleText.replace(/<[^>]*>/g, ''));
    } else {
      console.log('No content div found in title slide');
    }
  } else {
    console.log('No title slide found');
  }

  // Extract agenda slide - updated for current HTML structure
  const agendaMatch = html.match(/<div class="agenda-slide">([\s\S]*?)<\/div>/);
  console.log('Agenda match found:', !!agendaMatch);
  if (agendaMatch) {
    const agendaContent = agendaMatch[1];
    console.log('Agenda content preview:', agendaContent.substring(0, 300));
    const agendaItems = [];

    // Extract agenda items - the pattern that matches the actual HTML structure
    const itemMatches = agendaContent.matchAll(/<li class="agenda-item"><span class="agenda-checkmark">✓<\/span>\s*(.*?)<\/li>/gs);
    console.log('Item matches found with checkmarks:', [...itemMatches].length);

    for (const match of agendaContent.matchAll(/<li class="agenda-item"><span class="agenda-checkmark">✓<\/span>\s*(.*?)<\/li>/gs)) {
      let itemText = match[1].replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').trim();
      console.log('Extracted agenda item:', itemText);
      if (itemText) {
        agendaItems.push(itemText);
      }
    }

    // Fallback: If no items found with checkmarks, try without checkmarks
    if (agendaItems.length === 0) {
      console.log('Trying fallback pattern for agenda items...');
      const itemMatches2 = agendaContent.matchAll(/<li class="agenda-item"[^>]*>(.*?)<\/li>/gs);
      for (const match of itemMatches2) {
        let itemText = match[1].replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/✓/g, '').trim();
        console.log('Fallback extracted agenda item:', itemText);
        if (itemText) {
          agendaItems.push(itemText);
        }
      }
    }

    slides.push({
      type: 'agenda',
      title: 'Agenda',
      items: agendaItems
    });

    console.log(`Found agenda slide with ${agendaItems.length} items:`, agendaItems);
  } else {
    console.log('No agenda slide found');
  }

  // Extract content slides - updated for div structure
  const contentMatches = html.matchAll(/<div class="content-slide">([\s\S]*?)<div class="content-page-number">/g);

  for (const match of contentMatches) {
    const slideContent = match[1];
    const slideTitle = slideContent.match(/<div class="content-title"[^>]*>(.*?)<\/div>/)?.[1] || 'Content Slide';

    const paragraphs = [];
    const paragraphMatches = slideContent.matchAll(/<p class="content-paragraph"[^>]*>([\s\S]*?)<\/p>/g);

    for (const pMatch of paragraphMatches) {
      let text = pMatch[1].replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ').trim();
      if (text) {
        paragraphs.push(text);
      }
    }

    if (paragraphs.length > 0) {
      slides.push({
        type: 'content',
        title: slideTitle.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&'),
        content: paragraphs
      });

      console.log(`Found content slide: "${slideTitle.replace(/<[^>]*>/g, '')}" with ${paragraphs.length} paragraphs`);
    }
  }

  // Extract table slides - for div.table-slide structure
  const tableMatches = html.matchAll(/<div class="table-slide">([\s\S]*?)<div class="table-page-number">/g);
  console.log('Table matches found:', [...html.matchAll(/<div class="table-slide">/g)].length);

  for (const match of tableMatches) {
    const slideContent = match[1];
    const slideTitle = slideContent.match(/<div class="table-title"[^>]*>(.*?)<\/div>/)?.[1] || 'Table';
    console.log('Processing table slide with title:', slideTitle);

    // Extract table headers and rows from HTML table
    const tableElement = slideContent.match(/<table[^>]*>([\s\S]*?)<\/table>/)?.[1];
    if (tableElement) {
      console.log('Found table element:', tableElement.substring(0, 200));

      // Extract headers from th elements
      const headers = [];
      const headerMatches = tableElement.matchAll(/<th[^>]*>(.*?)<\/th>/g);
      for (const headerMatch of headerMatches) {
        const headerText = headerMatch[1].replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').trim();
        if (headerText) {
          headers.push(headerText);
        }
      }

      // Extract rows from tr elements (skip header row)
      const rows = [];
      const rowMatches = tableElement.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g);
      let isFirstRow = true;
      for (const rowMatch of rowMatches) {
        if (isFirstRow) {
          isFirstRow = false; // Skip header row
          continue;
        }

        const rowContent = rowMatch[1];
        const rowData = [];
        const cellMatches = rowContent.matchAll(/<td[^>]*>(.*?)<\/td>/g);
        for (const cellMatch of cellMatches) {
          const cellText = cellMatch[1].replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').trim();
          rowData.push(cellText);
        }

        if (rowData.length > 0) {
          rows.push(rowData);
        }
      }

      console.log(`Extracted table data - Headers: ${headers.length}, Rows: ${rows.length}`);
      console.log('Headers:', headers);
      console.log('Rows sample:', rows.slice(0, 2));

      if (headers.length > 0 && rows.length > 0) {
        slides.push({
          type: 'table',
          title: slideTitle.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&'),
          headers: headers,
          rows: rows
        });

        console.log(`Found table slide: "${slideTitle.replace(/<[^>]*>/g, '')}" with ${headers.length} headers and ${rows.length} rows`);
      } else {
        console.log('Table slide had no valid data, skipping');
      }
    } else {
      console.log('No table element found in table slide');
    }
  }

  console.log(`Total slides parsed: ${slides.length}`);
  return slides;
}

// Keep old function name for any other references
function parseHtmlSlides(html) {
  return parseHtmlSlidesRegex(html);
}

// PPT conversion endpoint using improved HTML processing
app.post('/convert-to-ppt', express.json(), async (req, res) => {
  try {
    const { presentationId, html } = req.body;

    if (!html) {
      return res.status(400).json({ error: 'No HTML content provided' });
    }

    console.log('Converting HTML to PPT using enhanced processing...');

    // Create PowerPoint presentation
    const pptx = new PptxGenJS();


    // Set slide size to 16:9
    pptx.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 5.625 });
    pptx.layout = 'LAYOUT_16x9';

    // Use AI to extract content from HTML
    const slides = await extractContentWithAI(html);

    if (!slides || slides.length === 0) {
      return res.status(400).json({ error: 'No slides found in HTML content' });
    }

    console.log(`Found ${slides.length} slides to convert`);

    // Process each slide using improved extraction
    for (let i = 0; i < slides.length; i++) {
      const slideData = slides[i];

      console.log(`Processing slide ${i + 1}: ${slideData.type}...`);

      // Only create slide for known types
      if (slideData.type === 'title') {
        const slide = pptx.addSlide();
        // Title slide with background image - simple sizing (works best with 16:9 images)
        slide.addImage({
          path: 'public/assets/image9.jpg',
          x: 0, y: 0, w: 10, h: 5.625
        });

        // Add purple overlay
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 5.625,
          fill: { color: '28295D', transparency: 20 }
        });

        // BCS Monthly Briefing header line
        const briefingHeader = slideData.briefing_header || `BCS Monthly Briefing: ${new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })}`;
        slide.addText(briefingHeader, {
          x: 0.5, y: 2.0, w: 8.5, h: 0.8,
          fontSize: 42, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', valign: 'middle', bold: false
        });

        // Brief title on line below
        if (slideData.title) {
          slide.addText(slideData.title, {
            x: 0.5, y: 3.0, w: 8.5, h: 0.8,
            fontSize: 32, color: validateColor('FFFFFF'), fontFace: 'Lato',
            align: 'left', valign: 'middle', bold: false
          });
        }

        // Logo in lower right corner with padding (2:1 aspect ratio, half height as padding)
        const logoHeight = 1.0;
        const logoWidth = logoHeight * 2; // 2:1 aspect ratio
        const padding = logoHeight / 2;
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 10 - logoWidth - padding, y: 5.625 - logoHeight - padding, w: logoWidth, h: logoHeight
        });

        console.log(`Created title slide: "${slideData.title}"`);

      } else if (slideData.type === 'agenda') {
        const slide = pptx.addSlide();
        // Agenda slide with background image - simple sizing (works best with 16:9 images)
        slide.addImage({
          path: 'public/assets/image7.jpg',
          x: 0, y: 0, w: 10, h: 5.625
        });

        // Add purple overlay
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 5.625,
          fill: { color: '28295D', transparency: 20 }
        });

        // Agenda title in middle upper left (not all caps, smaller font)
        slide.addText('Agenda', {
          x: 0.4, y: 1.4, w: 3.5, h: 1.0,
          fontSize: 42, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle'
        });

        // Right side - white background for agenda items (full height and width)
        slide.addShape(pptx.ShapeType.rect, {
          x: 4, y: 0, w: 6, h: 5.625,
          fill: { color: 'FFFFFF', transparency: 0 }
        });

        // Right side - agenda items with checkmarks
        let yPos = 1.4;
        const maxItems = Math.min(slideData.items.length, 6); // Limit to 6 items to prevent overflow

        for (let j = 0; j < maxItems; j++) {
          const item = slideData.items[j];
          if (!item || item.trim() === '') continue;

          // Add green checkmark
          slide.addText('✓', {
            x: 4.3, y: yPos, w: 0.3, h: 0.5,
            fontSize: 20, color: validateColor('7CB342'), fontFace: 'Lato',
            align: 'center', bold: true, valign: 'middle'
          });

          // Add agenda item text - smaller font
          const safeItem = (item || '').toString();
          const itemText = safeItem.length > 60 ? safeItem.substring(0, 57) + '...' : safeItem;
          slide.addText(itemText, {
            x: 4.8, y: yPos, w: 4.8, h: 0.5,
            fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
            align: 'left', valign: 'middle'
          });
          yPos += 0.6;
        }

        // Logo in lower left with less padding (moved more to the left)
        const logoHeight = 1.0;
        const logoWidth = logoHeight * 2; // 2:1 aspect ratio
        const padding = 0.3; // Reduced padding to move logo more to the left
        slide.addImage({
          path: 'public/assets/image8.png',
          x: padding, y: 5.625 - logoHeight - (logoHeight / 2), w: logoWidth, h: logoHeight
        });

        // Page number
        slide.addText('2', {
          x: 9.2, y: 5, w: 0.6, h: 0.4,
          fontSize: 20, color: validateColor('28295D'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        console.log(`Created agenda slide with ${slideData.items.length} items`);

      } else if (slideData.type === 'content') {
        const slide = pptx.addSlide();
        // Content slide
        slide.background = { color: 'F5F5F5' };

        // Purple header
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 1.1,
          fill: { color: '28295D' }
        });

        // Content title - truncate if too long
        const titleText = slideData.title.length > 80 ? slideData.title.substring(0, 77) + '...' : slideData.title;
        slide.addText(titleText, {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: 28, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle'
        });

        // Add brand logo in header (2:1 aspect ratio)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8.5, y: 0.2, w: 1.2, h: 0.6
        });

        

        // Content paragraphs - improved text handling to prevent overflow
        const safeContent = slideData.content && Array.isArray(slideData.content) ? slideData.content : ['No content available'];
        const allContent = safeContent.join('\n\n');

        // More conservative character limit for better formatting
        if (allContent.length > 600) {
          // Split content into smaller, manageable chunks
          const chunks = [];
          let currentChunk = '';

          for (const paragraph of slideData.content) {
            const testChunk = currentChunk + (currentChunk ? '\n\n' : '') + paragraph;
            if (testChunk.length > 600 && currentChunk.length > 0) {
              chunks.push(currentChunk.trim());
              currentChunk = paragraph;
            } else {
              currentChunk = testChunk;
            }
          }

          if (currentChunk.trim()) {
            chunks.push(currentChunk.trim());
          }

          // Add first chunk to current slide - prevent overflow
          slide.addText(chunks[0], {
            x: 0.6, y: 1.5, w: 8.8, h: 3.0, // Reduced height
            fontSize: 14, color: validateColor('333333'), fontFace: 'Lato', // Smaller font
            align: 'left', lineSpacing: 18, valign: 'top', // Tighter spacing
            wrap: true
          });

          // Add page number
          slide.addText(`${i + 1}`, {
            x: 9, y: 5.1, w: 0.8, h: 0.4,
            fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
            align: 'center', valign: 'middle'
          });

          // Create additional slides for remaining chunks
          for (let c = 1; c < chunks.length; c++) {
            const contSlide = pptx.addSlide();
            contSlide.background = { color: 'F5F5F5' };

            contSlide.addShape(pptx.ShapeType.rect, {
              x: 0, y: 0, w: 10, h: 1.1,
              fill: { color: '28295D' }
            });

            const contTitleText = titleText.length > 65 ? titleText.substring(0, 62) + '...' : titleText;
            contSlide.addText(`${contTitleText} (cont.)`, {
              x: 0.5, y: 0.15, w: 7, h: 0.8,
              fontSize: 28, color: validateColor('FFFFFF'), fontFace: 'Lato',
              align: 'left', bold: false, valign: 'middle'
            });

            // Add brand logo to continuation slide (2:1 aspect ratio)
            contSlide.addImage({
              path: 'public/assets/image8.png',
              x: 8.5, y: 0.2, w: 1.2, h: 0.6
            });

            contSlide.addText(chunks[c], {
              x: 0.6, y: 1.5, w: 8.8, h: 3.0, // Reduced height
              fontSize: 14, color: validateColor('333333'), fontFace: 'Lato', // Smaller font
              align: 'left', lineSpacing: 18, valign: 'top', // Tighter spacing
              wrap: true
            });

            contSlide.addText(`${i + 1}-${c + 1}`, {
              x: 9, y: 5.1, w: 0.8, h: 0.4,
              fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
              align: 'center', valign: 'middle'
            });
          }
        } else {
          // Content fits on one slide - prevent overflow
          slide.addText(allContent, {
            x: 0.6, y: 1.5, w: 8.8, h: 3.0, // Reduced height to prevent overflow
            fontSize: 14, color: validateColor('333333'), fontFace: 'Lato', // Smaller font
            align: 'left', lineSpacing: 18, valign: 'top', // Tighter line spacing
            wrap: true
          });

          // Add page number
          slide.addText(`${i + 1}`, {
            x: 9, y: 5.1, w: 0.8, h: 0.4,
            fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
            align: 'center', valign: 'middle'
          });
        }

        console.log(`Created content slide: "${slideData.title}" with ${slideData.content.length} paragraphs`);

      } else if (slideData.type === 'table') {
        const slide = pptx.addSlide();
        // Table slide
        slide.background = { color: 'F5F5F5' };


        // Purple header
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 1.1,
          fill: { color: '28295D' }
        });

        // Table title
        slide.addText(slideData.title || 'Table', {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: 28, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle'
        });

        // Add brand logo in header (2:1 aspect ratio)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8.5, y: 0.2, w: 1.2, h: 0.6
        });

        // Create table using manual row-by-row approach to get proper colors
        if (slideData.headers && slideData.rows) {
          const tableData = [slideData.headers, ...slideData.rows];

          // Create table manually with proper styling for each row
          const startY = 1.4;
          const rowHeight = 0.5; // Increased for better text spacing
          const colWidth = 9.2 / tableData[0].length; // Distribute width evenly

          tableData.forEach((row, rowIndex) => {
            row.forEach((cell, colIndex) => {
              const x = 0.4 + (colIndex * colWidth);
              const y = startY + (rowIndex * rowHeight);

              // Determine colors based on row
              let fillColor, textColor;
              if (rowIndex === 0) {
                // Header row
                fillColor = validateColor('28295D');
                textColor = validateColor('FFFFFF');
              } else if (rowIndex % 2 === 1) {
                // Odd data rows (1, 3, 5...) - teal
                fillColor = validateColor('A8D5D5');
                textColor = validateColor('333333');
              } else {
                // Even data rows (2, 4, 6...) - light gray
                fillColor = validateColor('E8E8E8');
                textColor = validateColor('333333');
              }

              // Add cell background
              slide.addShape(pptx.ShapeType.rect, {
                x: x, y: y, w: colWidth, h: rowHeight,
                fill: { color: fillColor },
                line: { color: validateColor('FFFFFF'), width: 2 }
              });

              // Add cell text with better padding and wrapping
              slide.addText(String(cell || ''), {
                x: x + 0.1, y: y + 0.1, w: colWidth - 0.2, h: rowHeight - 0.2,
                fontSize: rowIndex === 0 ? 12 : 10, // Smaller fonts for better fit
                color: textColor,
                fontFace: 'Lato',
                align: rowIndex === 0 ? 'center' : 'left',
                valign: 'middle',
                bold: rowIndex === 0,
                wrap: true, // Enable text wrapping
                lineSpacing: 14 // Better line spacing
              });
            });
          });
        }

        // Add page number
        slide.addText(`${i + 1}`, {
          x: 9, y: 5.1, w: 0.8, h: 0.4,
          fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        console.log(`Created table slide: "${slideData.title}"`);

      } else if (slideData.type === 'statistics') {
        const slide = pptx.addSlide();
        // Statistics slide
        slide.background = { color: 'F5F5F5' };

        // Purple header
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 1.1,
          fill: { color: '28295D' }
        });

        // Statistics title
        slide.addText(slideData.title || 'Slide Title', {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: 28, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle'
        });

        // Add brand logo in header (2:1 aspect ratio)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8.5, y: -0.2, w: 1.2, h: 0.6
        });

        // Right side statistics panel
        slide.addShape(pptx.ShapeType.rect, {
          x: 6, y: 1.3, w: 3.8, h: 3.5,
          fill: { color: '28295D' }
        });

        // Statistics panel title
        slide.addText('Statistics', {
          x: 6.2, y: 1.5, w: 3.4, h: 0.5,
          fontSize: 24, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: true, valign: 'middle'
        });

        // Description
        if (slideData.description) {
          slide.addText(slideData.description, {
            x: 6.2, y: 2.0, w: 3.4, h: 1.0,
            fontSize: 12, color: validateColor('FFFFFF'), fontFace: 'Lato',
            align: 'left', valign: 'top', wrap: true
          });
        }

        // Metrics grid
        if (slideData.metrics && slideData.metrics.length > 0) {
          let metricY = 3.2;
          for (let m = 0; m < Math.min(4, slideData.metrics.length); m++) {
            const metric = slideData.metrics[m];
            const xPos = m % 2 === 0 ? 6.2 : 8.0;
            const yPos = metricY + Math.floor(m / 2) * 0.8;

            // Metric value
            slide.addText(metric.value, {
              x: xPos, y: yPos, w: 1.6, h: 0.4,
              fontSize: 24, color: validateColor('FFFFFF'), fontFace: 'Lato',
              align: 'center', bold: true, valign: 'middle'
            });

            // Metric label
            slide.addText(metric.label, {
              x: xPos, y: yPos + 0.3, w: 1.6, h: 0.3,
              fontSize: 10, color: validateColor('FFFFFF'), fontFace: 'Lato',
              align: 'center', valign: 'top'
            });
          }
        }

        // Add page number
        slide.addText(`${i + 1}`, {
          x: 9, y: 5.1, w: 0.8, h: 0.4,
          fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        console.log(`Created statistics slide: "${slideData.title}"`);

      } else if (slideData.type === 'checklist') {
        const slide = pptx.addSlide();
        // Checklist slide
        slide.background = { color: '28295D' };

        // Header
        slide.addText(slideData.title || 'Slide Title', {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: 28, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle'
        });

        // Add brand logo in header (2:1 aspect ratio)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8.5, y: 0.2, w: 1.2, h: 0.6
        });

        // Left content area - prevent text overflow
        if (slideData.content && slideData.content.length > 0) {
          const safeChecklist = Array.isArray(slideData.content) ? slideData.content : [slideData.content];
          const contentText = safeChecklist.join('\n\n');
          // Truncate if too long to prevent overflow
          const maxLength = 500;
          const displayText = contentText.length > maxLength ? contentText.substring(0, maxLength) + '...' : contentText;

          slide.addText(displayText, {
            x: 0.5, y: 1.3, w: 6.7, h: 3.8,
            fontSize: 14, color: validateColor('FFFFFF'), fontFace: 'Lato',
            align: 'left', valign: 'top', wrap: true, lineSpacing: 18
          });
        }

        // Right checklist panel - full coverage from header to bottom, no padding
        slide.addShape(pptx.ShapeType.rect, {
          x: 8.0, y: 1.1, w: 2.0, h: 4.525,
          fill: { color: 'FFFFFF' }
        });

        // Checklist items
        if (slideData.checklist_items && slideData.checklist_items.length > 0) {
          let checkY = 1.2;
          for (const item of slideData.checklist_items.slice(0, 6)) {
            // Checkbox or checkmark
            const symbol = item.checked ? '✓' : '☐';
            const symbolColor = validateColor(item.checked ? '7CB342' : '999999');

            slide.addText(symbol, {
              x: 8.0, y: checkY, w: 0.3, h: 0.3,
              fontSize: 10, color: symbolColor, fontFace: 'Lato',
              align: 'center', valign: 'middle'
            });

            // Item text
            slide.addText(item.text || 'Item', {
              x: 8.2, y: checkY, w: 1.5, h: 0.3,
              fontSize: 6, color: validateColor('333333'), fontFace: 'Lato',
              align: 'left', valign: 'middle'
            });

            checkY += 0.4;
          }
        }

        // Add page number
        slide.addText(`${i + 1}`, {
          x: 9, y: 5.1, w: 0.8, h: 0.4,
          fontSize: 18, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        console.log(`Created checklist slide: "${slideData.title}"`);

      } else if (slideData.type === 'textbox') {
        const slide = pptx.addSlide();
        // Textbox slide
        slide.background = { color: 'F5F5F5' };

        // Purple header
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 1.1,
          fill: { color: '28295D' }
        });

        // Textbox title
        slide.addText(slideData.title || 'Slide Title', {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: 28, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle'
        });

        // Add brand logo in header (2:1 aspect ratio)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8.5, y: 0.2, w: 1.2, h: 0.6
        });

        // Add textboxes side by side with padding
        if (slideData.boxes && slideData.boxes.length > 0) {
          const boxWidth = 4.2;
          const boxHeight = 3.5;
          const padding = 0.3;

          for (let b = 0; b < Math.min(2, slideData.boxes.length); b++) {
            const box = slideData.boxes[b];
            const xPos = b === 0 ? 0.5 + padding : 5.0 + padding;
            const fillColor = validateColor((box && box.color === 'teal') ? 'A8D5D5' : 'E8E8E8');

            // Textbox header
            slide.addShape(pptx.ShapeType.rect, {
              x: xPos, y: 1.3 + padding, w: boxWidth, h: 0.6,
              fill: { color: '28295D' }
            });

            slide.addText(box.header || 'Header Text', {
              x: xPos + 0.2, y: 1.4 + padding, w: boxWidth - 0.4, h: 0.4,
              fontSize: 14, color: validateColor('FFFFFF'), fontFace: 'Lato',
              align: 'left', bold: true, valign: 'middle'
            });

            // Textbox content
            slide.addShape(pptx.ShapeType.rect, {
              x: xPos, y: 1.9 + padding, w: boxWidth, h: boxHeight - 0.6,
              fill: { color: fillColor }
            });

            // Truncate content if too long to prevent overflow
            const content = box.content || 'Content text goes here...';
            const maxLength = 300; // Limit content length
            const displayContent = content.length > maxLength ? content.substring(0, maxLength) + '...' : content;

            slide.addText(displayContent, {
              x: xPos + 0.2, y: 2.1 + padding, w: boxWidth - 0.4, h: boxHeight - 1.2,
              fontSize: 11, color: validateColor('333333'), fontFace: 'Lato',
              align: 'left', valign: 'top', wrap: true, lineSpacing: 15
            });
          }
        }

        // Add page number
        slide.addText(`${i + 1}`, {
          x: 9, y: 5.1, w: 0.8, h: 0.4,
          fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        console.log(`Created textbox slide: "${slideData.title}"`);

      } else if (slideData.type === 'transition') {
        const slide = pptx.addSlide();
        // Transition slide
        slide.addImage({
          path: 'public/assets/image1.jpg',
          x: 0, y: 0, w: 10, h: 5.625
        });

        // No overlay for transition slide

        // Transition content box (moved more to the left)
        slide.addShape(pptx.ShapeType.rect, {
          x: 0.5, y: 1.5, w: 6, h: 2.5,
          fill: { color: '28295D' }
        });

        // Transition title (all caps)
        const transitionTitle = (slideData.title || 'TRANSITION SLIDE').toUpperCase();
        slide.addText(transitionTitle, {
          x: 1, y: 2.2, w: 5, h: 1.1,
          fontSize: 36, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: true, valign: 'middle'
        });

        // Add brand logo (2:1 aspect ratio)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8, y: 0.2, w: 1.6, h: 0.8
        });

        // Add page number
        slide.addText(`${i + 1}`, {
          x: 9, y: 5.1, w: 0.8, h: 0.4,
          fontSize: 18, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        console.log(`Created transition slide: "${slideData.title}"`);

      } else if (slideData.type === 'thankyou') {
        const slide = pptx.addSlide();
        // Thank You slide - keep exactly as template
        slide.addImage({
          path: 'public/assets/image10.jpg',
          x: 0, y: 0, w: 10, h: 5.625
        });

        // Add purple overlay
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 5.625,
          fill: { color: '28295D', transparency: 15 }
        });

        // Thank you title
        slide.addText('THANK YOU!', {
          x: 1, y: 1.8, w: 6, h: 1.0,
          fontSize: 48, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle'
        });

        // Q&A subtitle
        slide.addText('Q&A', {
          x: 1, y: 2.8, w: 6, h: 0.8,
          fontSize: 48, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle'
        });

        // Disclaimer text
        const disclaimer = 'The material presented here is for general educational purposes only and is subject to change and law, rules, and regulations. It does not provide legal or tax opinions or advice.';
        slide.addText(disclaimer, {
          x: 1, y: 4.5, w: 5.5, h: 0.8,
          fontSize: 10, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', valign: 'top', wrap: true, transparency: 10
        });

        // Add brand logo in lower right corner (same as title page)
        const logoHeight = 1.0;
        const logoWidth = logoHeight * 2; // 2:1 aspect ratio
        const padding = logoHeight / 2;
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 10 - logoWidth - padding, y: 5.625 - logoHeight - padding, w: logoWidth, h: logoHeight
        });

        // Add page number
        slide.addText(`${i + 1}`, {
          x: 9, y: 5.1, w: 0.8, h: 0.4,
          fontSize: 18, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        console.log(`Created thank you slide - preserved as template`);

      } else {
        console.log(`Unknown slide type: ${slideData.type} - skipping slide creation`);
        // Don't create slide for unknown types
      }
    }

    // Generate PPT file
    const pptFilename = `presentation_${presentationId || Date.now()}.pptx`;
    const pptPath = path.join('public', 'downloads', pptFilename);

    // Ensure downloads directory exists
    await fs.ensureDir(path.join('public', 'downloads'));

    // Write PPT file
    await pptx.writeFile({ fileName: pptPath });

    res.json({
      success: true,
      downloadUrl: `/downloads/${pptFilename}`,
      filename: pptFilename
    });

  } catch (error) {
    console.error('Error converting to PPT:', error);
    res.status(500).json({
      error: 'Failed to convert to PowerPoint: ' + error.message
    });
  }
});

// Helper function for fallback content extraction with proper nested structure
function extractSlideContent(slideHtml) {
  if (slideHtml.includes('class="slide"')) {
    // For title slide, look inside content div
    const contentDiv = slideHtml.match(/<div class="content">([\s\S]*?)<\/div>/);
    if (contentDiv) {
      const title = contentDiv[1].match(/<div class="title"[^>]*>(.*?)<\/div>/)?.[1] || '';
      const subtitle = contentDiv[1].match(/<div class="subtitle"[^>]*>(.*?)<\/div>/)?.[1] || '';
      return {
        type: 'title',
        title: title.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&'),
        subtitle: subtitle.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&')
      };
    }
  } else if (slideHtml.includes('class="agenda-slide"')) {
    const items = [];
    const itemMatches = slideHtml.matchAll(/<li class="agenda-item"><span class="agenda-checkmark">✓<\/span>\s*(.*?)<\/li>/gs);
    for (const match of itemMatches) {
      items.push(match[1].replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').trim());
    }
    return { type: 'agenda', items };
  } else if (slideHtml.includes('class="content-slide"')) {
    const title = slideHtml.match(/<div class="content-title"[^>]*>(.*?)<\/div>/)?.[1] || '';
    const content = [];
    const paragraphMatches = slideHtml.matchAll(/<p class="content-paragraph"[^>]*>([\s\S]*?)<\/p>/g);
    for (const match of paragraphMatches) {
      content.push(match[1].replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').trim());
    }
    return {
      type: 'content',
      title: title.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&'),
      content
    };
  }
  return { type: 'unknown' };
}

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});