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
2. AGENDA PAGE: div.agenda-slide - Agenda with checkmarks (use concise article titles)
3. CONTENT SLIDE: div.content-slide - Standard content with BULLETED LISTS (no long paragraphs)
4. TABLE LAYOUT: div.table-slide - For tabular data and structured information
5. TRANSITION SLIDES: div.transition-slide OR div.transition-alt-slide - Section breaks (vary by category: "In the News", "Federal Update", "Hot Topics", "Question of the Month")
6. GRAY TEXT BOXES: div.textbox-slide with textbox-content-gray - Highlighted content boxes
7. TEAL TEXT BOXES: div.textbox-slide with textbox-content-teal - Emphasized content boxes
8. CHECKLIST SIDEBAR: div.checklist-slide - For employer implications and action items (SIDEBAR FORMAT - HIGHLY USEFUL)
9. THANK YOU: div.thankyou-slide - Closing slide with Q&A and disclaimer (ALWAYS INCLUDE)

EXCLUDED COMPONENTS (DO NOT USE):
- STATISTICS: div.statistics-slide - DO NOT use this component unless explicitly requested

INTELLIGENT COMPONENT SELECTION:
- Analyze document content to determine appropriate slide types
- Use TABLE LAYOUT for data, comparisons, or structured lists (auto-resize to fit data)
- DO NOT USE STATISTICS slides - they are excluded by default
- Use CHECKLIST SIDEBAR for employer implications (HIGHLY USEFUL - ALWAYS RETAIN THIS FORMAT)
- Use CHECKLIST for requirements, action items, or compliance steps
- Use TEAL TEXT BOXES for highlighted important information
- Use GRAY TEXT BOXES for quotes, highlights, or callouts
- Use TRANSITION slides between article categories - vary by category:
  * "In the News" - use transition-slide
  * "Federal Update" - use transition-alt-slide
  * "Hot Topics" - use transition-slide
  * "Question of the Month" - use transition-alt-slide
- Use THANK YOU as final slide for presentations (MANDATORY - always include with disclaimer)
- Use standard CONTENT slides for regular text content - FORMAT AS BULLETED LISTS, NOT PARAGRAPHS

CREATE PRESENTATION:
1. Start with TITLE PAGE (extract compelling title from document)
2. Add AGENDA PAGE with concise article titles aligned with content
3. For EACH ARTICLE, create a 3-SLIDE STRUCTURE:
   a) BACKGROUND + APPLICABLE EMPLOYERS slide (CONTENT slide with bulleted format)
   b) "GO DEEPER" slide (dedicated slide preserving nuance and context from article's "go deeper" section)
   c) EMPLOYER IMPLICATIONS slide (use CHECKLIST SIDEBAR format - this is the highly useful sidebar template)
4. Between article categories, add TRANSITION slides:
   - "In the News" → transition-slide
   - "Federal Update" → transition-alt-slide
   - "Hot Topics" → transition-slide
   - "Question of the Month" → transition-alt-slide
5. Intelligently select appropriate components for document content:
   - Tables/data → TABLE LAYOUT (ensure proper auto-resize)
   - DO NOT use STATISTICS slides
   - Requirements/tasks → CHECKLIST (sidebar format)
   - Employer implications → TEAL TEXT BOXES (highly useful - always retain)
   - Important highlights → GRAY TEXT BOXES
   - Regular content → CONTENT SLIDE with BULLETED LISTS (extract main points and sub-points as bullets and sub-bullets)
6. End with THANK YOU slide (MANDATORY - MUST copy this EXACT structure from template):
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
2. AGENDA PAGE: div.agenda-slide - Agenda with checkmarks (use concise article titles)
3. CONTENT SLIDE: div.content-slide - Standard content with BULLETED LISTS (no long paragraphs)
4. TABLE LAYOUT: div.table-slide - For tabular data and structured information
5. TRANSITION SLIDES: div.transition-slide OR div.transition-alt-slide - Section breaks (vary by category: "In the News", "Federal Update", "Hot Topics", "Question of the Month")
6. GRAY TEXT BOXES: div.textbox-slide with textbox-content-gray - Highlighted content boxes
7. TEAL TEXT BOXES: div.textbox-slide with textbox-content-teal - Emphasized content boxes
8. CHECKLIST SIDEBAR: div.checklist-slide - For employer implications and action items (SIDEBAR FORMAT - HIGHLY USEFUL)
9. THANK YOU: div.thankyou-slide - Closing slide with Q&A and disclaimer (ALWAYS INCLUDE)

EXCLUDED COMPONENTS (DO NOT USE):
- STATISTICS: div.statistics-slide - DO NOT use this component unless explicitly requested

INTELLIGENT COMPONENT SELECTION:
- Analyze document content to determine appropriate slide types
- Use TABLE LAYOUT for data, comparisons, or structured lists (auto-resize to fit data)
- DO NOT USE STATISTICS slides - they are excluded by default
- Use CHECKLIST SIDEBAR for employer implications (HIGHLY USEFUL - ALWAYS RETAIN THIS FORMAT)
- Use CHECKLIST for requirements, action items, or compliance steps
- Use TEAL TEXT BOXES for highlighted important information
- Use GRAY TEXT BOXES for quotes, highlights, or callouts
- Use TRANSITION slides between article categories - vary by category:
  * "In the News" - use transition-slide
  * "Federal Update" - use transition-alt-slide
  * "Hot Topics" - use transition-slide
  * "Question of the Month" - use transition-alt-slide
- Use THANK YOU as final slide for presentations (MANDATORY - always include with disclaimer)
- Use standard CONTENT slides for regular text content - FORMAT AS BULLETED LISTS, NOT PARAGRAPHS

CREATE PRESENTATION:
1. Start with TITLE PAGE (extract compelling title from document)
2. Add AGENDA PAGE with concise article titles aligned with content
3. For EACH ARTICLE, create a 3-SLIDE STRUCTURE:
   a) BACKGROUND + APPLICABLE EMPLOYERS slide (CONTENT slide with bulleted format)
   b) "GO DEEPER" slide (dedicated slide preserving nuance and context from article's "go deeper" section)
   c) EMPLOYER IMPLICATIONS slide (use CHECKLIST SIDEBAR format - this is the highly useful sidebar template)
4. Between article categories, add TRANSITION slides:
   - "In the News" → transition-slide
   - "Federal Update" → transition-alt-slide
   - "Hot Topics" → transition-slide
   - "Question of the Month" → transition-alt-slide
5. Intelligently select appropriate components for document content:
   - Tables/data → TABLE LAYOUT (ensure proper auto-resize)
   - DO NOT use STATISTICS slides
   - Requirements/tasks → CHECKLIST (sidebar format)
   - Employer implications → TEAL TEXT BOXES (highly useful - always retain)
   - Important highlights → GRAY TEXT BOXES
   - Regular content → CONTENT SLIDE with BULLETED LISTS (extract main points and sub-points as bullets and sub-bullets)
6. End with THANK YOU slide (MANDATORY - MUST copy this EXACT structure from template):
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

// Helper function to calculate dynamic font size for slide headers
function getHeaderFontSize(title) {
  const length = title.length;
  if (length > 80) return 20;
  if (length > 60) return 22;
  if (length > 45) return 24;
  if (length > 35) return 26;
  return 28;
}

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
      "items": ["Main item 1", "  Sub-item 1a", "  Sub-item 1b", "Main item 2"]
    },
    {
      "type": "content",
      "title": "slide title",
      "content": ["• Main point 1", "  ◦ Sub-point 1a", "  ◦ Sub-point 1b", "• Main point 2"],
      "bullets": true
    },
    {
      "type": "go_deeper",
      "title": "Go Deeper: [Article Title]",
      "content": ["• Detailed point 1", "• Detailed point 2"]
    },
    {
      "type": "table",
      "title": "table slide title",
      "headers": ["col1", "col2", "col3"],
      "rows": [["data1", "data2", "data3"], ["data4", "data5", "data6"]]
    },
    {
      "type": "checklist",
      "title": "checklist title",
      "content": ["item 1", "item 2"],
      "checklist_heading": "Action Items",
      "checklist_panel_text": "Brief paragraph text before checklist",
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
      "type": "transition_alt",
      "title": "alternate transition text"
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
3. div.content-slide = content slide (use for background + applicable employers)
4. div.content-slide with "Go Deeper" in title = go_deeper slide type
5. div.table-slide = table slide (extract table headers and data)
6. div.checklist-slide = checklist slide (SIDEBAR FORMAT - use for employer implications)
7. div.textbox-slide = textbox slide (extract textbox content and colors)
8. div.transition-slide OR div.transition-alt-slide = transition slide
9. div.thankyou-slide = thank you slide

IMPORTANT: Employer Implications should ALWAYS use div.checklist-slide (sidebar format), NOT textbox

CRITICAL EXTRACTION RULES:
- Extract actual content, NOT placeholder text
- For briefing_header: Always format as "BCS Monthly Briefing: [Month Day, Year]" with current date

BULLET FORMATTING (MANDATORY - NESTED LIST DETECTION):
- ALWAYS preserve bullet hierarchies from nested <ul> lists in HTML
- Detect nested lists: parent <li> contains nested <ul> with child <li> items
- For parent bullets: NO prefix, just the text (e.g., "Background")
- For sub-bullets (nested <ul><li>): Add DASH PREFIX "- " at start (e.g., "- The ACA has become...")
- Extract format example: ["Background", "- Sub-point 1", "- Sub-point 2", "Next Topic", "- Another sub"]
- NEVER use • or ◦ characters - use DASH for sub-bullets only
- Parent bullets have NO prefix at all

AGENDA FORMATTING (CRITICAL - MANDATORY):
- MUST preserve hierarchical structure with TWO-SPACE indentation for nested items
- For items like "Hot Topics: Title" followed by nested items:
  * Extract parent as just the category: "Hot Topics"
  * Extract nested item with TWO SPACES: "  1094/1095 IRS e-filing..."
- Main items: NO spaces at start, e.g., "In the News"
- Sub-items: EXACTLY TWO SPACES at start, e.g., "  Sub-topic"
- Detect nested items: <li class="agenda-item nested"> has class "nested"
- If an item has a colon and is followed by a nested item, split at colon and use first part as parent
- Extract format: ["In the News", "Hot Topics", "  ACA Reporting Deadlines", "Federal Update"]
- DO NOT trim() or strip() leading spaces from sub-items
- The two-space prefix "  " is REQUIRED for proper PPT indentation

SLIDE TYPE DETECTION:
- "Go Deeper" in title or heading → type: "go_deeper"
- "Employer Implications" or "Takeaways" in title → type: "checklist" (sidebar format)
- "Implications" or "Action Items" in title → type: "checklist" (sidebar format)
- Sidebar/checklist format → type: "checklist"
- Tables with headers/rows → type: "table"
- Transition headers for "In the News" or "Hot Topics" → type: "transition"
- Transition headers for "Federal Update" or "Question of the Month" → type: "transition_alt"
- Other transition headers → type: "transition"
- NEVER create type: "statistics" (excluded by default)

OTHER RULES:
- For tables: extract th elements as headers, td elements as row data
- For checklists:
  * Extract checklist_heading from h2.checklist-content-heading (e.g., "Action Items")
  * Extract checklist_panel_text from div.checklist-panel-content (brief paragraph before checklist)
  * Extract checklist_items array with text and checked status from ul.checklist-list li elements
  * Extract main content from div.checklist-content div.checklist-item elements
- For textboxes: detect gray/teal colors and extract header/content pairs
- IMPORTANT: Employer implications should ALWAYS be type: "checklist" (not textbox)
- Identify article categories and use appropriate transition slide types
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

      // Clean agenda items - PRESERVE leading spaces for sub-bullets
      if (cleanSlide.items) {
        cleanSlide.items = cleanSlide.items.map(item => {
          if (!item) return '';
          // Clean text but preserve leading spaces for indentation
          const leadingSpaces = item.match(/^(\s*)/)[0];
          const cleanedText = cleanTextContent(item);
          // Re-add leading spaces if they were removed
          return leadingSpaces + cleanedText;
        });
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

    // Extract agenda items - detect nested items with class="nested"
    console.log('Item matches found with checkmarks:', [...agendaContent.matchAll(/<li class="agenda-item/g)].length);

    for (const match of agendaContent.matchAll(/<li class="agenda-item(?:\s+nested)?"[^>]*>.*?<span class="agenda-checkmark">✓<\/span>\s*(.*?)<\/li>/gs)) {
      const fullMatch = match[0];
      const isNested = fullMatch.includes('class="agenda-item nested"');
      let itemText = match[1].replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').trim();

      // Add two-space prefix for nested items
      if (isNested && !itemText.startsWith('  ')) {
        itemText = '  ' + itemText;
      }

      console.log('Extracted agenda item:', itemText, isNested ? '(nested)' : '');
      if (itemText) {
        agendaItems.push(itemText);
      }
    }

    // Fallback: If no items found with checkmarks, try without checkmarks
    if (agendaItems.length === 0) {
      console.log('Trying fallback pattern for agenda items...');
      const itemMatches2 = agendaContent.matchAll(/<li class="agenda-item(?:\s+nested)?"[^>]*>(.*?)<\/li>/gs);
      for (const match of itemMatches2) {
        const fullMatch = match[0];
        const isNested = fullMatch.includes('class="agenda-item nested"');
        let itemText = match[1].replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/✓/g, '').trim();

        // Add two-space prefix for nested items
        if (isNested && !itemText.startsWith('  ')) {
          itemText = '  ' + itemText;
        }

        console.log('Fallback extracted agenda item:', itemText, isNested ? '(nested)' : '');
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

  // Extract content slides - updated for div structure with bullet support
  const contentMatches = html.matchAll(/<div class="content-slide">([\s\S]*?)<div class="content-page-number">/g);

  for (const match of contentMatches) {
    const slideContent = match[1];
    const slideTitle = slideContent.match(/<div class="content-title"[^>]*>(.*?)<\/div>/)?.[1] || 'Content Slide';
    const cleanTitle = slideTitle.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&');

    // Determine if this is a "Go Deeper" slide
    const isGoDeeper = cleanTitle.toLowerCase().includes('go deeper');

    const bullets = [];

    // Try to extract bullets from <li> elements first
    const listItemMatches = slideContent.matchAll(/<li[^>]*>([\s\S]*?)<\/li>/g);
    for (const liMatch of listItemMatches) {
      let text = liMatch[1].replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ').trim();
      if (text && !text.startsWith('•')) {
        text = '• ' + text;
      }
      if (text) {
        bullets.push(text);
      }
    }

    // If no list items, try paragraphs and convert to bullets
    if (bullets.length === 0) {
      const paragraphMatches = slideContent.matchAll(/<p class="content-paragraph"[^>]*>([\s\S]*?)<\/p>/g);
      for (const pMatch of paragraphMatches) {
        let text = pMatch[1].replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ').trim();
        if (text && !text.startsWith('•')) {
          text = '• ' + text;
        }
        if (text) {
          bullets.push(text);
        }
      }
    }

    if (bullets.length > 0) {
      slides.push({
        type: isGoDeeper ? 'go_deeper' : 'content',
        title: cleanTitle,
        content: bullets,
        bullets: true
      });

      console.log(`Found ${isGoDeeper ? 'Go Deeper' : 'content'} slide: "${cleanTitle}" with ${bullets.length} bullets`);
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
          fontSize: 32, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', valign: 'middle', bold: false
        });

        // Brief title on line below
        if (slideData.title) {
          slide.addText(slideData.title, {
            x: 0.5, y: 3.0, w: 8.5, h: 0.8,
            fontSize: 24, color: validateColor('FFFFFF'), fontFace: 'Lato',
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

        // Right side - agenda items with checkmarks - centered vertically with proper spacing
        const maxItems = Math.min(slideData.items.length, 12);
        const baseLineHeight = 0.5;  // Increased spacing for wrapped items

        // Calculate total height needed for all items
        const totalContentHeight = maxItems * baseLineHeight;

        // Center vertically: start position = (slide height - content height) / 2
        let yPos = (5.625 - totalContentHeight) / 2;

        for (let j = 0; j < maxItems; j++) {
          const item = slideData.items[j];
          if (!item || item.trim() === '') continue;

          // Detect if item is a sub-bullet (starts with spaces or specific characters)
          const isSubBullet = item.startsWith('  ') || item.startsWith('◦') || item.startsWith('- ');
          const indent = isSubBullet ? 0.3 : 0;

          // Add green checkmark bullet
          slide.addText('✓', {
            x: 4.2 + indent, y: yPos, w: 0.3, h: 0.4,
            fontSize: isSubBullet ? 16 : 18,
            color: validateColor('7CB342'), fontFace: 'Lato',
            align: 'center', bold: true, valign: 'top'
          });

          // Add agenda item text with wrapping and proper height for multi-line
          const safeItem = (item || '').toString().trim().replace(/^[◦\-\s]*/, '');
          slide.addText(safeItem, {
            x: 4.6 + indent, y: yPos, w: 5.0 - indent, h: baseLineHeight - 0.05,
            fontSize: isSubBullet ? 13 : 15,
            color: validateColor('28295D'), fontFace: 'Lato',
            align: 'left', valign: 'top', bold: false,
            wrap: true,
            lineSpacing: 16
          });
          yPos += baseLineHeight;
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

        // Content title with dynamic font sizing
        const titleText = slideData.title;
        const headerFontSize = getHeaderFontSize(titleText);
        slide.addText(titleText, {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: headerFontSize, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle', wrap: true
        });

        // Add brand logo in header (2:1 aspect ratio)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8.5, y: 0.2, w: 1.2, h: 0.6
        });



        // Content as bulleted lists - add each bullet separately for proper rendering
        const safeContent = slideData.content && Array.isArray(slideData.content) ? slideData.content : ['No content available'];

        // Add bullets directly - calculate height based on text length for wrapping
        let yPos = 1.5;
        const maxYPos = 5.0; // Don't go below this

        for (let j = 0; j < safeContent.length && yPos < maxYPos; j++) {
          const item = safeContent[j];
          const trimmed = item.trim();
          const isSubBullet = trimmed.startsWith('-');
          const text = isSubBullet ? trimmed.substring(1).trim() : trimmed;

          const indent = isSubBullet ? 0.4 : 0;
          const bulletSymbol = isSubBullet ? '◦' : '•';

          // Estimate height based on text length (rough calculation)
          // Assume ~100 chars per line at 14pt font with given width
          const textWidth = 8.5 - indent;
          const charsPerLine = textWidth * 12; // Rough estimate: 12 chars per inch
          const numLines = Math.ceil(text.length / charsPerLine);
          const itemHeight = Math.max(0.35, numLines * 0.25); // Min 0.35", ~0.25" per line

          // Stop if this item won't fit
          if (yPos + itemHeight > maxYPos) break;

          // Add bullet symbol
          slide.addText(bulletSymbol, {
            x: 0.6 + indent, y: yPos, w: 0.2, h: 0.3,
            fontSize: isSubBullet ? 12 : 14,
            color: validateColor('333333'), fontFace: 'Lato',
            align: 'left', valign: 'top'
          });

          // Add text with calculated height
          slide.addText(text, {
            x: 0.9 + indent, y: yPos, w: 8.5 - indent, h: itemHeight,
            fontSize: 14, color: validateColor('333333'), fontFace: 'Lato',
            align: 'left', valign: 'top', wrap: true, lineSpacing: 18
          });

          yPos += itemHeight + 0.05; // Add small gap between items
        }

        // Add page number
        slide.addText(`${i + 1}`, {
          x: 9, y: 5.1, w: 0.8, h: 0.4,
          fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        console.log(`Created content slide: "${slideData.title}" with ${slideData.content.length} paragraphs`);

      } else if (slideData.type === 'go_deeper') {
        // Go Deeper slide - add each bullet separately for proper rendering
        const titleText = slideData.title;
        const safeContent = slideData.content && Array.isArray(slideData.content) ? slideData.content : ['No content available'];

        const slide = pptx.addSlide();
        slide.background = { color: 'F5F5F5' };

        // Purple header
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 1.1,
          fill: { color: '28295D' }
        });

        // Go Deeper title with dynamic font
        const headerFontSize = getHeaderFontSize(titleText);
        slide.addText(titleText, {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: headerFontSize, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle', wrap: true
        });

        // Add brand logo in header (2:1 aspect ratio)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8.5, y: 0.2, w: 1.2, h: 0.6
        });

        // Add "Go Deeper" section header with gray background
        slide.addShape(pptx.ShapeType.rect, {
          x: 0.6, y: 1.3, w: 8.8, h: 0.5,
          fill: { color: 'E8E8E8' }
        });

        slide.addText('Go Deeper', {
          x: 0.8, y: 1.35, w: 8.4, h: 0.4,
          fontSize: 16, color: validateColor('28295D'), fontFace: 'Lato',
          align: 'left', bold: true, valign: 'middle'
        });

        // Add bullets directly - calculate height based on text length for wrapping
        let yPos = 2.0;
        const maxYPos = 5.0; // Don't go below this

        for (let j = 0; j < safeContent.length && yPos < maxYPos; j++) {
          const item = safeContent[j];
          const trimmed = item.trim();
          const isSubBullet = trimmed.startsWith('-');
          const text = isSubBullet ? trimmed.substring(1).trim() : trimmed;

          const indent = isSubBullet ? 0.4 : 0;
          const bulletSymbol = isSubBullet ? '◦' : '•';

          // Estimate height based on text length (rough calculation)
          // Assume ~100 chars per line at 13pt font with given width
          const textWidth = 8.5 - indent;
          const charsPerLine = textWidth * 12; // Rough estimate: 12 chars per inch
          const numLines = Math.ceil(text.length / charsPerLine);
          const itemHeight = Math.max(0.35, numLines * 0.25); // Min 0.35", ~0.25" per line

          // Stop if this item won't fit
          if (yPos + itemHeight > maxYPos) break;

          // Add bullet symbol
          slide.addText(bulletSymbol, {
            x: 0.6 + indent, y: yPos, w: 0.2, h: 0.3,
            fontSize: isSubBullet ? 11 : 13,
            color: validateColor('333333'), fontFace: 'Lato',
            align: 'left', valign: 'top'
          });

          // Add text with calculated height
          slide.addText(text, {
            x: 0.9 + indent, y: yPos, w: 8.5 - indent, h: itemHeight,
            fontSize: 13, color: validateColor('333333'), fontFace: 'Lato',
            align: 'left', valign: 'top', wrap: true, lineSpacing: 16
          });

          yPos += itemHeight + 0.05; // Add small gap between items
        }

        // Add page number
        slide.addText(`${i + 1}`, {
          x: 9, y: 5.1, w: 0.8, h: 0.4,
          fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        console.log(`Created Go Deeper slide: "${slideData.title}"`);

      } else if (slideData.type === 'table') {
        // Dynamic table creation with smart sizing and splitting
        if (slideData.headers && slideData.rows) {
          const headers = slideData.headers;
          const rows = slideData.rows;
          const numCols = headers.length;
          const totalRows = rows.length;

          // Calculate content metrics for intelligent sizing
          const columnWidths = new Array(numCols).fill(0);
          const allData = [headers, ...rows];

          allData.forEach(row => {
            row.forEach((cell, colIndex) => {
              const cellLength = String(cell || '').length;
              columnWidths[colIndex] = Math.max(columnWidths[colIndex], cellLength);
            });
          });

          // Normalize column widths
          const totalContentWidth = columnWidths.reduce((a, b) => a + b, 0);
          const availableWidth = 9.2;
          const normalizedWidths = columnWidths.map(w => (w / totalContentWidth) * availableWidth);

          // Determine optimal font sizes based on content
          const maxContentLength = Math.max(...columnWidths);
          let baseFontSize = 10;
          let headerFontSize = 12;

          if (maxContentLength > 80) {
            baseFontSize = 7;
            headerFontSize = 9;
          } else if (maxContentLength > 50) {
            baseFontSize = 8;
            headerFontSize = 10;
          } else if (maxContentLength > 30) {
            baseFontSize = 9;
            headerFontSize = 11;
          }

          // Calculate how many rows fit per slide (max available height / min row height)
          const availableHeight = 3.8;
          const minRowHeight = 0.35;
          const maxRowHeight = 0.6;

          // Dynamic row height based on number of rows and content
          let rowHeight = Math.max(minRowHeight, Math.min(maxRowHeight, availableHeight / (totalRows + 1)));

          // Adjust row height if content is very long (needs wrapping)
          if (maxContentLength > 50) {
            rowHeight = Math.max(rowHeight, 0.45);
          }

          // Calculate rows per slide
          const rowsPerSlide = Math.floor(availableHeight / rowHeight) - 1; // -1 for header
          const numTableSlides = Math.ceil(totalRows / rowsPerSlide);

          console.log(`Table split into ${numTableSlides} slides (${rowsPerSlide} rows per slide)`);

          // Create table slides
          for (let slideNum = 0; slideNum < numTableSlides; slideNum++) {
            const slide = pptx.addSlide();
            slide.background = { color: 'F5F5F5' };

            // Purple header
            slide.addShape(pptx.ShapeType.rect, {
              x: 0, y: 0, w: 10, h: 1.1,
              fill: { color: '28295D' }
            });

            // Table title with continuation indicator and dynamic font
            const titleSuffix = numTableSlides > 1 ? ` (${slideNum + 1}/${numTableSlides})` : '';
            const tableTitle = (slideData.title || 'Table') + titleSuffix;
            const headerFontSize = getHeaderFontSize(tableTitle);
            slide.addText(tableTitle, {
              x: 0.5, y: 0.15, w: 7, h: 0.8,
              fontSize: headerFontSize, color: validateColor('FFFFFF'), fontFace: 'Lato',
              align: 'left', bold: false, valign: 'middle', wrap: true
            });

            // Add brand logo in header (2:1 aspect ratio)
            slide.addImage({
              path: 'public/assets/image8.png',
              x: 8.5, y: 0.2, w: 1.2, h: 0.6
            });

            // Determine which rows to display on this slide
            const startRow = slideNum * rowsPerSlide;
            const endRow = Math.min(startRow + rowsPerSlide, totalRows);
            const slideRows = rows.slice(startRow, endRow);

            // Create header + data for this slide
            const tableData = [headers, ...slideRows];
            const startY = 1.4;
            let currentX = 0.4;

            tableData.forEach((row, rowIndex) => {
              currentX = 0.4; // Reset X for each row
              row.forEach((cell, colIndex) => {
                const colWidth = normalizedWidths[colIndex];
                const y = startY + (rowIndex * rowHeight);

                // Determine colors based on row
                let fillColor, textColor;
                if (rowIndex === 0) {
                  // Header row
                  fillColor = validateColor('28295D');
                  textColor = validateColor('FFFFFF');
                } else if (rowIndex % 2 === 1) {
                  // Odd data rows - teal
                  fillColor = validateColor('A8D5D5');
                  textColor = validateColor('333333');
                } else {
                  // Even data rows - light gray
                  fillColor = validateColor('E8E8E8');
                  textColor = validateColor('333333');
                }

                // Add cell background
                slide.addShape(pptx.ShapeType.rect, {
                  x: currentX, y: y, w: colWidth, h: rowHeight,
                  fill: { color: fillColor },
                  line: { color: validateColor('FFFFFF'), width: 2 }
                });

                // Calculate font size dynamically per cell
                const cellContent = String(cell || '');
                let cellFontSize = rowIndex === 0 ? headerFontSize : baseFontSize;

                // For header cells, adjust font based on BOTH cell length AND column width
                if (rowIndex === 0) {
                  // Calculate optimal header font size based on content and width
                  const charsPerInch = 8; // Rough estimate for readability
                  const maxChars = colWidth * charsPerInch;

                  if (cellContent.length > maxChars) {
                    // Content won't fit comfortably, reduce font
                    const ratio = maxChars / cellContent.length;
                    cellFontSize = Math.max(7, Math.floor(headerFontSize * ratio));
                  } else if (cellContent.length > 25) {
                    // Medium length header
                    cellFontSize = Math.min(headerFontSize, 10);
                  }
                }

                // Further reduce font if cell content is very long (for data cells)
                if (rowIndex > 0 && cellContent.length > 60) {
                  cellFontSize = Math.max(6, cellFontSize - 1);
                }

                // Add cell text with dynamic sizing
                slide.addText(cellContent, {
                  x: currentX + 0.05, y: y + 0.05, w: colWidth - 0.1, h: rowHeight - 0.1,
                  fontSize: cellFontSize,
                  color: textColor,
                  fontFace: 'Lato',
                  align: rowIndex === 0 ? 'center' : 'left',
                  valign: 'middle',
                  bold: rowIndex === 0,
                  wrap: true,
                  lineSpacing: 10
                });

                currentX += colWidth;
              });
            });

            // Add page number
            slide.addText(`${i + 1}${slideNum > 0 ? '-' + (slideNum + 1) : ''}`, {
              x: 9, y: 5.1, w: 0.8, h: 0.4,
              fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
              align: 'center', valign: 'middle'
            });

            console.log(`Created table slide ${slideNum + 1}/${numTableSlides}: "${slideData.title}" with ${slideRows.length} rows`);
          }
        }

      } else if (slideData.type === 'statistics') {
        const slide = pptx.addSlide();
        // Statistics slide
        slide.background = { color: 'F5F5F5' };

        // Purple header
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 1.1,
          fill: { color: '28295D' }
        });

        // Statistics title with dynamic font
        const statTitle = slideData.title || 'Slide Title';
        const headerFontSize = getHeaderFontSize(statTitle);
        slide.addText(statTitle, {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: headerFontSize, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle', wrap: true
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

        // Header with dynamic font
        const checklistTitle = slideData.title || 'Slide Title';
        const headerFontSize = getHeaderFontSize(checklistTitle);
        slide.addText(checklistTitle, {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: headerFontSize, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle', wrap: true
        });

        // Add brand logo in header (2:1 aspect ratio)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8.5, y: 0.2, w: 1.2, h: 0.6
        });

        // Left content area with heading
        let contentStartY = 1.3;

        // Add checklist heading if present (e.g., "Action Items")
        if (slideData.checklist_heading) {
          slide.addText(slideData.checklist_heading, {
            x: 0.5, y: 1.3, w: 6.7, h: 0.4,
            fontSize: 18, color: validateColor('FFFFFF'), fontFace: 'Lato',
            align: 'left', bold: true, valign: 'middle'
          });
          contentStartY = 1.8;
        }

        // Left content area with bullets - prevent text overflow
        if (slideData.content && slideData.content.length > 0) {
          const safeChecklist = Array.isArray(slideData.content) ? slideData.content : [slideData.content];

          // Process content to ensure bullets are present
          const bulletContent = safeChecklist.map(item => {
            const trimmed = item.trim();
            // Remove existing bullet/number if present
            const cleaned = trimmed.replace(/^[\d\)•◦\-]\s*/, '');
            return cleaned;
          });

          const contentText = bulletContent.join('\n');
          // Truncate if too long to prevent overflow
          const maxLength = 500;
          const displayText = contentText.length > maxLength ? contentText.substring(0, maxLength) + '...' : contentText;

          const contentHeight = 5.1 - contentStartY - 0.3;
          slide.addText(displayText, {
            x: 0.5, y: contentStartY, w: 6.7, h: contentHeight,
            fontSize: 14, color: validateColor('FFFFFF'), fontFace: 'Lato',
            align: 'left', valign: 'top', wrap: true, lineSpacing: 18,
            bullet: true  // Add bullets to checklist content
          });
        }

        // Right checklist panel - full coverage from header to bottom, no padding
        slide.addShape(pptx.ShapeType.rect, {
          x: 7.5, y: 1.1, w: 2.5, h: 4.525,
          fill: { color: 'FFFFFF' }
        });

        let checkY = 1.3;

        // Add brief paragraph before checklist if present
        if (slideData.checklist_panel_text) {
          slide.addText(slideData.checklist_panel_text, {
            x: 7.6, y: checkY, w: 2.3, h: 1.0,
            fontSize: 8, color: validateColor('333333'), fontFace: 'Lato',
            align: 'left', valign: 'top', wrap: true, lineSpacing: 12
          });
          checkY += 1.1; // Move down after paragraph
        }

        // Checklist items
        if (slideData.checklist_items && slideData.checklist_items.length > 0) {
          for (const item of slideData.checklist_items.slice(0, 8)) {
            // Checkbox or checkmark
            const symbol = item.checked ? '✓' : '☐';
            const symbolColor = validateColor(item.checked ? '7CB342' : '999999');

            slide.addText(symbol, {
              x: 7.6, y: checkY, w: 0.25, h: 0.25,
              fontSize: 10, color: symbolColor, fontFace: 'Lato',
              align: 'center', valign: 'middle'
            });

            // Item text
            slide.addText(item.text || 'Item', {
              x: 7.9, y: checkY, w: 1.9, h: 0.25,
              fontSize: 7, color: validateColor('333333'), fontFace: 'Lato',
              align: 'left', valign: 'middle', wrap: true
            });

            checkY += 0.35;
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

        // Textbox title with dynamic font
        const textboxTitle = slideData.title || 'Slide Title';
        const headerFontSize = getHeaderFontSize(textboxTitle);
        slide.addText(textboxTitle, {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: headerFontSize, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle', wrap: true
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

      } else if (slideData.type === 'transition' || slideData.type === 'transition_alt') {
        const slide = pptx.addSlide();

        // Use alternate transition style for transition_alt
        const isAltTransition = slideData.type === 'transition_alt';

        // Transition slide - use image1.jpg for standard, image5.jpg for alt (matching template)
        slide.addImage({
          path: isAltTransition ? 'public/assets/image5.jpg' : 'public/assets/image1.jpg',
          x: 0, y: 0, w: 10, h: 5.625
        });

        // Add overlay for alt transition only
        if (isAltTransition) {
          slide.addShape(pptx.ShapeType.rect, {
            x: 0, y: 0, w: 10, h: 5.625,
            fill: { color: '28295D', transparency: 20 }  // 0.8 opacity = 20% transparency
          });
        }

        // Transition title (all caps)
        const transitionTitle = (slideData.title || 'TRANSITION SLIDE').toUpperCase();

        if (isAltTransition) {
          // Alternative transition: centered white bordered box
          slide.addShape(pptx.ShapeType.rect, {
            x: 2.0, y: 1.69, w: 6.0, h: 2.25,  // Centered box (60% width, 40% height)
            fill: { type: 'solid', color: '000000', transparency: 100 },  // Transparent fill
            line: { color: 'FFFFFF', width: 4 }  // White border
          });

          slide.addText(transitionTitle, {
            x: 2.0, y: 1.69, w: 6.0, h: 2.25,
            fontSize: 40, color: validateColor('FFFFFF'), fontFace: 'Lato',
            align: 'center', bold: true, valign: 'middle'
          });
        } else {
          // Standard transition: purple box on left
          slide.addShape(pptx.ShapeType.rect, {
            x: 0.5, y: 1.5, w: 6, h: 2.5,
            fill: { color: '28295D' }
          });

          slide.addText(transitionTitle, {
            x: 1, y: 2.2, w: 5, h: 1.1,
            fontSize: 36, color: validateColor('FFFFFF'), fontFace: 'Lato',
            align: 'left', bold: true, valign: 'middle'
          });
        }

        // Add brand logo (2:1 aspect ratio) - top right for both
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

        console.log(`Created ${isAltTransition ? 'alternate' : 'standard'} transition slide: "${slideData.title}"`);

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