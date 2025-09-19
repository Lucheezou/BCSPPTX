const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs-extra');
const mammoth = require('mammoth');
const OpenAI = require('openai');
const Groq = require('groq-sdk');
const { encoding_for_model } = require('tiktoken');
const PptxGenJS = require('pptxgenjs');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Initialize Groq client
const groq = new Groq({
  apiKey: process.env.GROQ_API_KEY
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
    // First chunk gets title page + agenda + content slides
    prompt = `You are an expert at converting document content into comprehensive presentation slides.

Document content:
"""
${chunk}
"""

Template styles to use:
"""
${templateStyles}
"""

Create a complete presentation starting with:

1. TITLE PAGE: Use the exact structure from template (div.slide with background-image, overlay, content with title/subtitle, and brand-logo)
   - Extract a compelling title from the document content
   - Add an appropriate subtitle
   - Keep all styling classes exactly as in template
   - Use: <img src="/assets/image8.png" alt="Brand Logo"> for the logo

2. AGENDA PAGE: Use template structure (div.agenda-slide with agenda-left/agenda-right layout)
   - Create agenda items based on document topics/sections
   - Use simple ✓ symbols for checkmarks: <span class="agenda-checkmark">✓</span>
   - Include page number "2"
   - Use: <img src="/assets/image8.png" alt="BCS Logo"> for the logo

3. CONTENT SLIDES: Create 2-4 content slides using div.content-slide structure
   - Extract main topics from document content
   - Use content-header with titles and content-body with paragraphs
   - Include appropriate page numbers starting from 3
   - Use: <img src="assets/image8.png" alt="BCS Logo"> for all logos

CRITICAL FORMATTING REQUIREMENTS:
- Return ONLY raw HTML elements - no markdown, no code blocks, no explanations
- Do NOT use code blocks or backticks anywhere in your response
- Do NOT number slides or add section headers like "1. Title Page:"
- Start immediately with <div class="slide"> for the first element
- Use exact CSS classes from template
- Use "/assets/image8.png" for ALL logo references
- Use simple ✓ for checkmarks, not FontAwesome or other icons

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

  const completion = await groq.chat.completions.create({
    model: "openai/gpt-oss-120b",
    messages: [{ role: "user", content: prompt }],
    max_tokens: 65536,
    temperature: 0.7,
  });

  return completion.choices[0].message.content;
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
  res.json({ message: 'Document to PPT Converter API' });
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

    // Process entire document at once with Groq's high token limit
    console.log(`Processing entire document with Groq`);

    const prompt = `You are an expert at converting document content into comprehensive presentation slides.

Document content:
"""
${documentText}
"""

Template styles to use:
"""
${templateStyles}
"""

Create a complete presentation starting with:

1. TITLE PAGE: Use the exact structure from template (div.slide with background-image, overlay, content with title/subtitle, and brand-logo)
   - Extract a compelling title from the document content
   - Add an appropriate subtitle
   - Keep all styling classes exactly as in template
   - Use: <img src="/assets/image8.png" alt="Brand Logo"> for the logo

2. AGENDA PAGE: Use template structure (div.agenda-slide with agenda-left/agenda-right layout)
   - Create agenda items based on document topics/sections
   - Use simple ✓ symbols for checkmarks: <span class="agenda-checkmark">✓</span>
   - Include page number "2"
   - Use: <img src="/assets/image8.png" alt="BCS Logo"> for the logo

3. CONTENT SLIDES: Create multiple content slides using div.content-slide structure
   - Extract all main topics from document content
   - Use content-header with titles and content-body with paragraphs
   - Include appropriate page numbers starting from 3
   - Use: <img src="assets/image8.png" alt="BCS Logo"> for all logos
   - Create as many slides as needed to cover all content comprehensively

CRITICAL FORMATTING REQUIREMENTS:
- Return ONLY raw HTML elements - no markdown, no code blocks, no explanations
- Do NOT use code blocks or backticks anywhere in your response
- Do NOT number slides or add section headers like "1. Title Page:"
- Start immediately with <div class="slide"> for the first element
- Use exact CSS classes from template
- Use "/assets/image8.png" for ALL logo references
- Use simple ✓ for checkmarks, not FontAwesome or other icons

Return ONLY the complete slide HTML elements without html/head/body tags.`;

    const completion = await groq.chat.completions.create({
      model: "openai/gpt-oss-120b",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 65536,
      temperature: 0.7,
    });

    const allSlides = [completion.choices[0].message.content];

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
    console.log('Using AI to extract content from HTML...');

    // Clean the HTML input first to remove problematic characters
    const cleanedHtml = cleanTextContent(html);

    const prompt = `Analyze this HTML presentation content and extract the structured data for PowerPoint conversion.

HTML Content:
"""
${cleanedHtml}
"""

Extract and return ONLY a JSON object with this exact structure:
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
    }
  ]
}

Rules:
1. Extract the actual text content from HTML elements, NOT the default placeholder text
2. For title slide: find the real title and subtitle in div.title and div.subtitle
3. For briefing_header: Always format as "BCS Monthly Briefing: [Month Day, Year]" with the current date
4. For agenda slide: extract all li.agenda-item text content (without the checkmarks)
5. For content slides: extract div.content-title and all p.content-paragraph text
6. Convert HTML entities and special characters to proper text (e.g., &amp; to &, – to -, ' to ')
7. Clean up any encoding issues and ensure all text uses standard ASCII/UTF-8 characters
8. Return ONLY valid JSON, no explanations or markdown
9. If a slide type is not found, don't include it in the array`;

    const completion = await groq.chat.completions.create({
      model: "openai/gpt-oss-120b",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 4096,
      temperature: 0.1,
    });

    const response = completion.choices[0].message.content.trim();
    console.log('AI response:', response.substring(0, 300));

    // Clean and parse JSON response
    const cleanResponse = response.replace(/```json\s*/g, '').replace(/```\s*$/g, '').trim();
    const extractedData = JSON.parse(cleanResponse);

    // Clean up text content to fix encoding issues
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
      const slide = pptx.addSlide();

      console.log(`Processing slide ${i + 1}: ${slideData.type}...`);

      if (slideData.type === 'title') {
        // Title slide with background image - simple sizing (works best with 16:9 images)
        slide.addImage({
          path: 'public/assets/image9.jpg',
          x: 0, y: 0, w: 10, h: 5.625
        });

        // Add purple overlay
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 5.625,
          fill: { color: '28295d', transparency: 20 }
        });

        // Add BCS Monthly Briefing header
        const briefingHeader = slideData.briefing_header || `BCS Monthly Briefing: ${new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })}`;
        slide.addText(briefingHeader, {
          x: 0.5, y: 1.2, w: 8.5, h: 0.6,
          fontSize: 24, color: 'FFFFFF', fontFace: 'Lato',
          align: 'left', valign: 'middle', bold: true
        });

        // Main title - smaller font size
        slide.addText(slideData.title, {
          x: 0.5, y: 2.0, w: 8.5, h: 1.5,
          fontSize: 44, color: 'FFFFFF', fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle',
          lineSpacing: 50
        });

        if (slideData.subtitle) {
          slide.addText(slideData.subtitle, {
            x: 0.5, y: 3.7, w: 8.5, h: 1,
            fontSize: 24, color: 'FFFFFF', fontFace: 'Lato',
            align: 'left', transparency: 10, valign: 'middle'
          });
        }

        // Add brand logo (bigger, adjusted position)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8, y: 3.8, w: 2, h: 2
        });

        console.log(`Created title slide: "${slideData.title}"`);

      } else if (slideData.type === 'agenda') {
        // Agenda slide with background image - simple sizing (works best with 16:9 images)
        slide.addImage({
          path: 'public/assets/image7.jpg',
          x: 0, y: 0, w: 10, h: 5.625
        });

        // Add purple overlay
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 5.625,
          fill: { color: '28295d', transparency: 20 }
        });

        // Left side - AGENDA title (aligned left)
        slide.addText('AGENDA', {
          x: 0.4, y: 2.2, w: 3.5, h: 1.2,
          fontSize: 54, color: 'FFFFFF', fontFace: 'Lato',
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
            fontSize: 20, color: '7cb342', fontFace: 'Lato',
            align: 'center', bold: true, valign: 'middle'
          });

          // Add agenda item text - truncate if too long
          const itemText = item.length > 60 ? item.substring(0, 57) + '...' : item;
          slide.addText(itemText, {
            x: 4.8, y: yPos, w: 4.8, h: 0.5,
            fontSize: 20, color: '28295d', fontFace: 'Lato',
            align: 'left', valign: 'middle'
          });
          yPos += 0.6;
        }

        // Add brand logo on lower left (bigger, adjusted position)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 0.2, y: 3.8, w: 1.8, h: 1.8
        });

        // Page number
        slide.addText('2', {
          x: 9.2, y: 5, w: 0.6, h: 0.4,
          fontSize: 20, color: '28295d', fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        console.log(`Created agenda slide with ${slideData.items.length} items`);

      } else if (slideData.type === 'content') {
        // Content slide
        slide.background = { color: 'F5F5F5' };

        // Purple header
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 1.1,
          fill: { color: '28295d' }
        });

        // Content title - truncate if too long
        const titleText = slideData.title.length > 80 ? slideData.title.substring(0, 77) + '...' : slideData.title;
        slide.addText(titleText, {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: 28, color: 'FFFFFF', fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle'
        });

        // Add brand logo in header (bigger, slight adjustment up)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8.5, y: -0.2, w: 1.5, h: 1.5
        });

        // Content paragraphs - improved text handling to prevent overflow
        const allContent = slideData.content.join('\n\n');

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

          // Add first chunk to current slide with better dimensions
          slide.addText(chunks[0], {
            x: 0.6, y: 1.5, w: 8.8, h: 3.3,
            fontSize: 16, color: '333333', fontFace: 'Lato',
            align: 'left', lineSpacing: 20, valign: 'top',
            wrap: true
          });

          // Add page number
          slide.addText(`${i + 1}`, {
            x: 9, y: 5.1, w: 0.8, h: 0.4,
            fontSize: 18, color: '28295d', fontFace: 'Lato',
            align: 'center', valign: 'middle'
          });

          // Create additional slides for remaining chunks
          for (let c = 1; c < chunks.length; c++) {
            const contSlide = pptx.addSlide();
            contSlide.background = { color: 'F5F5F5' };

            contSlide.addShape(pptx.ShapeType.rect, {
              x: 0, y: 0, w: 10, h: 1.1,
              fill: { color: '28295d' }
            });

            const contTitleText = titleText.length > 65 ? titleText.substring(0, 62) + '...' : titleText;
            contSlide.addText(`${contTitleText} (cont.)`, {
              x: 0.5, y: 0.15, w: 7, h: 0.8,
              fontSize: 28, color: 'FFFFFF', fontFace: 'Lato',
              align: 'left', bold: false, valign: 'middle'
            });

            // Add brand logo to continuation slide (bigger, slight adjustment up)
            contSlide.addImage({
              path: 'public/assets/image8.png',
              x: 8.5, y: -0.2, w: 1.5, h: 1.5
            });

            contSlide.addText(chunks[c], {
              x: 0.6, y: 1.5, w: 8.8, h: 3.3,
              fontSize: 16, color: '333333', fontFace: 'Lato',
              align: 'left', lineSpacing: 20, valign: 'top',
              wrap: true
            });

            contSlide.addText(`${i + 1}-${c + 1}`, {
              x: 9, y: 5.1, w: 0.8, h: 0.4,
              fontSize: 18, color: '28295d', fontFace: 'Lato',
              align: 'center', valign: 'middle'
            });
          }
        } else {
          // Content fits on one slide
          slide.addText(allContent, {
            x: 0.6, y: 1.5, w: 8.8, h: 3.3,
            fontSize: 16, color: '333333', fontFace: 'Lato',
            align: 'left', lineSpacing: 20, valign: 'top',
            wrap: true
          });

          // Add page number
          slide.addText(`${i + 1}`, {
            x: 9, y: 5.1, w: 0.8, h: 0.4,
            fontSize: 18, color: '28295d', fontFace: 'Lato',
            align: 'center', valign: 'middle'
          });
        }

        console.log(`Created content slide: "${slideData.title}" with ${slideData.content.length} paragraphs`);
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