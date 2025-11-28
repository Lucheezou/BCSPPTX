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
- Use TRANSITION slides between article categories - use transition_alt type for ALL transitions with rotating background images
  * "In the News" - use transition_alt
  * "Federal Updates" - use transition_alt
  * "Hot Topics" - use transition_alt
  * "Question of the Month" - use transition_alt
- Use THANK YOU as final slide for presentations (MANDATORY - always include with disclaimer)
- Use standard CONTENT slides for regular text content - FORMAT AS CONCISE BULLETED LISTS (max 5 bullets, max 12 words each)

CONTENT DENSITY & SLIDE COUNT RULES:
- Maximum 5 bullets per slide
- Maximum 12 words per bullet
- DO NOT ABBREVIATE - spell out all terms in full (use "Question of the Month" not "QotM", write out acronyms)
- Classify each article by importance:
  * CRITICAL: 2 slides max (overview + employer actions with quick check)
  * STANDARD: 1 slide max (overview + actions combined)
  * MINOR: Include as 1-2 bullets in a section roundup slide
- Use importance cues: "Applies to", employer actions, dates, litigation impact
- Push detailed explanations, examples, and context to speaker notes (not visible on slides)
- Keep slides concise for presenter to speak; put supporting detail in notes

CREATE PRESENTATION:
1. Start with TITLE PAGE (extract compelling title from document)
2. Add AGENDA PAGE with ALL article titles (not just section headers):
   - Include transition headers as main items: "In the News", "Federal Updates", "Hot Topics", "Question of the Month"
   - Under each section, list ALL article titles as sub-items with "  " prefix for indentation
   - Example: "In the News", "  Trial Court Vacates Exemptions", "  Tobacco Incentive Lawsuits", etc.
3. For EACH ARTICLE, determine importance level and create slides accordingly:

   CRITICAL articles (2 slides max with SECTION HEADINGS):

   a) SLIDE 1 - Overview with CENTERED section headings (NO bullets on headings):
      • Section headings WITHOUT bullets will be centered, bold, purple
      • Typical sections: "Background and Timeline", "What is changing and why?"
      • After each heading, add 2-4 regular bullets (12 words each)
      • Use "  -" prefix for sub-bullets (indented)
      • CRITICAL: In JSON content array, headings have NO bullet, regular items have "• ", sub-items have "  -"
      JSON Example:
        "content": [
          "Background and Timeline",
          "• ACA required contraceptive coverage",
          "• 2017 rules expanded exemptions",
          "What is changing and why?",
          "• Trial court vacated 2017 exemptions",
          "  - Did not reasonably address stated problem"
        ]

   b) SLIDE 2 - Employer Actions (CHECKLIST format):
      • Start with <h3 class="content-section-heading">What's next?</h3> (NO bullet) + brief context paragraph
      • Then <h2 class="checklist-content-heading">Employer Action Steps</h2> (NO colon, NO bullet, centered)
      • List 5-8 action items with bullets
      • IMPORTANT: Headings have NO bullet character, only action items have bullets
      Example:
        <h3 class="content-section-heading">What's next?</h3>
        <p style="margin-bottom:15px;font-size:16px;">Appeal expected to Third Circuit</p>
        <h2 class="checklist-content-heading">Employer Action Steps</h2>
        <p class="checklist-item">• Objecting employers seek counsel</p>
        <p class="checklist-item">• Stop relying on 2017 exemptions</p>

   STANDARD articles (1 slide max):
   a) COMBINED slide (overview + actions, max 5 bullets, max 12 words each)

   MINOR articles:
   - Roll up into a single roundup slide with 1-2 bullets per item (max 10 words each)

   QUESTION OF THE MONTH (special case-style format):
   a) Use QotM structure - DO NOT include section header text as bullets:
      - Question as title
      - "scenario" array: ONLY the scenario facts (2-3 bullets), NOT "Scenario" as a bullet
      - "rule" array: ONLY what the rule says (2-3 bullets), NOT "What the rule says" as a bullet
      - "action" array: ONLY employer actions (2-3 bullets), NOT "What employers should do" as a bullet
      - Section headers render automatically with styled colored boxes
   - NO "Go Deeper/Who Applies/Employer Implications" pattern for QotM

4. Between article categories, add TRANSITION slides:
   - "In the News" → transition-slide
   - "Federal Updates" → transition-alt-slide (same visual style as Hot Topics and In the News)
   - "Hot Topics" → transition-slide
   - "Question of the Month" → transition-alt-slide
5. Intelligently select appropriate components for document content:
   - Tables/data → TABLE LAYOUT (ensure proper auto-resize)
   - DO NOT use STATISTICS slides
   - Requirements/tasks → CHECKLIST (sidebar format)
   - Employer implications → CHECKLIST SIDEBAR (highly useful - always retain)
   - Important highlights → GRAY TEXT BOXES
   - Regular content → CONTENT SLIDE with CONCISE BULLETED LISTS (max 5 bullets, max 12 words each)
6. SPEAKER NOTES REQUIREMENT (CRITICAL):
   🚨 MANDATORY HTML STRUCTURE FOR NOTES - DO NOT USE PLAIN TEXT PARAGRAPHS! 🚨

   - For every content, go_deeper, checklist, table, and qotm slide, ADD A HIDDEN NOTES SECTION
   - Add this IMMEDIATELY AFTER the slide's closing </div>
   - Inside the slide-notes div, you MUST use proper HTML tags:
     * <p>text here</p> for each paragraph
     * <ul><li>item</li><li>item</li></ul> for bulleted lists
     * <strong>Label:</strong> for section headings
   - DO NOT write plain text without HTML tags - it will become one massive unreadable paragraph!
   - DO NOT use • bullet characters in plain text - use <ul><li> tags instead!
   - CORRECT Example (MUST use HTML tags):
     <div class="slide-notes" data-slide-type="content" style="display:none;">
       <p>Appeal expected to the Third Circuit; until then, expanded exemptions cannot be relied upon</p>

       <p><strong>Applies to:</strong> Employers sponsoring a non-grandfathered medical plan who object to covering one or more of the women's contraceptive/sterilization benefits.</p>

       <ul>
         <li>Objecting employers without another explicit exemption should consult counsel on next steps</li>
         <li>Identify if your plan relied on 2017 religious/moral exemptions; assume they are unavailable pending appeal</li>
         <li>Assess eligibility for 2013 outright exemption (limited religious entities) or accommodation process</li>
         <li>If pursuing accommodation, coordinate EBSA Form 700 or DOL/HHS notice</li>
       </ul>

       <p><strong>Background context:</strong></p>
       <p>ACA requires non-grandfathered plans to cover women's sterilization and contraceptives without cost-sharing.</p>

       <p><strong>Court's critique:</strong></p>
       <ul>
         <li>Religious exemption: did not reasonably address stated problem</li>
         <li>Moral exemption: relied on improper factors not authorized by Congress</li>
       </ul>
     </div>

   - WRONG Example (DO NOT generate plain text like this):
     <div class="slide-notes" style="display:none;">
       Plain text paragraph without HTML tags will create one massive unreadable block.
     </div>
   NOTES FORMATTING REQUIREMENTS (CRITICAL):
   - MUST use proper HTML: <p> for paragraphs, <ul><li> for bulleted lists, <strong> for bold labels
   - DO NOT use plain text with • bullet characters - use <ul><li> tags
   - <p> tags create blank lines between sections in PowerPoint notes
   - <ul><li> tags convert to bulleted lists (• item) in PowerPoint notes
   - This HTML structure converts to properly formatted notes with line breaks and structure

7. End with THANK YOU slide (MANDATORY - MUST copy this EXACT structure from template):
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

// ===== ARTICLE PARSING HELPER =====
// Parses document text to extract individual articles with their metadata
function parseArticlesFromDocument(documentText) {
  const articles = [];

  // Try to split by numbered articles first
  const numberedSplit = documentText.split(/\n(?=\d+\.\s+)/);

  if (numberedSplit.length > 1) {
    // Document has numbered articles
    numberedSplit.forEach(chunk => {
      const titleMatch = chunk.match(/^\d+\.\s+([^\n]+)/);
      if (titleMatch) {
        const title = titleMatch[1].trim();
        const content = chunk.substring(titleMatch[0].length).trim();

        // Extract metadata hints from content
        const metadata = {};
        if (content.match(/\[CRITICAL\]/i)) metadata.importance_hint = 'critical';
        if (content.match(/\[STANDARD\]/i)) metadata.importance_hint = 'standard';
        if (content.match(/\[MINOR\]/i)) metadata.importance_hint = 'minor';

        articles.push({ title, content, metadata });
      }
    });
  } else {
    // Fallback: treat entire document as one article
    articles.push({
      title: 'Document Content',
      content: documentText,
      metadata: {}
    });
  }

  console.log(`📄 Parsed ${articles.length} articles from document`);
  return articles;
}

// ===== CLASSIFICATION GUIDANCE BUILDER =====
// Builds structured guidance for AI based on classification results
function buildClassificationGuidance(classificationResults) {
  const critical = classificationResults.filter(r => r.importance === 'critical');
  const standard = classificationResults.filter(r => r.importance === 'standard');
  const minor = classificationResults.filter(r => r.importance === 'minor');

  let guidance = '\n\n🎯 ARTICLE CLASSIFICATION & SLIDE COUNT ENFORCEMENT:\n\n';

  if (critical.length > 0) {
    guidance += 'CRITICAL ARTICLES (2 slides max each):\n';
    critical.forEach(item => {
      guidance += `  - "${item.title}": Create overview slide + employer implications checklist (MAX 2 SLIDES)\n`;
    });
    guidance += '\n';
  }

  if (standard.length > 0) {
    guidance += 'STANDARD ARTICLES (1 slide max each):\n';
    standard.forEach(item => {
      guidance += `  - "${item.title}": Create single combined slide with overview + actions (MAX 1 SLIDE)\n`;
    });
    guidance += '\n';
  }

  if (minor.length > 0) {
    guidance += 'MINOR ITEMS (rollup into section roundup slides):\n';
    minor.forEach(item => {
      guidance += `  - "${item.title}": Include as 1-2 bullets in section roundup (NO DEDICATED SLIDES)\n`;
    });
    guidance += '\n';
    guidance += '⚠️ Create section roundup slides that combine multiple minor items:\n';
    guidance += '   - Group minor items by section (In the News, Federal Updates, Hot Topics)\n';
    guidance += '   - Each minor item gets 1-2 bullets max (10 words each)\n';
    guidance += '   - Title: "[Section Name] - Quick Updates" or "[Section Name] Roundup"\n';
    guidance += '\n';
  }

  guidance += '🚨 HARD ENFORCEMENT RULES:\n';
  guidance += '  - CRITICAL articles: NEVER exceed 2 slides\n';
  guidance += '  - STANDARD articles: NEVER exceed 1 slide\n';
  guidance += '  - MINOR items: NEVER create dedicated slides, ONLY rollup bullets\n';
  guidance += '  - If content exceeds slide limits, move detail to speaker notes\n';
  guidance += '  - Prioritize conciseness over completeness on slides\n';

  return guidance;
}

// ===== SLIDE COUNT ENFORCEMENT =====
// Validates and enforces slide count caps based on article classification
function enforceSlideCountCaps(slides, classificationResults) {
  console.log('🔍 Enforcing slide count caps...');

  // Build article title → classification map for quick lookup
  const classificationMap = new Map();
  classificationResults.forEach(result => {
    // Normalize title for fuzzy matching (lowercase, remove special chars)
    const normalizedTitle = result.title.toLowerCase().replace(/[^a-z0-9\s]/g, '');
    classificationMap.set(normalizedTitle, result);
  });

  // Track slides per article
  const articleSlideGroups = new Map();

  // Group slides by article (assumes consecutive slides belong to same article)
  const validatedSlides = [];

  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];

    // Skip structural slides (always include)
    if (['title', 'agenda', 'transition', 'transition-alt', 'thankyou'].includes(slide.type)) {
      validatedSlides.push(slide);
      continue;
    }

    // Try to identify which article this slide belongs to
    const slideTitle = (slide.title || '').toLowerCase().replace(/[^a-z0-9\s]/g, '');

    // Check if this slide title matches any classified article
    let matchedArticle = null;
    for (const [articleTitle, classification] of classificationMap) {
      // Fuzzy match: if slide title contains significant portion of article title or vice versa
      const words = articleTitle.split(/\s+/).filter(w => w.length > 3);
      const matchCount = words.filter(word => slideTitle.includes(word)).length;

      if (matchCount >= Math.min(2, words.length)) {
        matchedArticle = classification;
        currentArticle = classification.title;
        break;
      }
    }

    // If matched, enforce slide count for this article
    if (matchedArticle) {
      const articleKey = matchedArticle.title;

      if (!articleSlideGroups.has(articleKey)) {
        articleSlideGroups.set(articleKey, []);
      }

      const articleSlides = articleSlideGroups.get(articleKey);
      const allowedCount = matchedArticle.slideCount;

      if (articleSlides.length < allowedCount) {
        // Within limit, include slide
        articleSlides.push(slide);
        validatedSlides.push(slide);
        console.log(`   ✅ Slide "${slide.title?.substring(0, 30)}..." → Article "${articleKey.substring(0, 30)}..." (${articleSlides.length}/${allowedCount})`);
      } else {
        // Exceeds limit, skip slide
        console.log(`   ⚠️ SKIPPED slide "${slide.title?.substring(0, 30)}..." → Exceeds ${matchedArticle.importance.toUpperCase()} limit (${allowedCount} slides)`);
      }
    } else {
      // Unknown article or content slide - include it (don't over-filter)
      validatedSlides.push(slide);
      console.log(`   ℹ️ Slide "${slide.title?.substring(0, 30)}..." → No article match, including`);
    }
  }

  console.log(`🔍 Enforcement complete: ${slides.length} → ${validatedSlides.length} slides`);
  return validatedSlides;
}

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

    // ===== IMPORTANCE CLASSIFICATION STEP =====
    // Parse document to identify articles and classify them
    console.log('📊 Classifying articles by importance...');

    const articles = parseArticlesFromDocument(documentText);
    const classificationResults = articles.map(article => {
      const importance = classifyArticleImportance(article);
      return {
        title: article.title,
        importance: importance,
        slideCount: importance === 'critical' ? 2 : importance === 'standard' ? 1 : 0 // 0 means roundup
      };
    });

    console.log(`📊 Classified ${classificationResults.length} articles:`);
    classificationResults.forEach(result => {
      console.log(`   - "${result.title.substring(0, 40)}..." → ${result.importance.toUpperCase()} (${result.slideCount} slides)`);
    });

    // Build classification guidance for AI
    const classificationGuidance = buildClassificationGuidance(classificationResults);

    // Read template HTML and extract styles
    const templateHtml = await fs.readFile('template.html', 'utf8');
    const styleMatch = templateHtml.match(/<style>([\s\S]*?)<\/style>/);
    let templateStyles = styleMatch ? styleMatch[1] : '';

    // Update image paths in styles to use absolute paths from server root
    templateStyles = templateStyles.replace(/url\('desiredresults\/assets\/ppt\/media\//g, "url('/assets/");

    // Process entire document at once with GPT-5
    console.log(`Processing entire document with GPT-5 using classification data`);

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

${classificationGuidance}

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
- Use TRANSITION slides between article categories - use transition_alt type for ALL transitions with rotating background images
  * "In the News" - use transition_alt
  * "Federal Updates" - use transition_alt
  * "Hot Topics" - use transition_alt
  * "Question of the Month" - use transition_alt
- Use THANK YOU as final slide for presentations (MANDATORY - always include with disclaimer)
- Use standard CONTENT slides for regular text content - FORMAT AS CONCISE BULLETED LISTS (max 5 bullets, max 12 words each)

CONTENT DENSITY & SLIDE COUNT RULES:
- Maximum 5 bullets per slide
- Maximum 12 words per bullet
- DO NOT ABBREVIATE - spell out all terms in full (use "Question of the Month" not "QotM", write out acronyms)
- Classify each article by importance:
  * CRITICAL: 2 slides max (overview + employer actions with quick check)
  * STANDARD: 1 slide max (overview + actions combined)
  * MINOR: Include as 1-2 bullets in a section roundup slide
- Use importance cues: "Applies to", employer actions, dates, litigation impact
- Push detailed explanations, examples, and context to speaker notes (not visible on slides)
- Keep slides concise for presenter to speak; put supporting detail in notes

CREATE PRESENTATION:
1. Start with TITLE PAGE (extract compelling title from document)
2. Add AGENDA PAGE with ALL article titles (not just section headers):
   - Include transition headers as main items: "In the News", "Federal Updates", "Hot Topics", "Question of the Month"
   - Under each section, list ALL article titles as sub-items with "  " prefix for indentation
   - Example: "In the News", "  Trial Court Vacates Exemptions", "  Tobacco Incentive Lawsuits", etc.
3. For EACH ARTICLE, determine importance level and create slides accordingly:

   CRITICAL articles (2 slides max with SECTION HEADINGS):

   a) SLIDE 1 - Overview with CENTERED section headings (NO bullets on headings):
      • Section headings WITHOUT bullets will be centered, bold, purple
      • Typical sections: "Background and Timeline", "What is changing and why?"
      • After each heading, add 2-4 regular bullets (12 words each)
      • Use "  -" prefix for sub-bullets (indented)
      • CRITICAL: In JSON content array, headings have NO bullet, regular items have "• ", sub-items have "  -"
      JSON Example:
        "content": [
          "Background and Timeline",
          "• ACA required contraceptive coverage",
          "• 2017 rules expanded exemptions",
          "What is changing and why?",
          "• Trial court vacated 2017 exemptions",
          "  - Did not reasonably address stated problem"
        ]

   b) SLIDE 2 - Employer Actions (CHECKLIST format):
      • Start with <h3 class="content-section-heading">What's next?</h3> (NO bullet) + brief context paragraph
      • Then <h2 class="checklist-content-heading">Employer Action Steps</h2> (NO colon, NO bullet, centered)
      • List 5-8 action items with bullets
      • IMPORTANT: Headings have NO bullet character, only action items have bullets
      Example:
        <h3 class="content-section-heading">What's next?</h3>
        <p style="margin-bottom:15px;font-size:16px;">Appeal expected to Third Circuit</p>
        <h2 class="checklist-content-heading">Employer Action Steps</h2>
        <p class="checklist-item">• Objecting employers seek counsel</p>
        <p class="checklist-item">• Stop relying on 2017 exemptions</p>

   STANDARD articles (1 slide max):
   a) COMBINED slide (overview + actions, max 5 bullets, max 12 words each)

   MINOR articles:
   - Roll up into a single roundup slide with 1-2 bullets per item (max 10 words each)

   QUESTION OF THE MONTH (special case-style format):
   a) Use QotM structure - DO NOT include section header text as bullets:
      - Question as title
      - "scenario" array: ONLY the scenario facts (2-3 bullets), NOT "Scenario" as a bullet
      - "rule" array: ONLY what the rule says (2-3 bullets), NOT "What the rule says" as a bullet
      - "action" array: ONLY employer actions (2-3 bullets), NOT "What employers should do" as a bullet
      - Section headers render automatically with styled colored boxes
   - NO "Go Deeper/Who Applies/Employer Implications" pattern for QotM

4. Between article categories, add TRANSITION slides:
   - "In the News" → transition-slide
   - "Federal Updates" → transition-alt-slide (same visual style as Hot Topics and In the News)
   - "Hot Topics" → transition-slide
   - "Question of the Month" → transition-alt-slide
5. Intelligently select appropriate components for document content:
   - Tables/data → TABLE LAYOUT (ensure proper auto-resize)
   - DO NOT use STATISTICS slides
   - Requirements/tasks → CHECKLIST (sidebar format)
   - Employer implications → CHECKLIST SIDEBAR (highly useful - always retain)
   - Important highlights → GRAY TEXT BOXES
   - Regular content → CONTENT SLIDE with CONCISE BULLETED LISTS (max 5 bullets, max 12 words each)
6. SPEAKER NOTES REQUIREMENT (CRITICAL):
   🚨 MANDATORY HTML STRUCTURE FOR NOTES - DO NOT USE PLAIN TEXT PARAGRAPHS! 🚨

   - For every content, go_deeper, checklist, table, and qotm slide, ADD A HIDDEN NOTES SECTION
   - Add this IMMEDIATELY AFTER the slide's closing </div>
   - Inside the slide-notes div, you MUST use proper HTML tags:
     * <p>text here</p> for each paragraph
     * <ul><li>item</li><li>item</li></ul> for bulleted lists
     * <strong>Label:</strong> for section headings
   - DO NOT write plain text without HTML tags - it will become one massive unreadable paragraph!
   - DO NOT use • bullet characters in plain text - use <ul><li> tags instead!
   - CORRECT Example (MUST use HTML tags):
     <div class="slide-notes" data-slide-type="content" style="display:none;">
       <p>Appeal expected to the Third Circuit; until then, expanded exemptions cannot be relied upon</p>

       <p><strong>Applies to:</strong> Employers sponsoring a non-grandfathered medical plan who object to covering one or more of the women's contraceptive/sterilization benefits.</p>

       <ul>
         <li>Objecting employers without another explicit exemption should consult counsel on next steps</li>
         <li>Identify if your plan relied on 2017 religious/moral exemptions; assume they are unavailable pending appeal</li>
         <li>Assess eligibility for 2013 outright exemption (limited religious entities) or accommodation process</li>
         <li>If pursuing accommodation, coordinate EBSA Form 700 or DOL/HHS notice</li>
       </ul>

       <p><strong>Background context:</strong></p>
       <p>ACA requires non-grandfathered plans to cover women's sterilization and contraceptives without cost-sharing.</p>

       <p><strong>Court's critique:</strong></p>
       <ul>
         <li>Religious exemption: did not reasonably address stated problem</li>
         <li>Moral exemption: relied on improper factors not authorized by Congress</li>
       </ul>
     </div>

   - WRONG Example (DO NOT generate plain text like this):
     <div class="slide-notes" style="display:none;">
       Plain text paragraph without HTML tags will create one massive unreadable block.
     </div>
   NOTES FORMATTING REQUIREMENTS (CRITICAL):
   - MUST use proper HTML: <p> for paragraphs, <ul><li> for bulleted lists, <strong> for bold labels
   - DO NOT use plain text with • bullet characters - use <ul><li> tags
   - <p> tags create blank lines between sections in PowerPoint notes
   - <ul><li> tags convert to bulleted lists (• item) in PowerPoint notes
   - This HTML structure converts to properly formatted notes with line breaks and structure

7. End with THANK YOU slide (MANDATORY - MUST copy this EXACT structure from template):
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
      presentationId: timestamp,
      classificationResults: classificationResults // Pass classification data to client
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

// Importance classifier for articles
function classifyArticleImportance(articleData) {
  const { title = '', content = '', metadata = {} } = articleData;

  // Manual override takes precedence
  if (metadata.importance_hint) {
    console.log(`📊 Manual importance override: ${metadata.importance_hint} for "${title}"`);
    return metadata.importance_hint;
  }

  const combinedText = (title + ' ' + content).toLowerCase();

  // CRITICAL indicators (score system)
  let criticalScore = 0;
  let standardScore = 0;
  let minorScore = 0;

  // Critical patterns (high impact)
  const criticalPatterns = [
    /applies to\s+all\s+(employers|groups)/i,
    /mandatory|required|must comply/i,
    /deadline.*\d{4}/i,
    /effective\s+date/i,
    /litigation|lawsuit|court\s+(order|ruling)/i,
    /penalty|fine|enforcement/i,
    /new\s+(federal|state)\s+(law|regulation|rule)/i,
    /emergency|urgent/i,
    /action\s+required/i,
    /compliance\s+deadline/i
  ];

  // Standard patterns (moderate impact)
  const standardPatterns = [
    /applies to/i,
    /employer\s+(action|obligations|responsibilities)/i,
    /guidance|clarification|update/i,
    /reporting\s+requirement/i,
    /notice\s+requirement/i,
    /action\s+items/i,
    /compliance\s+quick\s+check/i,
    /takeaways/i
  ];

  // Minor patterns (low impact, informational)
  const minorPatterns = [
    /reminder|note/i,
    /background|context/i,
    /fyi|for\s+your\s+information/i,
    /optional|voluntary/i,
    /best\s+practice/i,
    /tip/i
  ];

  // Score critical indicators
  criticalPatterns.forEach(pattern => {
    if (pattern.test(combinedText)) criticalScore += 2;
  });

  // Score standard indicators
  standardPatterns.forEach(pattern => {
    if (pattern.test(combinedText)) standardScore += 1;
  });

  // Score minor indicators
  minorPatterns.forEach(pattern => {
    if (pattern.test(combinedText)) minorScore += 1;
  });

  // Additional scoring based on content characteristics
  const hasActionItems = /action\s+items?:/i.test(combinedText);
  const hasQuickCheck = /compliance\s+quick\s+check/i.test(combinedText);
  const hasDeadline = /deadline|due\s+date|by\s+\w+\s+\d+/i.test(combinedText);
  const hasAppliesTo = /applies\s+to:/i.test(combinedText);
  const hasPenalty = /penalty|fine|sanction/i.test(combinedText);

  if (hasActionItems) standardScore += 2;
  if (hasQuickCheck) standardScore += 2;
  if (hasDeadline) criticalScore += 1;
  if (hasAppliesTo) standardScore += 1;
  if (hasPenalty) criticalScore += 2;

  // Content length analysis
  const wordCount = combinedText.split(/\s+/).length;
  if (wordCount > 500) criticalScore += 1; // Long content suggests important topic
  if (wordCount < 100) minorScore += 2; // Short content suggests minor update

  // Decision logic - adjusted thresholds for better classification
  let importance;
  if (criticalScore >= 2) {
    // Articles with 2+ critical indicators are critical (court ruling, mandatory compliance, penalties, etc.)
    importance = 'critical';
  } else if (criticalScore >= 1 || standardScore >= 2) {
    // Articles with 1 critical indicator OR 2+ standard indicators are standard
    importance = 'standard';
  } else if (minorScore >= 2 || (criticalScore === 0 && standardScore <= 1)) {
    // Articles with 2+ minor indicators OR no critical/standard indicators are minor
    importance = 'minor';
  } else {
    // Default to standard for edge cases
    importance = 'standard';
  }

  console.log(`📊 Classified "${title.substring(0, 50)}..." as ${importance.toUpperCase()} (critical: ${criticalScore}, standard: ${standardScore}, minor: ${minorScore})`);

  return importance;
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

// WORKAROUND: Auto-format plain text notes into structured HTML
function autoFormatPlainTextNotes(plainText) {
  console.log('🔧 Auto-formatting plain text notes into HTML structure...');

  // Split into sentences and identify structure
  let formatted = plainText.trim();

  // Step 1: Identify and wrap section labels (e.g., "Applies to:", "Background:", etc.)
  formatted = formatted.replace(
    /^([A-Z][^:.!?]*:)/gm,
    '<p><strong>$1</strong></p><p>'
  );

  // Step 2: Split into sentences (periods followed by space and capital letter)
  const sentences = formatted.split(/\.\s+(?=[A-Z])/);

  // Step 3: Identify lists - sentences that are likely part of a list
  const structuredParts = [];
  let currentList = [];

  sentences.forEach((sentence, idx) => {
    // Check if this looks like a list item (starts with action word, has "should", etc.)
    const isListItem = /^(Employers?|Objecting|Review|Update|Consult|Identify|Assess|Coordinate|Train|Monitor|Track|File|Align|Ensure|Note|Consider)/i.test(sentence.trim());

    if (isListItem && sentence.length < 200) {
      // Add to current list
      currentList.push(sentence.trim() + '.');
    } else {
      // Not a list item - flush current list if any
      if (currentList.length > 0) {
        structuredParts.push({
          type: 'list',
          items: currentList
        });
        currentList = [];
      }

      // Add as paragraph
      if (sentence.trim()) {
        structuredParts.push({
          type: 'paragraph',
          text: sentence.trim() + (idx < sentences.length - 1 ? '.' : '')
        });
      }
    }
  });

  // Flush any remaining list
  if (currentList.length > 0) {
    structuredParts.push({
      type: 'list',
      items: currentList
    });
  }

  // Step 4: Convert to HTML
  let html = '';
  structuredParts.forEach(part => {
    if (part.type === 'paragraph') {
      html += `<p>${part.text}</p>\n\n`;
    } else if (part.type === 'list') {
      html += '<ul>\n';
      part.items.forEach(item => {
        html += `  <li>${item}</li>\n`;
      });
      html += '</ul>\n\n';
    }
  });

  console.log(`✅ Auto-formatted ${structuredParts.length} sections from plain text`);
  return html.trim();
}

// AI-powered HTML content extraction
async function extractContentWithAI(html) {
  try {
    console.log('Using GPT-5 to extract content from HTML...');

    // Clean the HTML input first to remove problematic characters
    const cleanedHtml = cleanTextContent(html);

    // PRE-PROCESS: Extract notes sections with regex for guaranteed extraction
    const notesRegex = /<div class="slide-notes"[^>]*>([\s\S]*?)<\/div>/gi;
    let notesMatch;
    const notesArray = [];

    while ((notesMatch = notesRegex.exec(cleanedHtml)) !== null) {
      let notesContent = notesMatch[1];

      // WORKAROUND: If AI generated plain text without HTML tags, auto-format it
      const hasHTMLTags = /<(p|ul|li|strong)>/i.test(notesContent);

      if (!hasHTMLTags && notesContent.length > 100) {
        console.log('⚠️ Detected plain text notes without HTML - auto-formatting...');

        // Auto-format plain text into structured HTML
        notesContent = autoFormatPlainTextNotes(notesContent);
      }

      // Convert HTML formatting to plain text with proper line breaks
      notesContent = notesContent
        .replace(/<br\s*\/?>/gi, '\n')           // Convert <br> to newlines
        .replace(/<\/p>/gi, '\n\n')              // Convert </p> to double newlines
        .replace(/<p[^>]*>/gi, '')               // Remove <p> opening tags
        .replace(/<\/li>/gi, '')                 // Remove </li> tags
        .replace(/<li[^>]*>/gi, '\n• ')          // Convert <li> to newline + bullet
        .replace(/<ul[^>]*>/gi, '\n')            // Add newline before list
        .replace(/<\/ul>/gi, '\n')               // Add newline after list
        .replace(/<strong>/gi, '')               // Remove <strong> tags (keep text)
        .replace(/<\/strong>/gi, '')             // Remove </strong> closing tags
        .replace(/<h[1-6][^>]*>/gi, '\n')        // Add newline before headings
        .replace(/<\/h[1-6]>/gi, '\n')           // Add newline after headings
        .replace(/<[^>]*>/g, '')                 // Remove remaining HTML tags
        .replace(/&bull;/g, '•')                 // Normalize bullets
        .replace(/\n\s*\n\s*\n+/g, '\n\n')       // Collapse multiple blank lines to max 2
        .replace(/\n\s+/g, '\n')                 // Remove leading spaces on lines
        .trim();

      notesArray.push(notesContent);
    }

    console.log(`📝 Found ${notesArray.length} slide-notes sections via regex pre-processing`);
    if (notesArray.length > 0) {
      console.log(`📝 First notes sample: ${notesArray[0].substring(0, 100)}...`);
    }

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
      "items": ["In the News", "  Trial Court Vacates Exemptions", "  Tobacco Lawsuits", "Federal Updates", "  ACA FAQs Part 71", "  STLDI Rules", "Hot Topics", "  MLR Rebates", "Question of the Month"]
    },
    {
      "type": "content",
      "title": "slide title",
      "content": ["• Main point 1", "  ◦ Sub-point 1a", "  ◦ Sub-point 1b", "• Main point 2"],
      "bullets": true,
      "notes": "Detailed speaker notes with full explanations, examples, citations"
    },
    {
      "type": "content",
      "title": "CRITICAL Article Title (2-slide format, slide 1)",
      "content": ["Background and Timeline", "• ACA required coverage", "• 2017 rules expanded", "What is changing and why?", "• Court vacated 2017 rules", "  - Did not address problem"],
      "bullets": true,
      "notes": "Speaker notes..."
    },
    {
      "type": "go_deeper",
      "title": "Go Deeper: [Article Title]",
      "content": ["• Detailed point 1", "• Detailed point 2"],
      "notes": "Speaker notes with context"
    },
    {
      "type": "qotm",
      "title": "Question text as title",
      "scenario": ["Employee requests mid-year spouse removal", "Court order instructs coverage for child"],
      "rule": ["Legal separation rarely changes eligibility unless plan says so", "Orders requiring coverage can trigger mid-year enrollments"],
      "action": ["Follow written procedures and require timely orders", "Coordinate continuation coverage like COBRA"],
      "notes": "Additional context for presenter"
    },
    {
      "type": "table",
      "title": "table slide title",
      "headers": ["col1", "col2", "col3"],
      "rows": [["data1", "data2", "data3"], ["data4", "data5", "data6"]],
      "notes": "Context about the table data"
    },
    {
      "type": "checklist",
      "title": "checklist title",
      "content": ["item 1", "item 2"],
      "checklist_heading": "Action Items",
      "checklist_panel_text": "Brief paragraph text before checklist",
      "checklist_items": [{"text": "item", "checked": false}],
      "notes": "Detailed explanation of action items"
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
5. div.qotm-slide = Question of the Month case-style slide
6. div.table-slide = table slide (extract table headers and data)
7. div.checklist-slide = checklist slide (SIDEBAR FORMAT - use for employer implications)
8. div.textbox-slide = textbox slide (extract textbox content and colors)
9. div.transition-slide OR div.transition-alt-slide = transition slide
10. div.thankyou-slide = thank you slide

IMPORTANT: Employer Implications should ALWAYS use div.checklist-slide (sidebar format), NOT textbox

CHECKLIST MODE:
- All checklist items should have "checked": false by default (empty checkboxes)
- This creates an interactive checklist for users to check off items
- DO NOT pre-check items unless explicitly marked as examples

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

SPEAKER NOTES EXTRACTION (CRITICAL - MANDATORY):
- ALWAYS look for hidden notes sections that appear IMMEDIATELY AFTER each slide
- Pattern: <div class="slide-notes" data-slide-type="..." style="display:none;">notes text here</div>
- Extract the FULL TEXT CONTENT from inside the slide-notes div
- Include this content in the "notes" field for that slide
- Notes field must be a single string (not array) with full sentences and paragraphs

EXAMPLE HTML INPUT:
  <div class="content-slide">
    <ul><li>Bullet 1</li><li>Bullet 2</li></ul>
  </div>
  <div class="slide-notes" data-slide-type="content" style="display:none;">
    The Affordable Care Act requires Applicable Large Employers (50+ full-time employees) to report health coverage information annually. The 2024 forms include updates to Part II, Column F...
  </div>

EXAMPLE JSON OUTPUT:
  {
    "type": "content",
    "title": "Slide Title",
    "content": ["Bullet 1", "Bullet 2"],
    "notes": "The Affordable Care Act requires Applicable Large Employers (50+ full-time employees) to report health coverage information annually. The 2024 forms include updates to Part II, Column F..."
  }

- If no slide-notes div exists, the "notes" field should be omitted or empty string
- CRITICAL: Search the HTML for each slide-notes div and extract its inner text

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
    // Also inject pre-extracted notes from regex
    let notesIndex = 0;
    const cleanedSlides = extractedData.slides.map((slide, slideIndex) => {
      const cleanSlide = { ...slide };

      // Clean title and subtitle
      if (cleanSlide.title) {
        cleanSlide.title = cleanTextContent(cleanSlide.title);
      }
      if (cleanSlide.subtitle) {
        cleanSlide.subtitle = cleanTextContent(cleanSlide.subtitle);
      }

      // INJECT NOTES from regex pre-processing
      // Skip title and agenda slides (first 2), then map notes to content slides
      const slideTypesWithNotes = ['content', 'go_deeper', 'table', 'checklist', 'qotm', 'textbox'];
      if (slideTypesWithNotes.includes(cleanSlide.type)) {
        if (notesIndex < notesArray.length) {
          cleanSlide.notes = notesArray[notesIndex];
          console.log(`📝 Injected notes into slide ${slideIndex} (${cleanSlide.type}): "${cleanSlide.title}"`);
          notesIndex++;
        }
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

      // Clean notes field (if AI extracted them or if we injected them above)
      // NOTE: Do NOT use cleanTextContent() on notes - it strips newlines!
      // Notes need to preserve \n for proper formatting in PowerPoint
      if (cleanSlide.notes) {
        console.log(`✅ Slide "${cleanSlide.title}" has notes (${cleanSlide.notes.length} chars)`);
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
    const { presentationId, html, classificationResults } = req.body;

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

    // ===== ENFORCE SLIDE COUNT CAPS =====
    // Validate that AI respected classification guidance (if classification data is provided)
    let validatedSlides = slides;
    if (classificationResults && Array.isArray(classificationResults) && classificationResults.length > 0) {
      validatedSlides = enforceSlideCountCaps(slides, classificationResults);
      console.log(`✅ Slide count validation: ${slides.length} → ${validatedSlides.length} slides after enforcement`);
    } else {
      console.log('ℹ️ No classification data provided - skipping slide count enforcement');
    }

    // Process each slide using improved extraction
    for (let i = 0; i < validatedSlides.length; i++) {
      const slideData = validatedSlides[i];

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

        // Right side - agenda items with checkmarks - dynamic sizing based on item count
        const totalItems = slideData.items.length;
        const maxItems = Math.min(totalItems, 15); // Allow up to 15 items

        // Dynamic font sizing and spacing based on item count
        let baseFontSize = 15;
        let subFontSize = 13;
        let checkFontSize = 18;
        let baseLineHeight = 0.5;
        let lineSpacing = 16;

        if (totalItems > 12) {
          baseFontSize = 12;
          subFontSize = 11;
          checkFontSize = 14;
          baseLineHeight = 0.38;
          lineSpacing = 14;
        } else if (totalItems > 10) {
          baseFontSize = 13;
          subFontSize = 12;
          checkFontSize = 16;
          baseLineHeight = 0.42;
          lineSpacing = 15;
        }

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
            fontSize: isSubBullet ? checkFontSize - 2 : checkFontSize,
            color: validateColor('7CB342'), fontFace: 'Lato',
            align: 'center', bold: true, valign: 'top'
          });

          // Add agenda item text with wrapping and proper height for multi-line
          const safeItem = (item || '').toString().trim().replace(/^[◦\-\s]*/, '');
          slide.addText(safeItem, {
            x: 4.6 + indent, y: yPos, w: 5.0 - indent, h: baseLineHeight - 0.05,
            fontSize: isSubBullet ? subFontSize : baseFontSize,
            color: validateColor('28295D'), fontFace: 'Lato',
            align: 'left', valign: 'top', bold: false,
            wrap: true,
            lineSpacing: lineSpacing
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

          // Check if this is a section heading - specific known headings for CRITICAL articles
          const sectionHeadingPatterns = [
            /^Background and Timeline/i,
            /^What is changing and why/i,
            /^What changed/i,
            /^What's next/i,
            /^Key Points/i,
            /^Overview/i,
            /^Court Decision/i,
            /^Employer Impact/i,
            /^What happened/i,
            /^Timeline/i,
            /^The Decision/i
          ];
          const isSectionHeading = !trimmed.startsWith('•') && !trimmed.startsWith('-') &&
                                   sectionHeadingPatterns.some(pattern => pattern.test(trimmed));

          const isSubBullet = trimmed.startsWith('-');

          if (isSectionHeading) {
            // Section heading - centered, bold, no bullet, purple color
            const headingHeight = 0.4;
            if (yPos + headingHeight > maxYPos) break;

            slide.addText(trimmed, {
              x: 0.5, y: yPos, w: 9, h: headingHeight,
              fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
              align: 'center', valign: 'middle', bold: true
            });

            yPos += headingHeight + 0.15; // Space after heading
          } else {
            // Regular bullet point
            const text = trimmed.startsWith('•') || trimmed.startsWith('-') ? trimmed.substring(1).trim() : trimmed;
            const indent = isSubBullet ? 0.4 : 0;
            const bulletSymbol = '•'; // Always use solid bullet for both main and sub-bullets

            // Estimate height based on text length
            const textWidth = 8.5 - indent;
            const charsPerLine = textWidth * 12;
            const numLines = Math.ceil(text.length / charsPerLine);
            const itemHeight = Math.max(0.35, numLines * 0.25);

            // Stop if this item won't fit
            if (yPos + itemHeight > maxYPos) break;

            // Add bullet symbol
            slide.addText(bulletSymbol, {
              x: 0.6 + indent, y: yPos, w: 0.2, h: 0.3,
              fontSize: 14, // Same size for both main and sub-bullets
              color: validateColor('333333'), fontFace: 'Lato',
              align: 'left', valign: 'top'
            });

            // Add text
            slide.addText(text, {
              x: 0.9 + indent, y: yPos, w: 8.5 - indent, h: itemHeight,
              fontSize: 14, color: validateColor('333333'), fontFace: 'Lato',
              align: 'left', valign: 'top', wrap: true,
              lineSpacing: safeContent.length >= 4 ? 20 : 18,
              paraSpaceAfter: safeContent.length >= 4 ? 8 : 6
            });

            yPos += itemHeight + 0.05;
          }
        }

        // Add page number
        slide.addText(`${i + 1}`, {
          x: 9, y: 5.1, w: 0.8, h: 0.4,
          fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        // Add speaker notes if present
        if (slideData.notes && slideData.notes.trim()) {
          slide.addNotes(slideData.notes);
        }

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

          // Add text with calculated height and increased spacing for better readability
          slide.addText(text, {
            x: 0.9 + indent, y: yPos, w: 8.5 - indent, h: itemHeight,
            fontSize: 13, color: validateColor('333333'), fontFace: 'Lato',
            align: 'left', valign: 'top', wrap: true,
            lineSpacing: safeContent.length >= 4 ? 18 : 16,  // Increased spacing for dense slides
            paraSpaceAfter: safeContent.length >= 4 ? 8 : 6   // Paragraph spacing for visual separation
          });

          yPos += itemHeight + 0.05; // Add small gap between items
        }

        // Add page number
        slide.addText(`${i + 1}`, {
          x: 9, y: 5.1, w: 0.8, h: 0.4,
          fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        // Add speaker notes if present
        if (slideData.notes && slideData.notes.trim()) {
          slide.addNotes(slideData.notes);
        }

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

            // Add speaker notes to first table slide only
            if (slideNum === 0 && slideData.notes && slideData.notes.trim()) {
              slide.addNotes(slideData.notes);
            }

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

        // Add speaker notes if present
        if (slideData.notes && slideData.notes.trim()) {
          slide.addNotes(slideData.notes);
        }

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

        // Rotate through background images for variety
        const bgImages = ['image1.jpg', 'image5.jpg', 'image7.jpg', 'image9.jpg'];
        const bgIndex = i % bgImages.length;  // Rotate based on slide index
        const bgImage = bgImages[bgIndex];

        // Transition slide - rotating background images
        slide.addImage({
          path: `public/assets/${bgImage}`,
          x: 0, y: 0, w: 10, h: 5.625
        });

        // Add purple overlay
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 5.625,
          fill: { color: '28295D', transparency: 20 }
        });

        // Transition title (all caps)
        const transitionTitle = (slideData.title || 'TRANSITION SLIDE').toUpperCase();

        // Centered white bordered box
        slide.addShape(pptx.ShapeType.rect, {
          x: 2.0, y: 1.69, w: 6.0, h: 2.25,  // Centered box
          fill: { type: 'solid', color: '000000', transparency: 100 },  // Transparent fill
          line: { color: 'FFFFFF', width: 4 }  // White border
        });

        slide.addText(transitionTitle, {
          x: 2.0, y: 1.69, w: 6.0, h: 2.25,
          fontSize: 40, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'center', bold: true, valign: 'middle'
        });

        // Add brand logo (2:1 aspect ratio) - top right
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

        console.log(`Created transition slide with ${bgImage}: "${slideData.title}"`);

      } else if (slideData.type === 'qotm') {
        // Question of the Month - case-style layout
        const slide = pptx.addSlide();
        slide.background = { color: 'F5F5F5' };

        // Purple header
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 1.1,
          fill: { color: '28295D' }
        });

        // Question as title with dynamic font
        const qotmTitle = slideData.title || 'Question of the Month';
        const headerFontSize = getHeaderFontSize(qotmTitle);
        slide.addText(qotmTitle, {
          x: 0.5, y: 0.15, w: 7, h: 0.8,
          fontSize: headerFontSize, color: validateColor('FFFFFF'), fontFace: 'Lato',
          align: 'left', bold: false, valign: 'middle', wrap: true
        });

        // Add brand logo in header (2:1 aspect ratio)
        slide.addImage({
          path: 'public/assets/image8.png',
          x: 8.5, y: 0.2, w: 1.2, h: 0.6
        });

        // Dynamic sizing for QotM based on total content
        const totalBullets = (slideData.scenario?.length || 0) + (slideData.rule?.length || 0) + (slideData.action?.length || 0);
        let bulletFontSize = 13;
        let bulletHeight = 0.4;
        let bulletSpacing = 0.45;
        let sectionHeaderFontSize = 14;

        if (totalBullets > 7) {
          bulletFontSize = 11;
          bulletHeight = 0.35;
          bulletSpacing = 0.38;
          sectionHeaderFontSize = 12;
        } else if (totalBullets > 5) {
          bulletFontSize = 12;
          bulletHeight = 0.38;
          bulletSpacing = 0.42;
          sectionHeaderFontSize = 13;
        }

        let yPos = 1.4;

        // Scenario section
        if (slideData.scenario && slideData.scenario.length > 0) {
          slide.addShape(pptx.ShapeType.rect, {
            x: 0.6, y: yPos, w: 8.8, h: 0.4,
            fill: { color: 'A8D5D5' }
          });
          slide.addText('Scenario', {
            x: 0.8, y: yPos + 0.05, w: 8.4, h: 0.3,
            fontSize: sectionHeaderFontSize, color: validateColor('28295D'), fontFace: 'Lato',
            align: 'left', bold: true, valign: 'middle'
          });
          yPos += 0.5;

          for (const bullet of slideData.scenario.slice(0, 3)) {
            slide.addText('•', {
              x: 0.8, y: yPos, w: 0.2, h: 0.25,
              fontSize: bulletFontSize, color: validateColor('333333'), fontFace: 'Lato',
              align: 'left', valign: 'top'
            });
            slide.addText(bullet, {
              x: 1.1, y: yPos, w: 8.3, h: bulletHeight,
              fontSize: bulletFontSize, color: validateColor('333333'), fontFace: 'Lato',
              align: 'left', valign: 'top', wrap: true
            });
            yPos += bulletSpacing;
          }
        }

        // What the rule says section
        if (slideData.rule && slideData.rule.length > 0) {
          yPos += 0.1;
          slide.addShape(pptx.ShapeType.rect, {
            x: 0.6, y: yPos, w: 8.8, h: 0.4,
            fill: { color: 'E8E8E8' }
          });
          slide.addText('What the rule says', {
            x: 0.8, y: yPos + 0.05, w: 8.4, h: 0.3,
            fontSize: sectionHeaderFontSize, color: validateColor('28295D'), fontFace: 'Lato',
            align: 'left', bold: true, valign: 'middle'
          });
          yPos += 0.5;

          for (const bullet of slideData.rule.slice(0, 3)) {
            slide.addText('•', {
              x: 0.8, y: yPos, w: 0.2, h: 0.25,
              fontSize: bulletFontSize, color: validateColor('333333'), fontFace: 'Lato',
              align: 'left', valign: 'top'
            });
            slide.addText(bullet, {
              x: 1.1, y: yPos, w: 8.3, h: bulletHeight,
              fontSize: bulletFontSize, color: validateColor('333333'), fontFace: 'Lato',
              align: 'left', valign: 'top', wrap: true
            });
            yPos += bulletSpacing;
          }
        }

        // What employers should do section
        if (slideData.action && slideData.action.length > 0) {
          yPos += 0.1;
          slide.addShape(pptx.ShapeType.rect, {
            x: 0.6, y: yPos, w: 8.8, h: 0.4,
            fill: { color: 'A8D5D5' }
          });
          slide.addText('What employers should do', {
            x: 0.8, y: yPos + 0.05, w: 8.4, h: 0.3,
            fontSize: sectionHeaderFontSize, color: validateColor('28295D'), fontFace: 'Lato',
            align: 'left', bold: true, valign: 'middle'
          });
          yPos += 0.5;

          for (const bullet of slideData.action.slice(0, 3)) {
            slide.addText('•', {
              x: 0.8, y: yPos, w: 0.2, h: 0.25,
              fontSize: bulletFontSize, color: validateColor('333333'), fontFace: 'Lato',
              align: 'left', valign: 'top'
            });
            slide.addText(bullet, {
              x: 1.1, y: yPos, w: 8.3, h: bulletHeight,
              fontSize: bulletFontSize, color: validateColor('333333'), fontFace: 'Lato',
              align: 'left', valign: 'top', wrap: true
            });
            yPos += bulletSpacing;
          }
        }

        // Add page number
        slide.addText(`${i + 1}`, {
          x: 9, y: 5.1, w: 0.8, h: 0.4,
          fontSize: 18, color: validateColor('28295D'), fontFace: 'Lato',
          align: 'center', valign: 'middle'
        });

        // Add speaker notes if present
        if (slideData.notes && slideData.notes.trim()) {
          slide.addNotes(slideData.notes);
        }

        console.log(`Created Question of the Month slide: "${slideData.title}"`);

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