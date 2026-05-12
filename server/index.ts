import express from 'express';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { promisify } from 'util';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  Header,
  ImageRun,
  TextWrappingType,
  TextWrappingSide,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom,
} from 'docx';

// @ts-ignore
import libre from 'libreoffice-convert';

const libreConvertAsync = promisify(libre.convert);

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json({ limit: '50mb' }));

// ─── Helper: Parse HTML into docx Paragraphs ───────────────────────────────
function htmlToDocxParagraphs(html: string): Paragraph[] {
  // Create a simple parser for server-side (no DOM)
  const paragraphs: Paragraph[] = [];

  // Strip the HTML into blocks by splitting on block-level tags
  // We'll handle: <h1>, <h2>, <h3>, <p>, <li>
  const blockRegex = /<(h[1-6]|p|li)([^>]*)>([\s\S]*?)<\/\1>/gi;
  let match;

  while ((match = blockRegex.exec(html)) !== null) {
    const tagName = match[1].toLowerCase();
    const attributes = match[2] || '';
    const innerHtml = match[3];

    // Determine alignment
    const isCenter = attributes.includes('ql-align-center') || attributes.includes('text-align: center') || attributes.includes('text-align:center');
    const isRight = attributes.includes('ql-align-right') || attributes.includes('text-align: right') || attributes.includes('text-align:right');
    const isJustify = attributes.includes('ql-align-justify') || attributes.includes('text-align: justify') || attributes.includes('text-align:justify');

    const alignment = isCenter ? AlignmentType.CENTER :
                      isRight ? AlignmentType.RIGHT :
                      isJustify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT;

    // Parse inline content to create TextRun children
    const children = parseInlineContent(innerHtml, tagName);

    // Skip empty paragraphs that are just <br>
    const textContent = innerHtml.replace(/<[^>]*>?/gm, '').trim();
    if (!textContent && innerHtml.includes('<br>')) {
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: '', size: 22 })],
        spacing: { after: 200 },
        alignment,
      }));
      continue;
    }

    const paragraphOptions: any = {
      children,
      spacing: { after: 200 },
      alignment,
    };

    if (tagName === 'li') {
      paragraphOptions.bullet = { level: 0 };
    }

    paragraphs.push(new Paragraph(paragraphOptions));
  }

  // If no block elements found, treat the entire HTML as plain text
  if (paragraphs.length === 0) {
    const plainText = html.replace(/<[^>]*>?/gm, '').replace(/&nbsp;/gi, ' ');
    paragraphs.push(new Paragraph({
      children: [new TextRun({ text: plainText, size: 22 })],
    }));
  }

  return paragraphs;
}

// ─── Helper: Parse inline HTML content into TextRun array ───────────────────
function parseInlineContent(html: string, parentTag: string): TextRun[] {
  const runs: TextRun[] = [];
  const isHeading = parentTag.startsWith('h');
  const headingSizes: Record<string, number> = {
    'h1': 32, 'h2': 28, 'h3': 24, 'h4': 22, 'h5': 20, 'h6': 18
  };
  const baseFontSize = isHeading ? (headingSizes[parentTag] || 22) : 22;

  // Split by inline tags to preserve formatting
  // We'll process <strong>, <b>, <em>, <i>, and plain text
  const parts = splitInlineHtml(html);

  for (const part of parts) {
    const text = part.text.replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&quot;/gi, '"');
    if (!text) continue;

    const options: any = {
      text,
      size: baseFontSize,
      font: 'Cormorant Garamond',
    };

    if (part.bold || isHeading) options.bold = true;
    if (part.italic) options.italics = true;
    if (part.underline) options.underline = {};

    runs.push(new TextRun(options));
  }

  if (runs.length === 0) {
    runs.push(new TextRun({ text: '', size: baseFontSize }));
  }

  return runs;
}

interface InlinePart {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
}

function splitInlineHtml(html: string): InlinePart[] {
  const parts: InlinePart[] = [];

  // Remove <br> tags
  let cleaned = html.replace(/<br\s*\/?>/gi, '');

  // Use a simple state machine to parse nested inline tags
  let pos = 0;
  const stack: { tag: string }[] = [];

  while (pos < cleaned.length) {
    // Check for opening tag
    const tagMatch = cleaned.substring(pos).match(/^<(strong|b|em|i|u|strike|s)([^>]*)>/i);
    if (tagMatch) {
      stack.push({ tag: tagMatch[1].toLowerCase() });
      pos += tagMatch[0].length;
      continue;
    }

    // Check for closing tag
    const closeMatch = cleaned.substring(pos).match(/^<\/(strong|b|em|i|u|strike|s)>/i);
    if (closeMatch) {
      const closingTag = closeMatch[1].toLowerCase();
      // Pop matching tag from stack
      for (let i = stack.length - 1; i >= 0; i--) {
        if (stack[i].tag === closingTag || 
            (closingTag === 'b' && stack[i].tag === 'strong') ||
            (closingTag === 'strong' && stack[i].tag === 'b') ||
            (closingTag === 'i' && stack[i].tag === 'em') ||
            (closingTag === 'em' && stack[i].tag === 'i')) {
          stack.splice(i, 1);
          break;
        }
      }
      pos += closeMatch[0].length;
      continue;
    }

    // Skip other HTML tags
    const otherTag = cleaned.substring(pos).match(/^<[^>]*>/);
    if (otherTag) {
      pos += otherTag[0].length;
      continue;
    }

    // Collect text content
    let textEnd = cleaned.indexOf('<', pos);
    if (textEnd === -1) textEnd = cleaned.length;

    const text = cleaned.substring(pos, textEnd);
    if (text) {
      const isBold = stack.some(s => s.tag === 'strong' || s.tag === 'b');
      const isItalic = stack.some(s => s.tag === 'em' || s.tag === 'i');
      const isUnderline = stack.some(s => s.tag === 'u');

      parts.push({ text, bold: isBold, italic: isItalic, underline: isUnderline });
    }

    pos = textEnd;
  }

  return parts;
}

// ─── API Endpoint: Generate PDF ─────────────────────────────────────────────
app.post('/api/generate-pdf', async (req, res) => {
  try {
    const { filledHtml } = req.body;

    if (!filledHtml) {
      return res.status(400).json({ error: 'filledHtml is required' });
    }

    console.log('[PDF] Received request, generating DOCX...');

    // 1) Parse HTML into docx paragraphs
    const docChildren = htmlToDocxParagraphs(filledHtml);

    // 2) Load letterhead background image
    let headers: any = undefined;
    const bgImagePath = path.resolve(__dirname, '..', 'public', 'a.png');
    let bgFileBytes: Uint8Array | null = null;

    try {
      if (fs.existsSync(bgImagePath)) {
        bgFileBytes = new Uint8Array(fs.readFileSync(bgImagePath));
        console.log('[PDF] Letterhead image loaded successfully');
      }
    } catch (e) {
      console.warn('[PDF] Could not load a.png:', e);
    }

    if (bgFileBytes) {
      const imageWidth = 794;
      const imageHeight = 1123;

      headers = {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new ImageRun({
                  data: bgFileBytes,
                  transformation: {
                    width: imageWidth,
                    height: imageHeight,
                  },
                  floating: {
                    horizontalPosition: {
                      relative: HorizontalPositionRelativeFrom.PAGE,
                      offset: 0,
                    },
                    verticalPosition: {
                      relative: VerticalPositionRelativeFrom.PAGE,
                      offset: 0,
                    },
                    behindDocument: true,
                    wrap: {
                      type: TextWrappingType.NONE,
                      side: TextWrappingSide.BOTH_SIDES,
                    },
                  },
                }),
              ],
            }),
          ],
        }),
      };
    }

    // 3) Create DOCX document
    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: 4500,
              bottom: 2800,
              left: 1400,
              right: 1400,
            },
          },
        },
        headers,
        children: docChildren,
      }],
    });

    const docxBuffer = await Packer.toBuffer(doc);
    console.log(`[PDF] DOCX generated: ${docxBuffer.length} bytes`);

    // 4) Convert DOCX to PDF using LibreOffice
    console.log('[PDF] Converting DOCX to PDF via LibreOffice...');
    const pdfBuffer = await libreConvertAsync(docxBuffer, '.pdf', undefined);
    console.log(`[PDF] PDF generated: ${pdfBuffer.length} bytes`);

    // 5) Send the PDF back
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="Termo_Preenchido_${Date.now()}.pdf"`);
    res.setHeader('Content-Length', pdfBuffer.length);
    res.send(pdfBuffer);

  } catch (error: any) {
    console.error('[PDF] Error:', error);
    res.status(500).json({
      error: 'Falha ao gerar PDF',
      details: error?.message || String(error),
    });
  }
});

// ─── API Endpoint: Generate DOCX (optional, for future use) ─────────────────
app.post('/api/generate-docx', async (req, res) => {
  try {
    const { filledHtml } = req.body;

    if (!filledHtml) {
      return res.status(400).json({ error: 'filledHtml is required' });
    }

    const docChildren = htmlToDocxParagraphs(filledHtml);

    let headers: any = undefined;
    const bgImagePath = path.resolve(__dirname, '..', 'public', 'a.png');
    let bgFileBytes: Uint8Array | null = null;

    try {
      if (fs.existsSync(bgImagePath)) {
        bgFileBytes = new Uint8Array(fs.readFileSync(bgImagePath));
      }
    } catch (e) {
      console.warn('[DOCX] Could not load a.png:', e);
    }

    if (bgFileBytes) {
      headers = {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new ImageRun({
                  data: bgFileBytes,
                  transformation: { width: 794, height: 1123 },
                  floating: {
                    horizontalPosition: {
                      relative: HorizontalPositionRelativeFrom.PAGE,
                      offset: 0,
                    },
                    verticalPosition: {
                      relative: VerticalPositionRelativeFrom.PAGE,
                      offset: 0,
                    },
                    behindDocument: true,
                    wrap: {
                      type: TextWrappingType.NONE,
                      side: TextWrappingSide.BOTH_SIDES,
                    },
                  },
                }),
              ],
            }),
          ],
        }),
      };
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: 4500,
              bottom: 2800,
              left: 1400,
              right: 1400,
            },
          },
        },
        headers,
        children: docChildren,
      }],
    });

    const docxBuffer = await Packer.toBuffer(doc);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="Termo_Preenchido_${Date.now()}.docx"`);
    res.setHeader('Content-Length', docxBuffer.length);
    res.send(docxBuffer);

  } catch (error: any) {
    console.error('[DOCX] Error:', error);
    res.status(500).json({
      error: 'Falha ao gerar DOCX',
      details: error?.message || String(error),
    });
  }
});

// ─── Health check ───────────────────────────────────────────────────────────
app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

app.listen(PORT, () => {
  console.log(`\n🚀 AutoTermos Backend running on http://localhost:${PORT}`);
  console.log(`   POST /api/generate-pdf   → Gera DOCX e converte para PDF`);
  console.log(`   POST /api/generate-docx  → Gera DOCX diretamente`);
  console.log(`   GET  /api/health         → Health check\n`);
});
