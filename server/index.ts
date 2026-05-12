import express from 'express';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import helmet from 'helmet';
import dotenv from 'dotenv';
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
import { GoogleGenAI } from "@google/genai";


// Load environment variables
dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3001;

// --- Middleware ---
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'", "'unsafe-inline'"],
      styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com"],
      fontSrc: ["'self'", "https://fonts.gstatic.com"],
      imgSrc: ["'self'", "data:", "blob:"],
      connectSrc: ["'self'"],
    },
  },
}));

app.use(cors({
  origin: ['http://localhost:3000', 'http://127.0.0.1:3000'],
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type'],
}));
app.use(express.json({ limit: '50mb' }));


// --- Helper: Parse HTML into docx Paragraphs ---
function htmlToDocxParagraphs(html: string): Paragraph[] {
  const paragraphs: Paragraph[] = [];
  const blockRegex = /<(h[1-6]|p|li)([^>]*)>([\s\S]*?)<\/\1>/gi;
  let match;

  while ((match = blockRegex.exec(html)) !== null) {
    const tagName = match[1].toLowerCase();
    const attributes = match[2] || '';
    const innerHtml = match[3];

    const isCenter = attributes.includes('ql-align-center') || attributes.includes('text-align: center');
    const isRight = attributes.includes('ql-align-right') || attributes.includes('text-align: right');
    const isJustify = attributes.includes('ql-align-justify') || attributes.includes('text-align: justify');

    const alignment = isCenter ? AlignmentType.CENTER :
                      isRight ? AlignmentType.RIGHT :
                      isJustify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT;

    const children = parseInlineContent(innerHtml, tagName);

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

  if (paragraphs.length === 0) {
    const plainText = html.replace(/<[^>]*>?/gm, '').replace(/&nbsp;/gi, ' ');
    paragraphs.push(new Paragraph({
      children: [new TextRun({ text: plainText, size: 22, font: 'Cormorant Garamond' })],
    }));
  }

  return paragraphs;
}

function parseInlineContent(html: string, parentTag: string): TextRun[] {
  const runs: TextRun[] = [];
  const isHeading = parentTag.startsWith('h');
  const headingSizes: Record<string, number> = {
    'h1': 32, 'h2': 28, 'h3': 24, 'h4': 22, 'h5': 20, 'h6': 18
  };
  const baseFontSize = isHeading ? (headingSizes[parentTag] || 22) : 22;

  const parts = splitInlineHtml(html);

  for (const part of parts) {
    const text = part.text
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"');
      
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
    runs.push(new TextRun({ text: '', size: baseFontSize, font: 'Cormorant Garamond' }));
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
  let cleaned = html.replace(/<br\s*\/?>/gi, '');
  let pos = 0;
  const stack: { tag: string }[] = [];

  while (pos < cleaned.length) {
    const tagMatch = cleaned.substring(pos).match(/^<(strong|b|em|i|u|strike|s)([^>]*)>/i);
    if (tagMatch) {
      stack.push({ tag: tagMatch[1].toLowerCase() });
      pos += tagMatch[0].length;
      continue;
    }

    const closeMatch = cleaned.substring(pos).match(/^<\/(strong|b|em|i|u|strike|s)>/i);
    if (closeMatch) {
      const closingTag = closeMatch[1].toLowerCase();
      for (let i = stack.length - 1; i >= 0; i--) {
        if (stack[i].tag === closingTag || 
            (closingTag === 'b' && stack[i].tag === 'strong') ||
            (closingTag === 'strong' && stack[i].tag === 'b')) {
          stack.splice(i, 1);
          break;
        }
      }
      pos += closeMatch[0].length;
      continue;
    }

    const otherTag = cleaned.substring(pos).match(/^<[^>]*>/);
    if (otherTag) {
      pos += otherTag[0].length;
      continue;
    }

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

// --- API Endpoints ---
app.post('/api/generate-docx', async (req, res) => {
  try {
    const { filledHtml } = req.body;
    if (!filledHtml) return res.status(400).json({ error: 'filledHtml is required' });

    const docChildren = htmlToDocxParagraphs(filledHtml);
    let headers: any = undefined;
    const bgImagePath = path.resolve(__dirname, '..', 'public', 'a.png');
    let bgFileBytes: Uint8Array | null = null;

    if (fs.existsSync(bgImagePath)) {
      bgFileBytes = new Uint8Array(fs.readFileSync(bgImagePath));
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
                    horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 0 },
                    verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 0 },
                    behindDocument: true,
                    wrap: { type: TextWrappingType.NONE, side: TextWrappingSide.BOTH_SIDES },
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
          page: { margin: { top: 4500, bottom: 2800, left: 1400, right: 1400 } },
        },
        headers,
        children: docChildren,
      }],
    });

    const docxBuffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="Termo_${Date.now()}.docx"`);
    res.send(docxBuffer);
  } catch (error: any) {
    console.error('[DOCX] Error:', error);
    res.status(500).json({ error: 'Falha ao gerar DOCX', details: error?.message });
  }
});

app.post('/api/analyze-document', async (req, res) => {
  try {
    const { fileBase64, mimeType, originalText } = req.body;
    const apiKey = process.env.GEMINI_API_KEY;

    if (!apiKey) {
      return res.status(500).json({ error: 'GEMINI_API_KEY não configurada no servidor.' });
    }

    const ai = new GoogleGenAI(apiKey);
    const prompt = `Analise este documento e retorne o texto integral dele, mas substituindo campos que claramente são variáveis (como nomes, CPFs, datas, endereços, nomes de empresas) por marcadores entre colchetes, como [NOME], [CPF], [DATA], [ENDEREÇO], [NOME DA EMPRESA]. 
    
    Se o documento já contiver marcadores entre colchetes, mantenha-os e apenas identifique se faltou algum campo importante para ser transformado em variável.
    
    Retorne APENAS o texto do modelo resultante. Nada de conversas, apenas o conteúdo do documento com os colchetes.`;

    const model = ai.getGenerativeModel({ model: "gemini-1.5-flash" });
    
    const parts: any[] = [{ text: prompt }];
    if (mimeType === 'application/pdf' && fileBase64) {
      parts.push({ inlineData: { data: fileBase64, mimeType } });
    } else if (originalText) {
      parts.push({ text: `Conteúdo extraído do arquivo: \n\n${originalText}` });
    }

    const result = await model.generateContent({
      contents: [{ role: "user", parts }]
    });

    res.json({ text: result.response.text() });
  } catch (error: any) {
    console.error('[AI] Error:', error);
    res.status(500).json({ error: 'Falha na análise por IA', details: 'Ocorreu um erro ao processar o documento.' });
  }
});


app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok', version: '1.0.0', timestamp: new Date().toISOString() });
});

app.listen(PORT, () => {
  console.log(`\n🚀 AutoTermos Backend v1 running on http://localhost:${PORT}\n`);
});
