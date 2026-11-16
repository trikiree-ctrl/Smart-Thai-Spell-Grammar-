import mammoth from 'mammoth';
import { Document, Paragraph, TextRun, HeadingLevel, Packer } from 'docx';
import JSZip from 'jszip';
import * as pdfjsLib from 'pdfjs-dist';

// Configure pdfjs-dist worker for browser environments
if (typeof window !== 'undefined' && pdfjsLib.GlobalWorkerOptions) {
  try {
    pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
      'pdfjs-dist/build/pdf.worker.min.mjs',
      import.meta.url
    ).href;
  } catch {
    // Fallback if environment doesn't support relative worker URL
  }
}

export interface ExtractedDocument {
  text: string;
  fileType: string;
  fileName: string;
  originalBuffer?: ArrayBuffer;
  pageCount?: number;
  isWarning?: boolean;
}

/**
 * Extract plain text from supported document files (.docx, .pdf, .txt, .md)
 * Defensively handles binary files and scanned PDFs without producing alien/garbled characters
 */
export async function extractTextFromFile(file: File): Promise<ExtractedDocument> {
  const fileName = file.name;
  const extension = fileName.substring(fileName.lastIndexOf('.')).toLowerCase();

  // 1. Modern Word Documents (.docx)
  if (extension === '.docx') {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer });
      const text = (result.value || '').trim();
      if (!text) {
        return {
          text: '⚠️ ไม่พบข้อความในไฟล์ .docx นี้ หรือเอกสารไม่มีเนื้อหาที่เป็นข้อความ',
          fileType: 'docx',
          fileName,
          isWarning: true
        };
      }
      return {
        text,
        fileType: 'docx',
        fileName,
        originalBuffer: arrayBuffer,
        isWarning: false
      };
    } catch (docxErr) {
      console.warn('Failed to parse docx:', docxErr);
      return {
        text: '⚠️ ไม่สามารถอ่านไฟล์ .docx นี้ได้ กรุณาตรวจสอบว่าไฟล์ไม่เสียหาย หรือบันทึกใหม่ใน Microsoft Word',
        fileType: 'docx',
        fileName,
        isWarning: true
      };
    }
  }

  // 2. Legacy Word Documents (.doc - Word 97-2003 binary format)
  if (extension === '.doc') {
    return {
      text: '⚠️ ไฟล์นี้เป็นรูปแบบ Microsoft Word รุ่นเก่า (.doc ซึ่งเป็นไฟล์ไบนารีเข้ารหัสเดิม)\n\nกรุณาเปิดไฟล์ในโปรแกรม Microsoft Word แล้วเลือก "บันทึกเป็น (Save As)" เป็นรูปแบบ .docx หรือคัดลอกข้อความจากเอกสารมาวางในช่องนี้โดยตรงครับ',
      fileType: 'doc',
      fileName,
      isWarning: true
    };
  }

  // 3. PDF Documents (.pdf)
  if (extension === '.pdf') {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const { text, pageCount } = await extractTextFromPdfWithPdfJs(arrayBuffer);
      if (!text || text.trim().length === 0) {
        return {
          text: '⚠️ ไม่พบข้อความที่สามารถคัดลอกได้ในไฟล์ PDF นี้\n\nสาเหตุ: เอกสารนี้อาจเป็น "ภาพถ่ายหรือภาพสแกน (Scanned Image PDF)" ซึ่งไม่มี Text Layer ข้อความภายในไฟล์\n\nวิธีแก้ไข: กรุณาคัดลอกข้อความจากเอกสารมาวางในช่องนี้โดยตรง หรือใช้ไฟล์ต้นฉบับ .docx / .txt ครับ',
          fileType: 'pdf',
          fileName,
          isWarning: true
        };
      }
      return {
        text,
        fileType: 'pdf',
        fileName,
        originalBuffer: arrayBuffer,
        pageCount,
        isWarning: false
      };
    } catch (pdfErr) {
      console.warn('PDF.js parsing error:', pdfErr);
      return {
        text: '⚠️ ไม่สามารถดึงข้อความจากไฟล์ PDF นี้ได้ เอกสารอาจมีการเข้ารหัสป้องกันหรือเป็นภาพสแกน กรุณาคัดลอกข้อความมาวางโดยตรงครับ',
        fileType: 'pdf',
        fileName,
        isWarning: true
      };
    }
  }

  // 4. Plain Text (.txt, .md, .csv)
  try {
    const text = await file.text();
    // Safety check: verify the text doesn't contain binary null bytes or compressed data
    if (isBinaryGarbage(text)) {
      return {
        text: '⚠️ รูปแบบไฟล์ไม่รองรับ (ตรวจพบข้อมูลไบนารี) กรุณาใช้ไฟล์เอกสาร .docx, .pdf หรือ .txt ที่ถูกต้อง',
        fileType: extension.replace('.', '') || 'unknown',
        fileName,
        isWarning: true
      };
    }
    return {
      text,
      fileType: extension.replace('.', '') || 'txt',
      fileName,
      isWarning: false
    };
  } catch (err) {
    return {
      text: '⚠️ ไม่สามารถเปิดไฟล์นี้ได้ กรุณาลองใหม่อีกครั้งหรือคัดลอกข้อความมาวาง',
      fileType: extension.replace('.', '') || 'unknown',
      fileName,
      isWarning: true
    };
  }
}

/**
 * Check if extracted text is actually binary garbage (e.g. raw binary file opened as text)
 */
function isBinaryGarbage(text: string): boolean {
  if (!text || text.length === 0) return false;
  // Check for null characters or high concentration of unprintable control characters
  const sample = text.slice(0, 1000);
  let unprintableCount = 0;
  for (let i = 0; i < sample.length; i++) {
    const code = sample.charCodeAt(i);
    // null byte or control chars (except \r, \n, \t)
    if (code === 0 || (code < 32 && code !== 9 && code !== 10 && code !== 13)) {
      unprintableCount++;
    }
  }
  return unprintableCount > 10;
}

/**
 * Robust PDF text extractor using pdfjs-dist
 * Decodes all font CMaps, ToUnicode maps, and FlateDecode streams accurately
 * Embeds page markers so page locations are tracked precisely
 */
async function extractTextFromPdfWithPdfJs(buffer: ArrayBuffer): Promise<{ text: string; pageCount: number }> {
  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    verbosity: 0,
  });

  const pdf = await loadingTask.promise;
  const pageTexts: string[] = [];

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();
    const strings = textContent.items
      .map((item: any) => item.str || '')
      .filter((s: string) => s.trim().length > 0);

    if (strings.length > 0) {
      const pageBody = strings.join(' ');
      pageTexts.push(`--- หน้า ${pageNum} ---\n${pageBody}`);
    }
  }

  return {
    text: pageTexts.join('\n\n\f'),
    pageCount: pdf.numPages
  };
}

/**
 * Patch an original Word (.docx) document buffer by replacing confirmed typos
 * inside word/document.xml, word/header*.xml, word/footer*.xml
 * This PRESERVES 100% of original document formatting, tables, images, margins, fonts, headers and footers!
 */
export async function patchOriginalDocx(
  originalBuffer: ArrayBuffer,
  corrections: Array<{ original: string; replacement: string }>
): Promise<Blob> {
  const zip = await JSZip.loadAsync(originalBuffer);

  const replaceInXml = (xml: string): string => {
    let updated = xml;
    for (const corr of corrections) {
      if (!corr.original || !corr.replacement || corr.original === corr.replacement) continue;
      updated = replaceOpenXmlTextNode(updated, corr.original, corr.replacement);
    }
    return updated;
  };

  // 1. Process word/document.xml
  const docXmlFile = zip.file('word/document.xml');
  if (docXmlFile) {
    const xml = await docXmlFile.async('text');
    const modified = replaceInXml(xml);
    zip.file('word/document.xml', modified);
  }

  // 2. Process headers and footers
  const allFiles = Object.keys(zip.files);
  for (const path of allFiles) {
    if (path.startsWith('word/header') || path.startsWith('word/footer')) {
      const hfFile = zip.file(path);
      if (hfFile) {
        const xml = await hfFile.async('text');
        const modified = replaceInXml(xml);
        zip.file(path, modified);
      }
    }
  }

  // Generate output docx blob preserving all original files
  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });
}

/**
 * Safely replace target word inside OpenXML text elements (<w:t>)
 */
export function replaceOpenXmlTextNode(xml: string, target: string, replacement: string): string {
  if (!xml || !target) return xml;

  const escapeXml = (str: string) =>
    str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');

  const escapedTarget = escapeXml(target);
  const escapedReplacement = escapeXml(replacement);

  let result = xml.replace(/(<w:t(?:\s+[^>]*)?>)([\s\S]*?)(<\/w:t>)/g, (match, openTag, textContent, closeTag) => {
    if (textContent.includes(escapedTarget)) {
      return `${openTag}${textContent.split(escapedTarget).join(escapedReplacement)}${closeTag}`;
    }
    if (textContent.includes(target)) {
      return `${openTag}${textContent.split(target).join(escapedReplacement)}${closeTag}`;
    }
    return match;
  });

  return result;
}

/**
 * Generate and download a formatted Microsoft Word (.docx) document
 */
export async function exportToDocx(
  text: string,
  fileName: string = 'เอกสารที่ผ่านการตรวจคำผิด.docx',
  meta?: { docTitle?: string; date?: string; errorCount?: number }
) {
  const paragraphs = text.split('\n').map(line => {
    return new Paragraph({
      spacing: { line: 360, after: 120 },
      children: [
        new TextRun({
          text: line,
          font: 'Sarabun',
          size: 32, // 16pt in half-points
        }),
      ],
    });
  });

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            heading: HeadingLevel.TITLE,
            spacing: { after: 240 },
            children: [
              new TextRun({
                text: meta?.docTitle || 'เอกสารที่ผ่านการตรวจทานและแก้ไขคำผิด',
                font: 'Sarabun',
                bold: true,
                size: 40,
              }),
            ],
          }),
          new Paragraph({
            spacing: { after: 240 },
            children: [
              new TextRun({
                text: `ตรวจทานโดย: ThaiDoc Inspector AI | วันที่: ${meta?.date || new Date().toLocaleDateString('th-TH')} | แก้ไขคำผิดแล้ว: ${meta?.errorCount || 0} จุด`,
                font: 'Sarabun',
                italics: true,
                size: 24,
                color: '666666',
              }),
            ],
          }),
          ...paragraphs,
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, fileName.endsWith('.docx') ? fileName : `${fileName}.docx`);
}

/**
 * Export plain text file
 */
export function exportToTxt(text: string, fileName: string = 'เอกสารแก้ไข.txt') {
  const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
  downloadBlob(blob, fileName.endsWith('.txt') ? fileName : `${fileName}.txt`);
}

/**
 * Trigger file download in browser
 */
export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
