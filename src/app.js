import * as pdfjsLib from 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs';

const TWIPS_PER_POINT = 20;
const EMUS_PER_POINT = 12700;
const RENDER_SCALE = 2.5;
const DEFAULT_FONT_SIZE = 22;
const LINE_Y_TOLERANCE = 4;
const encoder = new TextEncoder();
let selectedFile = null;

const fileInput = document.querySelector('#fileInput');
const dropZone = document.querySelector('#dropZone');
const convertButton = document.querySelector('#convertButton');
const fileName = document.querySelector('#fileName');
const fileMeta = document.querySelector('#fileMeta');
const progress = document.querySelector('#progress');
const message = document.querySelector('#message');

function xmlEscape(value) {
  return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;');
}

function pointsToTwips(value) {
  return Math.round(value * TWIPS_PER_POINT);
}

function pointsToEmus(value) {
  return Math.round(value * EMUS_PER_POINT);
}

function cleanFileName(name) {
  return name.replace(/\.pdf$/i, '').replace(/[\\/:*?"<>|]+/g, '-');
}

function setMessage(text, state = '') {
  message.textContent = text;
  message.className = `message ${state}`.trim();
}

function canvasToBlob(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => (blob ? resolve(blob) : reject(new Error('Could not encode the PDF page image.'))), 'image/png');
  });
}

function textRunXml(text, fontSize = DEFAULT_FONT_SIZE) {
  const escaped = xmlEscape(text);
  const preserveSpace = /^\s|\s$|\s{2,}/.test(text) ? ' xml:space="preserve"' : '';
  return `<w:r><w:rPr><w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:cs="Cambria Math"/></w:rPr><w:t${preserveSpace}>${escaped}</w:t></w:r>`;
}

function paragraphXml(line, pageWidthPt) {
  const leftTwips = Math.max(0, pointsToTwips(line.x));
  const availableWidth = Math.max(1, pageWidthPt - line.x);
  const fontSize = Math.max(14, Math.min(40, Math.round(line.fontSize * 2) || DEFAULT_FONT_SIZE));
  const spacingAfter = Math.max(0, pointsToTwips(line.gapAfter || 0));

  return `<w:p><w:pPr><w:ind w:left="${leftTwips}"/><w:spacing w:before="0" w:after="${spacingAfter}"/><w:tabs><w:tab w:val="left" w:pos="${pointsToTwips(availableWidth)}"/></w:tabs></w:pPr>${textRunXml(line.text, fontSize)}</w:p>`;
}

function groupItemsIntoEditableLines(items = []) {
  const rows = [];

  for (const item of items) {
    if (!item || typeof item.str !== 'string' || !Array.isArray(item.transform) || item.transform.length < 6) continue;

    const [, , scaleX, scaleY, x, y] = item.transform;
    const text = item.str;
    if (!text.trim()) continue;

    const fontSize = Math.hypot(scaleX, scaleY) || Math.abs(scaleY) || 11;
    let row = rows.find((candidate) => Math.abs(candidate.y - y) <= LINE_Y_TOLERANCE);
    if (!row) {
      row = { y, items: [] };
      rows.push(row);
    }
    row.items.push({ x, text, fontSize });
  }

  rows.sort((a, b) => b.y - a.y);

  return rows.map((row, index) => {
    row.items.sort((a, b) => a.x - b.x);
    const text = row.items.map((item, itemIndex) => {
      if (itemIndex === 0) return item.text;
      const previous = row.items[itemIndex - 1];
      const needsSpace = !previous.text.endsWith(' ') && !item.text.startsWith(' ');
      return `${needsSpace ? ' ' : ''}${item.text}`;
    }).join('');
    const nextRow = rows[index + 1];

    return {
      text,
      x: Math.min(...row.items.map((item) => item.x)),
      y: row.y,
      fontSize: row.items.reduce((sum, item) => sum + item.fontSize, 0) / row.items.length,
      gapAfter: nextRow ? Math.max(0, row.y - nextRow.y - 10) : 0,
    };
  });
}

async function extractTextbookPages(file) {
  const bytes = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
  const pages = [];
  progress.hidden = false;

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    progress.textContent = `Rendering exact page and extracting editable text ${pageNumber} of ${pdf.numPages}`;
    const page = await pdf.getPage(pageNumber);
    const nativeViewport = page.getViewport({ scale: 1 });
    const renderViewport = page.getViewport({ scale: RENDER_SCALE });
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d', { alpha: false });

    canvas.width = Math.ceil(renderViewport.width);
    canvas.height = Math.ceil(renderViewport.height);
    await page.render({ canvasContext: context, viewport: renderViewport }).promise;

    const textContent = await page.getTextContent({ includeMarkedContent: true });
    const imageBlob = await canvasToBlob(canvas);

    pages.push({
      imageData: new Uint8Array(await imageBlob.arrayBuffer()),
      widthPt: nativeViewport.width,
      heightPt: nativeViewport.height,
      pageNumber,
      lines: groupItemsIntoEditableLines(textContent.items),
    });
  }

  return pages;
}

function visualPageXml(page, relationshipId) {
  const widthEmus = pointsToEmus(page.widthPt);
  const heightEmus = pointsToEmus(page.heightPt);
  const widthTwips = pointsToTwips(page.widthPt);
  const heightTwips = pointsToTwips(page.heightPt);
  const orientation = page.widthPt > page.heightPt ? ' w:orient="landscape"' : '';
  const pageBreak = page.pageNumber === 1 ? '' : '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

  return `${pageBreak}<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${widthEmus}" cy="${heightEmus}"/><wp:docPr id="${page.pageNumber}" name="Exact PDF page ${page.pageNumber}" descr="Rendered page preserves exact textbook math and layout."/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${page.pageNumber}" name="page-${page.pageNumber}.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmus}" cy="${heightEmus}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="${widthTwips}" w:h="${heightTwips}"${orientation}/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
}

function editableTextPageXml(page) {
  const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const heading = `<w:p><w:pPr><w:spacing w:after="160"/></w:pPr>${textRunXml(`Editable text for page ${page.pageNumber}`, 28)}</w:p>`;
  const body = page.lines.length
    ? page.lines.map((line) => paragraphXml(line, page.widthPt)).join('')
    : '<w:p><w:r><w:t>No selectable text was found on this page. The exact visual page is preserved above.</w:t></w:r></w:p>';

  return `${pageBreak}${heading}${body}`;
}

function createTextbookDocxZip(pages) {
  const files = {};
  const imageRelationships = pages.map((page) => `<Relationship Id="rIdPage${page.pageNumber}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/page-${page.pageNumber}.png"/>`).join('');
  const visualPages = pages.map((page) => visualPageXml(page, `rIdPage${page.pageNumber}`)).join('');
  const editablePages = pages.map(editableTextPageXml).join('');

  files['[Content_Types].xml'] = encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  files['_rels/.rels'] = encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  files['word/_rels/document.xml.rels'] = encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${imageRelationships}</Relationships>`);
  files['word/styles.xml'] = encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:cs="Cambria Math"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style></w:styles>');
  files['word/document.xml'] = encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${visualPages}${editablePages}</w:body></w:document>`);

  for (const page of pages) {
    files[`word/media/page-${page.pageNumber}.png`] = page.imageData;
  }

  return window.fflate.zipSync(files, { level: 6 });
}

function downloadDocx(data, name) {
  const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  anchor.click();
  URL.revokeObjectURL(url);
}

function selectFile(file) {
  if (!file || (!file.name.toLowerCase().endsWith('.pdf') && file.type !== 'application/pdf')) {
    setMessage('Please choose a PDF file.', 'error');
    return;
  }
  selectedFile = file;
  fileName.textContent = file.name;
  fileMeta.textContent = `${(file.size / 1024 / 1024).toFixed(2)} MB selected`;
  convertButton.disabled = false;
  progress.hidden = true;
  setMessage('Ready. The DOCX will keep the textbook pages visually exact and add editable text pages after them.');
}

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (event) => event.preventDefault());
dropZone.addEventListener('drop', (event) => {
  event.preventDefault();
  selectFile(event.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => selectFile(fileInput.files[0]));

convertButton.addEventListener('click', async () => {
  if (!selectedFile) return;
  convertButton.disabled = true;
  setMessage('Rendering exact textbook pages and extracting editable text...');
  try {
    const pages = await extractTextbookPages(selectedFile);
    progress.textContent = 'Building the textbook .docx file...';
    const docx = createTextbookDocxZip(pages);
    const outputName = `${cleanFileName(selectedFile.name)}-textbook-editable.docx`;
    downloadDocx(docx, outputName);
    setMessage(`Created ${outputName}. Exact pages preserve math/layout; editable text pages are included after the visual pages.`, 'done');
  } catch (error) {
    setMessage(error instanceof Error ? error.message : 'Conversion failed. Please try another PDF.', 'error');
  } finally {
    convertButton.disabled = false;
  }
});
