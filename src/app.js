import * as pdfjsLib from 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs';

const TWIPS_PER_POINT = 20;
const EMUS_PER_POINT = 12700;
const RENDER_SCALE = 2.5;
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

async function renderPdf(file) {
  const bytes = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
  const pages = [];
  progress.hidden = false;

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    progress.textContent = `Rendering page ${pageNumber} of ${pdf.numPages}`;
    const page = await pdf.getPage(pageNumber);
    const nativeViewport = page.getViewport({ scale: 1 });
    const renderViewport = page.getViewport({ scale: RENDER_SCALE });
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d', { alpha: false });
    canvas.width = Math.ceil(renderViewport.width);
    canvas.height = Math.ceil(renderViewport.height);
    await page.render({ canvasContext: context, viewport: renderViewport }).promise;
    const blob = await canvasToBlob(canvas);
    pages.push({
      data: new Uint8Array(await blob.arrayBuffer()),
      widthPt: nativeViewport.width,
      heightPt: nativeViewport.height,
      pageNumber,
    });
  }

  return pages;
}

function sectionXml(page, relationshipId) {
  const widthTwips = pointsToTwips(page.widthPt);
  const heightTwips = pointsToTwips(page.heightPt);
  const widthEmus = pointsToEmus(page.widthPt);
  const heightEmus = pointsToEmus(page.heightPt);
  const orientation = page.widthPt > page.heightPt ? ' w:orient="landscape"' : '';
  const breakXml = page.pageNumber === 1 ? '' : '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

  return `${breakXml}<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="0"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${widthEmus}" cy="${heightEmus}"/><wp:docPr id="${page.pageNumber}" name="PDF page ${page.pageNumber}" descr="Full-page rendering used to preserve exact textbook PDF formatting."/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${page.pageNumber}" name="page-${page.pageNumber}.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmus}" cy="${heightEmus}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="${widthTwips}" w:h="${heightTwips}"${orientation}/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
}

function createDocxZip(pages) {
  const files = {};
  const rels = pages.map((page) => `<Relationship Id="rId${page.pageNumber}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/page-${page.pageNumber}.png"/>`).join('');
  const body = pages.map((page) => sectionXml(page, `rId${page.pageNumber}`)).join('');

  files['[Content_Types].xml'] = encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  files['_rels/.rels'] = encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  files['word/_rels/document.xml.rels'] = encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);
  files['word/document.xml'] = encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${body}</w:body></w:document>`);

  for (const page of pages) {
    files[`word/media/page-${page.pageNumber}.png`] = page.data;
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
  setMessage('Ready to convert with exact visual formatting.');
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
  setMessage('Rendering each PDF page at high resolution...');
  try {
    const pages = await renderPdf(selectedFile);
    progress.textContent = 'Building the .docx file...';
    const docx = createDocxZip(pages);
    const outputName = `${cleanFileName(selectedFile.name)}-exact-layout.docx`;
    downloadDocx(docx, outputName);
    setMessage(`Created ${xmlEscape(outputName)}. The Word document preserves the PDF exactly as page images, including equations and textbook formatting.`, 'done');
  } catch (error) {
    setMessage(error instanceof Error ? error.message : 'Conversion failed. Please try another PDF.', 'error');
  } finally {
    convertButton.disabled = false;
  }
});
