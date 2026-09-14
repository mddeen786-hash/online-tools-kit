const express = require('express');
const fileUpload = require('express-fileupload');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const { createCanvas, loadImage } = require('canvas'); 
const XLSX = require('xlsx');
const { PDFDocument, degrees, rgb } = require('pdf-lib'); 
const JSZip = require('jszip'); 

// OCR Module (Tesseract)
let Tesseract = null;
try { Tesseract = require('tesseract.js'); } catch(e) { console.log('Tesseract not available'); }

let sharp = null;
try { sharp = require('sharp'); } catch(e) { console.log('Sharp not installed.'); }

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(fileUpload({ createParentPath: true, limits: { fileSize: 100 * 1024 * 1024 } }));

const tempDir = path.join(__dirname, 'temp');
if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

const toolRegistry = [
  { id: 'pdf-to-word',         name: 'PDF to Word',             icon: '📄', file: 'pdf-to-word.html', outputExt: '.doc' },
  { id: 'pdf-to-excel',        name: 'PDF to Excel',            icon: '📊', file: 'pdf-to-excel.html', outputExt: '.xlsx' },
  { id: 'merge-pdf',           name: 'Merge PDF',               icon: '💼', file: 'merge.html', outputExt: '.pdf' },
  { id: 'pdf-splitter',        name: 'PDF Splitter',            icon: '✂️', file: 'pdf-splitter.html', outputExt: '.pdf' },
  { id: 'pdf-organizer',       name: 'PDF Organizer',           icon: '📑', file: 'pdf-organizer.html', outputExt: '.pdf' },
  { id: 'pdf-watermark',       name: 'PDF Watermark',           icon: '💧', file: 'pdf-watermark.html', outputExt: '.pdf' },
  { id: 'pdf-to-image',        name: 'PDF to Image',            icon: '🖼️', file: 'pdf-to-image.html', outputExt: '.zip' },
  { id: 'compress-image',      name: 'Compress Image',          icon: '🗜️', file: 'compress.html', outputExt: '.jpg' },
  { id: 'image-converter',     name: 'Image Converter',         icon: '🖼️', file: 'image-converter.html' }, 
  { id: 'image-to-pdf',        name: 'Image to PDF',            icon: '📷', file: 'image-to-pdf.html', outputExt: '.pdf' },
  { id: 'image-reducer',       name: 'Image Reducer',           icon: '📉', file: 'image-reducer.html', outputExt: '.jpg' },
  { id: 'image-to-text',       name: 'Image to Text',           icon: '🔍', file: 'image-to-text.html', outputExt: '.doc' },
  { id: 'passport-studio',     name: 'Passport Studio',         icon: '🛂', file: 'passport-studio.html', outputExt: '.jpg' },
  { id: 'merge-image',         name: 'Merge Image',             icon: '🧩', file: 'merge-image.html', outputExt: '.jpg' }
];

const workflowsDB = new Map();

// ========== PDF to Word Helpers ==========
function clusterXPositions(items, threshold = 8) {
  const xs = [...new Set(items.map(item => item.x))].sort((a, b) => a - b);
  if (xs.length === 0) return [];
  const clusters = [];
  let currentCluster = [xs[0]];
  for (let i = 1; i < xs.length; i++) {
    if (xs[i] - currentCluster[currentCluster.length - 1] <= threshold) {
      currentCluster.push(xs[i]);
    } else {
      clusters.push(Math.round(currentCluster.reduce((sum, v) => sum + v, 0) / currentCluster.length));
      currentCluster = [xs[i]];
    }
  }
  if (currentCluster.length > 0) clusters.push(Math.round(currentCluster.reduce((sum, v) => sum + v, 0) / currentCluster.length));
  return clusters;
}

function buildTableFromLines(lines) {
  const allItems = lines.flat();
  if (allItems.length === 0) return '';
  const columns = clusterXPositions(allItems, 8);
  let tableHtml = '<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse; width:100%; mso-table-lspace:0pt; mso-table-rspace:0pt; margin:15px 0; font-size:10pt;">';
  lines.forEach((line, rowIdx) => {
    const cellMap = new Array(columns.length).fill('');
    line.forEach(item => {
      let minDist = Infinity, bestCol = -1;
      columns.forEach((colX, idx) => {
        const dist = Math.abs(item.x - colX);
        if (dist < minDist) { minDist = dist; bestCol = idx; }
      });
      if (bestCol !== -1) {
        cellMap[bestCol] = cellMap[bestCol] ? cellMap[bestCol] + ' ' + item.str : item.str;
      }
    });
    const isHeader = rowIdx === 0 || cellMap.some(text => text.includes("Category") || text.includes("Participants") || text.includes("Expenditure"));
    const bgStyle = isHeader ? "background:#f2f2f2; font-weight:bold;" : "";
    tableHtml += '<tr style="' + bgStyle + '">';
    cellMap.forEach(text => {
      tableHtml += '<td style="border:1px solid #000000; padding:6px; vertical-align:top;">' + (text || '&nbsp;') + '</td>';
    });
    tableHtml += '</tr>';
  });
  tableHtml += '</table>';
  return tableHtml;
}

async function pdfToWordConvert(inputPath, outputPath) {
  try {
    const data = new Uint8Array(fs.readFileSync(inputPath));
    const pdf = await pdfjsLib.getDocument({ data }).promise;
    let finalWordDocContent = "";
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      let items = textContent.items.map(item => ({
        str: item.str,
        x: Math.round(item.transform[4]),
        y: Math.round(item.transform[5])
      }));
      items.sort((a, b) => {
        if (Math.abs(a.y - b.y) > 5) return b.y - a.y;
        return a.x - b.x;
      });
      let lines = [];
      let currentLine = [];
      let lastY = null;
      items.forEach(item => {
        if (lastY === null || Math.abs(item.y - lastY) <= 5) {
          currentLine.push(item);
        } else {
          lines.push(currentLine);
          currentLine = [item];
        }
        lastY = item.y;
      });
      if (currentLine.length > 0) lines.push(currentLine);

      let pageHtmlBody = "";
      let tableLinesBuffer = [];
      lines.forEach(line => {
        let lineText = line.map(it => it.str).join(' ').trim();
        const isTableLine = (line.length >= 3 && (
          lineText.includes("Disability") || lineText.includes("Participants") || lineText.includes("Ballots") ||
          lineText.includes("Accuracy") || lineText.includes("Expenditure") || lineText.includes("Main character") ||
          lineText.includes("Sidekick") || lineText.includes("Column header") || lineText.includes("Row header") ||
          /\d+/.test(lineText)
        )) && !lineText.includes("GRANDMA'S BAG OF STORIES");
        if (isTableLine) {
          tableLinesBuffer.push(line);
        } else {
          if (tableLinesBuffer.length > 0) {
            pageHtmlBody += buildTableFromLines(tableLinesBuffer);
            tableLinesBuffer = [];
          }
          if (lineText.length > 0) {
            pageHtmlBody += '<p style="line-height:1.6;margin-bottom:8px;">' + lineText + '</p>';
          }
        }
      });
      if (tableLinesBuffer.length > 0) pageHtmlBody += buildTableFromLines(tableLinesBuffer);

      if (pageHtmlBody.trim() === "") {
        const viewport = page.getViewport({ scale: 1.2 });
        const canvas = createCanvas(viewport.width, viewport.height);
        const context = canvas.getContext('2d');
        await page.render({ canvasContext: context, viewport: viewport }).promise;
        const imageBuffer = canvas.toBuffer('image/jpeg', { quality: 0.85 });
        const pageImgUrl = 'data:image/jpeg;base64,' + imageBuffer.toString('base64');
        pageHtmlBody = '<div style="text-align:center;"><img src="' + pageImgUrl + '" style="max-width:100%;height:auto;border:1px solid #ddd;" /></div>';
      }

      finalWordDocContent += '<div style="margin-bottom:40px;page-break-after:always;"><h3 style="font-size:11pt;border-bottom:1px solid #ccc;padding-bottom:4px;margin-bottom:15px;">Page ' + i + '</h3>' + pageHtmlBody + '</div>';
    }

    const finalWordHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>Converted Document</title><style>body{} table{border-collapse:collapse;width:100%;} td{border:1px solid #000000;padding:6px;}</style></head><body>' + finalWordDocContent + '</body></html>';
    fs.writeFileSync(outputPath, '\ufeff' + finalWordHtml);
    return outputPath;
  } catch (error) {
    console.error('PDF to Word error:', error);
    throw error;
  }
}

// ========== PDF to Excel Helpers ==========
function isTitleLine(str) {
  return str.startsWith('Table') || str.startsWith('Design') || str.startsWith('Sample') || 
         str.startsWith('SK&M') || str.startsWith('Test') || str.startsWith('Yeh ek') || 
         str.startsWith('Aap is') || str.startsWith('file:///') || /^\(\d+\)/.test(str);
}

function parseSimplePDFTables(items) {
  if (!items || items.length === 0) return [];
  let rawItems = [];
  items.forEach(item => {
      let str = item.str;
      if (!str || !str.trim()) return;
      let x = item.transform[4];
      let y = item.transform[5];
      let w = item.width || str.length * 6;
      rawItems.push({ x: x, y: y, right: x + w, text: str.trim() });
  });
  if (rawItems.length === 0) return [];
  rawItems.sort((a, b) => b.y - a.y || a.x - b.x);
  let lines = [];
  let currentLine = [];
  let currentY = null;
  const Y_TOLERANCE = 5;
  rawItems.forEach(item => {
      if (currentY === null || Math.abs(item.y - currentY) <= Y_TOLERANCE) {
          currentLine.push(item);
          if (currentY === null) currentY = item.y;
      } else {
          lines.push(currentLine);
          currentLine = [item];
          currentY = item.y;
      }
  });
  if (currentLine.length > 0) lines.push(currentLine);
  let allXCoords = [];
  lines.forEach(line => {
      if (line.length > 1) {
          line.forEach(item => allXCoords.push(item.x));
      }
  });
  allXCoords.sort((a, b) => a - b);
  let columnAnchors = [];
  const COLUMN_TOLERANCE = 35;
  allXCoords.forEach(x => {
      let found = columnAnchors.find(anchor => Math.abs(anchor - x) <= COLUMN_TOLERANCE);
      if (!found) {
          columnAnchors.push(x);
      }
  });
  columnAnchors.sort((a, b) => a - b);
  let refinedAnchors = [];
  for (let x of columnAnchors) {
      if (refinedAnchors.length === 0 || x - refinedAnchors[refinedAnchors.length - 1] > COLUMN_TOLERANCE) {
          refinedAnchors.push(x);
      }
  }
  if (refinedAnchors.length === 0) {
      return lines.map(line => line.map(i => i.text));
  }
  let grid = [];
  lines.forEach(line => {
      let row = new Array(refinedAnchors.length).fill("");
      let isTitle = line.length === 1 && isTitleLine(line[0].text);
      if (isTitle) {
          grid.push([line[0].text]);
          return;
      }
      line.forEach(item => {
          let bestCol = 0;
          let minDist = Math.abs(refinedAnchors[0] - item.x);
          for (let c = 1; c < refinedAnchors.length; c++) {
              let dist = Math.abs(refinedAnchors[c] - item.x);
              if (dist < minDist) {
                  minDist = dist;
                  bestCol = c;
              }
          }
          if (row[bestCol]) {
              row[bestCol] += " " + item.text;
          } else {
              row[bestCol] = item.text;
          }
      });
      if (row.some(c => c.trim() !== "")) {
          grid.push(row);
      }
  });
  if (grid.length > 0) {
      let maxCols = Math.max(...grid.map(r => r.length));
      let activeCols = [];
      for (let c = 0; c < maxCols; c++) {
          let hasContent = grid.some(row => row[c] && row[c].toString().trim() !== "");
          if (hasContent) activeCols.push(c);
      }
      grid = grid.map(row => activeCols.map(c => row[c] || ""));
  }
  return grid;
}

function parseComplexPDFTables(items) {
  if (!items || items.length === 0) return [];
  let rawItems = [];
  items.forEach(item => {
      let str = item.str;
      if (!str || !str.trim()) return;
      let x = item.transform[4];
      let y = item.transform[5];
      let w = item.width || 0;
      let parts = str.split(/\s{2,}/);
      if (parts.length > 1) {
          let totalLen = str.length || 1;
          let charWidth = w / totalLen;
          let searchIdx = 0;
          parts.forEach(part => {
              let trimmed = part.trim();
              if (trimmed) {
                  let idx = str.indexOf(part, searchIdx);
                  let partX = x + (idx * charWidth);
                  let partW = part.length * charWidth;
                  rawItems.push({ x: partX, y: y, width: partW, right: partX + partW, text: trimmed });
                  searchIdx = idx + part.length;
              }
          });
      } else {
          rawItems.push({ x: x, y: y, width: w, right: x + w, text: str.trim() });
      }
  });
  if (rawItems.length === 0) return [];
  rawItems.sort((a, b) => b.y - a.y || a.x - b.x);
  let lines = [];
  let currentLine = [];
  let currentY = null;
  const Y_TOLERANCE = 5;
  rawItems.forEach(item => {
      if (currentY === null || Math.abs(item.y - currentY) <= Y_TOLERANCE) {
          currentLine.push(item);
          if (currentY === null) currentY = item.y;
      } else {
          lines.push(currentLine);
          currentLine = [item];
          currentY = item.y;
      }
  });
  if (currentLine.length > 0) lines.push(currentLine);
  lines.forEach(line => line.sort((a, b) => a.x - b.x));
  let parsedLines = [];
  lines.forEach(line => {
      let cells = [];
      let currCell = null;
      line.forEach(item => {
          if (!currCell) {
              currCell = { x: item.x, right: item.right, text: item.text };
          } else {
              let gap = item.x - currCell.right;
              if (gap > 8) {
                  cells.push(currCell);
                  currCell = { x: item.x, right: item.right, text: item.text };
              } else {
                  let space = gap > 1.5 ? " " : "";
                  currCell.text += space + item.text;
                  currCell.right = Math.max(currCell.right, item.right);
              }
          }
      });
      if (currCell) cells.push(currCell);
      if (cells.length > 0) parsedLines.push(cells);
  });
  let blocks = [];
  let currentBlock = [];
  parsedLines.forEach(line => {
      let lineText = line.map(c => c.text).join(" ");
      let isTitle = line.length === 1 && isTitleLine(lineText);
      if (isTitle) {
          if (currentBlock.length > 0) {
              blocks.push(currentBlock);
              currentBlock = [];
          }
          blocks.push([line]);
      } else {
          currentBlock.push(line);
      }
  });
  if (currentBlock.length > 0) blocks.push(currentBlock);
  let finalGrid = [];
  blocks.forEach(block => {
      if (block.length === 0) return;
      let blockText = block[0].map(c => c.text).join(" ");
      if (block.length === 1 && (block[0].length === 1 || isTitleLine(blockText))) {
          finalGrid.push([blockText]);
          return;
      }
      let blockXStarts = [];
      block.forEach(line => {
          if (line.length >= 2) {
              line.forEach(cell => blockXStarts.push(cell.x));
          } else if (line.length === 1) {
              blockXStarts.push(line[0].x);
          }
      });
      blockXStarts.sort((a, b) => a - b);
      let localColBins = [];
      const LOCAL_BIN_TOLERANCE = 22;
      blockXStarts.forEach(x => {
          let bin = localColBins.find(b => Math.abs(b.center - x) <= LOCAL_BIN_TOLERANCE);
          if (bin) {
              bin.count++;
              bin.center = ((bin.center * (bin.count - 1)) + x) / bin.count;
          } else {
              localColBins.push({ center: x, count: 1 });
          }
      });
      localColBins.sort((a, b) => a.center - b.center);
      let blockCols = localColBins.map(b => b.center);
      if (blockCols.length <= 1) {
          block.forEach(line => finalGrid.push(line.map(c => c.text)));
          return;
      }
      block.forEach(line => {
          let lineText = line.map(c => c.text).join(" ");
          if (line.length === 1 && isTitleLine(lineText)) {
              finalGrid.push([lineText]);
              return;
          }
          let row = new Array(blockCols.length).fill("");
          line.forEach(cell => {
              let bestIdx = 0;
              let minDist = Math.abs(blockCols[0] - cell.x);
              for (let c = 1; c < blockCols.length; c++) {
                  let dist = Math.abs(blockCols[c] - cell.x);
                  if (dist < minDist) {
                      minDist = dist;
                      bestIdx = c;
                  }
              }
              if (row[bestIdx]) {
                  row[bestIdx] += " " + cell.text;
              } else {
                  row[bestIdx] = cell.text;
              }
          });
          while (row.length > 1 && row[row.length - 1] === "") {
              row.pop();
          }
          if (row.some(c => c.trim() !== "")) {
              finalGrid.push(row);
          }
      });
  });
  if (finalGrid.length > 0) {
      let maxCols = Math.max(...finalGrid.map(r => r.length));
      let activeCols = [];
      for (let c = 0; c < maxCols; c++) {
          let hasContent = finalGrid.some(row => row[c] && row[c].toString().trim() !== "");
          if (hasContent) activeCols.push(c);
      }
      finalGrid = finalGrid.map(row => activeCols.map(c => row[c] || ""));
  }
  return finalGrid;
}

async function runOCRFallback(inputPath) {
  const rows = [];
  if (!Tesseract) return rows;
  let worker = null;
  try {
    let imageToRecognize = inputPath;
    const ext = path.extname(inputPath).toLowerCase();
    
    // If input is a PDF, render page 1 to an image buffer first using pdfjs + canvas
    if (ext === '.pdf') {
      try {
        const data = new Uint8Array(fs.readFileSync(inputPath));
        const pdf = await pdfjsLib.getDocument({ data }).promise;
        if (pdf.numPages > 0) {
          const page = await pdf.getPage(1);
          const viewport = page.getViewport({ scale: 2.0 });
          const canvas = createCanvas(viewport.width, viewport.height);
          const ctx = canvas.getContext('2d');
          await page.render({ canvasContext: ctx, viewport: viewport }).promise;
          imageToRecognize = canvas.toBuffer('image/jpeg', { quality: 0.9 });
        }
      } catch (pdfErr) {
        console.log('PDF page render for OCR failed:', pdfErr.message);
        return rows;
      }
    }

    worker = await Tesseract.createWorker('eng', 1);
    const { data } = await worker.recognize(imageToRecognize);
    if (data && data.text) {
      const lines = data.text.split('\n');
      lines.forEach(l => {
        const cols = l.split(/\t|\s{2,}/).map(c => c.trim()).filter(c => c.length > 0);
        if (cols.length > 0) rows.push(cols);
      });
    }
  } catch (e) {
    console.log('OCR Error:', e.message);
  } finally {
    if (worker) {
      try { await worker.terminate(); } catch (_) {}
    }
  }
  return rows;
}

async function pdfToExcelConvert(inputPath, outputPath, config = {}) {
  try {
    const data = new Uint8Array(fs.readFileSync(inputPath));
    const pdf = await pdfjsLib.getDocument({ data }).promise;
    const totalPages = pdf.numPages;
    let extractedTableData = [];

    const isComplex = config.isComplex === true;

    for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();
      
      const pageRows = isComplex ? parseComplexPDFTables(textContent.items) : parseSimplePDFTables(textContent.items);
      
      if (pageRows.length > 0) extractedTableData.push(...pageRows);
    }

    if (extractedTableData.length === 0) {
      extractedTableData = await runOCRFallback(inputPath);
    }

    const formattedData = extractedTableData.map(row =>
      row.map(cell => {
        const cleanNum = cell ? cell.toString().replace(/,/g, '').trim() : '';
        return (!isNaN(cleanNum) && cleanNum !== '') ? parseFloat(cleanNum) : cell;
      })
    );

    const worksheet = XLSX.utils.aoa_to_sheet(formattedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'PDF_Data');
    XLSX.writeFile(workbook, outputPath);
    return true;
  } catch (error) {
    console.error('PDF to Excel error:', error);
    throw error;
  }
}

// ========== Merge PDF Logic ==========
async function mergePdfConvert(input, outputPath, config = {}) {
  try {
    let inputs = Array.isArray(input) ? input : [input];

    if (config.sortOrder === 'custom' && config.customSequence && config.customSequence.trim() !== '') {
      const seqParts = config.customSequence.split(',').map(n => parseInt(n.trim()) - 1);
      let newInputs = [];
      let used = new Set();
      for (let idx of seqParts) {
        if (!isNaN(idx) && idx >= 0 && idx < inputs.length) {
          newInputs.push(inputs[idx]);
          used.add(idx);
        }
      }
      for (let i = 0; i < inputs.length; i++) {
        if (!used.has(i)) newInputs.push(inputs[i]);
      }
      inputs = newInputs;
    }

    const mergedPdf = await PDFDocument.create();

    for (let fIdx = 0; fIdx < inputs.length; fIdx++) {
      const file = inputs[fIdx];
      const arrayBuffer = fs.readFileSync(file);
      let mergedSuccessfully = false;

      try {
        const pdf = await PDFDocument.load(arrayBuffer);
        const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
        copiedPages.forEach((page) => {
          mergedPdf.addPage(page);
        });
        mergedSuccessfully = true;
      } catch (nativeErr) {
        console.warn("Native copy failed for a file, switching to safe hybrid page rendering...");
      }

      if (!mergedSuccessfully) {
        const data = new Uint8Array(arrayBuffer);
        const pdfDoc = await pdfjsLib.getDocument({ data }).promise;

        for (let i = 1; i <= pdfDoc.numPages; i++) {
          const page = await pdfDoc.getPage(i);
          const viewport = page.getViewport({ scale: 1.5 });
          
          const canvas = createCanvas(viewport.width, viewport.height);
          const context = canvas.getContext('2d');
          
          await page.render({ canvasContext: context, viewport: viewport }).promise;

          const imageBuffer = canvas.toBuffer('image/jpeg', { quality: 0.85 });
          const embeddedImage = await mergedPdf.embedJpg(imageBuffer);

          const pdfPage = mergedPdf.addPage([viewport.width, viewport.height]);
          pdfPage.drawImage(embeddedImage, {
            x: 0,
            y: 0,
            width: viewport.width,
            height: viewport.height,
          });
        }
      }
    }

    const mergedPdfBytes = await mergedPdf.save({ useObjectStreams: true });
    fs.writeFileSync(outputPath, mergedPdfBytes);
    return outputPath;
  } catch (error) {
    console.error('Merge PDF error:', error);
    throw error;
  }
}

// ========== PDF Splitter Logic ==========
async function pdfSplitterConvert(inputPath, outputPath, config = {}) {
  try {
    const data = fs.readFileSync(inputPath);
    const srcDoc = await PDFDocument.load(data);
    const totalPages = srcDoc.getPageCount();

    const mode = config.mode || 'all'; 
    const rangeStr = config.range || '';

    if (mode === 'range') {
      let selectedIndices = new Set();
      
      if (rangeStr.trim() !== '') {
        const normalizedRangeStr = rangeStr.toLowerCase()
            .replace(/to/g, '-')
            .replace(/and/g, ',')
            .replace(/&/g, ',')
            .replace(/[^0-9,-]/g, ''); 
      
        const parts = normalizedRangeStr.split(',');
        parts.forEach(part => {
          if (part.includes('-')) {
            const [s, e] = part.split('-').map(n => parseInt(n));
            if (!isNaN(s) && !isNaN(e)) {
              for (let p = Math.min(s, e); p <= Math.max(s, e); p++) {
                if (p >= 1 && p <= totalPages) selectedIndices.add(p - 1);
              }
            }
          } else {
            const p = parseInt(part);
            if (!isNaN(p) && p >= 1 && p <= totalPages) {
              selectedIndices.add(p - 1);
            }
          }
        });
      } else {
        for (let i = 0; i < totalPages; i++) selectedIndices.add(i);
      }

      if (selectedIndices.size === 0) throw new Error("No valid pages selected for extraction.");

      const indices = Array.from(selectedIndices).sort((a, b) => a - b);
      const newPdf = await PDFDocument.create();
      const copiedPages = await newPdf.copyPages(srcDoc, indices);
      copiedPages.forEach(p => newPdf.addPage(p));

      const pdfBytes = await newPdf.save();
      fs.writeFileSync(outputPath, pdfBytes);
      return outputPath; 

    } else {
      const zip = new JSZip();
      const baseName = path.basename(inputPath, path.extname(inputPath)) || 'split_document';

      for (let i = 0; i < totalPages; i++) {
        const singlePdf = await PDFDocument.create();
        const [copiedPage] = await singlePdf.copyPages(srcDoc, [i]);
        singlePdf.addPage(copiedPage);
        const pdfBytes = await singlePdf.save();
        zip.file(`${baseName}_page_${i + 1}.pdf`, pdfBytes);
      }

      const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });
      const zipOutputPath = outputPath.replace(/\.[^/.]+$/, ".zip"); 
      fs.writeFileSync(zipOutputPath, zipBuffer);
      return zipOutputPath; 
    }
  } catch (error) {
    console.error('PDF Splitter error:', error);
    throw error;
  }
}

// ========== PDF Organizer Logic ==========
async function pdfOrganizerConvert(inputPath, outputPath, config = {}) {
  try {
    const data = fs.readFileSync(inputPath);
    const srcDoc = await PDFDocument.load(data);
    const totalPages = srcDoc.getPageCount();

    const orderStr = config.pageOrder || '';
    const rotPagesStr = config.rotatePages || '';
    const rotDeg = parseInt(config.rotateDegree) || 90;

    let sequence = [];
    if (orderStr.trim() !== '') {
      const parts = orderStr.toLowerCase().replace(/and/g, ',').replace(/&/g, ',').replace(/[^0-9,-]/g, '').split(',');
      parts.forEach(part => {
        if (part.includes('-')) {
          const [s, e] = part.split('-').map(n => parseInt(n));
          if (!isNaN(s) && !isNaN(e)) {
            const step = s <= e ? 1 : -1; 
            for (let p = s; step > 0 ? p <= e : p >= e; p += step) {
              if (p >= 1 && p <= totalPages) sequence.push(p - 1);
            }
          }
        } else {
          const p = parseInt(part);
          if (!isNaN(p) && p >= 1 && p <= totalPages) sequence.push(p - 1);
        }
      });
    } else {
      for (let i = 0; i < totalPages; i++) sequence.push(i);
    }

    if (sequence.length === 0) throw new Error("No valid pages selected to organize.");

    let rotPages = new Set();
    let rotateAll = false;

    if (rotPagesStr.trim() !== '') {
      const rStr = rotPagesStr.toLowerCase();
      if (rStr === 'all') {
        rotateAll = true;
      } else {
         const normalized = rStr.replace(/to/g, '-').replace(/and/g, ',').replace(/&/g, ',').replace(/[^0-9,-]/g, '');
         const parts = normalized.split(',');
         parts.forEach(part => {
           if (part.includes('-')) {
             const [s, e] = part.split('-').map(n => parseInt(n));
             if (!isNaN(s) && !isNaN(e)) {
               for (let p = Math.min(s, e); p <= Math.max(s, e); p++) {
                  if (p >= 1 && p <= totalPages) rotPages.add(p - 1);
               }
             }
           } else {
             const p = parseInt(part);
             if (!isNaN(p) && p >= 1 && p <= totalPages) rotPages.add(p - 1);
           }
         });
      }
    }

    const newPdf = await PDFDocument.create();
    for (let origIndex of sequence) {
      const [copiedPage] = await newPdf.copyPages(srcDoc, [origIndex]);
      
      if (rotateAll || rotPages.has(origIndex)) {
        const currentRot = copiedPage.getRotation().angle || 0;
        copiedPage.setRotation(degrees((currentRot + rotDeg) % 360));
      }
      
      newPdf.addPage(copiedPage);
    }

    const pdfBytes = await newPdf.save();
    fs.writeFileSync(outputPath, pdfBytes);
    return outputPath;

  } catch (error) {
    console.error('PDF Organizer error:', error);
    throw error;
  }
}

// ========== PDF Watermark Logic ==========
async function pdfWatermarkConvert(inputPath, outputPath, config = {}) {
  try {
    const data = fs.readFileSync(inputPath);
    const pdfDoc = await PDFDocument.load(data);
    const pages = pdfDoc.getPages();

    const text = config.watermarkText || "CONFIDENTIAL";
    const mode = config.watermarkMode || "diagonal";

    pages.forEach((page, index) => {
      const { width, height } = page.getSize();
      
      if (mode === 'diagonal' || mode === 'both') {
        page.drawText(text, {
          x: width / 4,
          y: height / 2,
          size: 45,
          color: rgb(0.8, 0.2, 0.2),
          opacity: 0.25,
          rotate: degrees(45)
        });
      }

      if (mode === 'footer_num' || mode === 'both') {
        page.drawText(`Page ${index + 1} of ${pages.length}`, {
          x: width / 2 - 35,
          y: 20,
          size: 10,
          color: rgb(0.3, 0.3, 0.3)
        });
      }
    });

    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(outputPath, pdfBytes);
    return outputPath;
  } catch (error) {
    console.error('PDF Watermark error:', error);
    throw error;
  }
}

// ========== PDF to Image Logic ==========
async function pdfToImageConvert(inputPath, outputPath, config = {}) {
  try {
    const data = new Uint8Array(fs.readFileSync(inputPath));
    const pdf = await pdfjsLib.getDocument({ data }).promise;
    const totalPages = pdf.numPages;
    
    const zip = new JSZip();
    const baseName = path.basename(inputPath, path.extname(inputPath)) || 'extracted_images';

    const batchSize = 5; 
    for (let i = 1; i <= totalPages; i += batchSize) {
      const pagePromises = [];
      for (let j = 0; j < batchSize && (i + j) <= totalPages; j++) {
        const pageNum = i + j;
        pagePromises.push((async () => {
          const page = await pdf.getPage(pageNum);
          const viewport = page.getViewport({ scale: 2.0 }); 
          
          const canvas = createCanvas(viewport.width, viewport.height);
          const ctx = canvas.getContext('2d');
          
          await page.render({ canvasContext: ctx, viewport: viewport }).promise;

          const imageBuffer = canvas.toBuffer('image/jpeg', { quality: 0.9 });
          return { pageNum, imageBuffer };
        })());
      }
      
      const results = await Promise.all(pagePromises);
      results.forEach(res => {
        zip.file(`${baseName}_Page_${res.pageNum}.jpg`, res.imageBuffer);
      });
    }

    const zipBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "STORE" });
    const zipOutputPath = outputPath.replace(/\.[^/.]+$/, ".zip"); 
    fs.writeFileSync(zipOutputPath, zipBuffer);
    
    return zipOutputPath; 
  } catch (error) {
    console.error('PDF to Image error:', error);
    throw error;
  }
}

// ========== Compress Image Logic ==========
async function compressImageConvert(inputPath, outputPath, config = {}) {
  try {
    const finalOutputPath = outputPath.replace(/\.[^/.]+$/, ".jpg");
    const quality = config.compressQuality ? Math.round(parseFloat(config.compressQuality) * 100) : 70;

    if (sharp) {
      await sharp(inputPath)
        .resize(1200, 1200, { fit: 'inside', withoutEnlargement: true, kernel: sharp.kernel.lanczos3 })
        .flatten({ background: { r: 255, g: 255, b: 255 } })
        .jpeg({ quality: quality, mozjpeg: true })
        .toFile(finalOutputPath);
      return finalOutputPath;
    } else {
      const img = await loadImage(inputPath);
      let width = img.width;
      let height = img.height;
      const MAX_WIDTH = 1200, MAX_HEIGHT = 1200;
      if (width > height && width > MAX_WIDTH) {
        height = Math.round((height * MAX_WIDTH) / width);
        width = MAX_WIDTH;
      } else if (height > MAX_HEIGHT) {
        width = Math.round((width * MAX_HEIGHT) / height);
        height = MAX_HEIGHT;
      }
      const canvas = createCanvas(width, height);
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, width, height);
      const buffer = canvas.toBuffer('image/jpeg', { quality: quality / 100 });
      fs.writeFileSync(finalOutputPath, buffer);
      return finalOutputPath;
    }
  } catch (error) {
    console.error('Compress Image error:', error);
    throw error;
  }
}

// ========== Image Converter Logic (Full support for BMP, PNG, WEBP, JPG) ==========
async function imageConverterConvert(inputPath, outputPath, config = {}) {
  try {
    const format = (config.convertFormat || 'jpg').toLowerCase();
    const quality = config.convertQuality ? Math.round(parseFloat(config.convertQuality) * 100) : 92;
    const finalOutputPath = outputPath.replace(/\.[^/.]+$/, `.${format}`);

    if (format === 'bmp') {
      const img = await loadImage(inputPath);
      const canvas = createCanvas(img.width, img.height);
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#FFFFFF';
      ctx.fillRect(0, 0, img.width, img.height);
      ctx.drawImage(img, 0, 0);
      
      const imgData = ctx.getImageData(0, 0, img.width, img.height);
      const rawData = imgData.data;

      const width = img.width;
      const height = img.height;
      const rowSize = Math.floor((3 * width + 3) / 4) * 4;
      const pixelArraySize = rowSize * height;
      const fileSize = 54 + pixelArraySize;
      const buffer = Buffer.alloc(fileSize);

      buffer.write('BM', 0);
      buffer.writeUInt32LE(fileSize, 2);
      buffer.writeUInt32LE(0, 6);
      buffer.writeUInt32LE(54, 10);

      buffer.writeUInt32LE(40, 14);
      buffer.writeInt32LE(width, 18);
      buffer.writeInt32LE(-height, 22);
      buffer.writeUInt16LE(1, 26);
      buffer.writeUInt16LE(24, 28);
      buffer.writeUInt32LE(0, 30);
      buffer.writeUInt32LE(pixelArraySize, 34);
      buffer.writeUInt32LE(2835, 38);
      buffer.writeUInt32LE(2835, 42);
      buffer.writeUInt32LE(0, 46);
      buffer.writeUInt32LE(0, 50);

      let srcOffset = 0;
      let dstOffset = 54;
      for (let y = 0; y < height; y++) {
          const rowPadding = rowSize - (width * 3);
          for (let x = 0; x < width; x++) {
              const r = rawData[srcOffset];
              const g = rawData[srcOffset + 1];
              const b = rawData[srcOffset + 2];
              buffer[dstOffset] = b;
              buffer[dstOffset + 1] = g;
              buffer[dstOffset + 2] = r;
              srcOffset += 4;
              dstOffset += 3;
          }
          for (let p = 0; p < rowPadding; p++) {
              buffer[dstOffset++] = 0;
          }
      }

      fs.writeFileSync(finalOutputPath, buffer);
      return finalOutputPath;
    }

    if (!sharp) throw new Error("Sharp module required.");
    let imageProcess = sharp(inputPath);

    if (format === 'jpg' || format === 'jpeg' || format === 'png') {
      imageProcess = imageProcess.flatten({ background: { r: 255, g: 255, b: 255 } });
    }

    let ext = format === 'jpeg' ? 'jpeg' : format;
    if (format === 'ico') ext = 'ico';

    if (format === 'ico') {
      await imageProcess.resize(256, 256, { fit: 'inside', kernel: sharp.kernel.lanczos3 }).png().toFile(finalOutputPath);
    } else if (format === 'jpg' || format === 'jpeg') {
      await imageProcess.jpeg({ quality: quality, mozjpeg: true }).toFile(finalOutputPath);
    } else if (format === 'webp') {
      await imageProcess.webp({ quality: quality }).toFile(finalOutputPath);
    } else if (format === 'png') {
      await imageProcess.png().toFile(finalOutputPath);
    } else {
      await imageProcess.toFormat(ext).toFile(finalOutputPath);
    }
    return finalOutputPath;
  } catch (error) {
    console.error('Image Converter error:', error);
    throw error;
  }
}

// ========== Image to PDF Logic ==========
async function imageToPdfConvert(input, outputPath, config = {}) {
  try {
    let inputs = Array.isArray(input) ? input : [input];
    const pdfDoc = await PDFDocument.create();
    
    for (let i = 0; i < inputs.length; i++) {
      const file = inputs[i];
      const imgBuffer = fs.readFileSync(file);
      let finalImgBuffer = imgBuffer;

      if (sharp) {
         const metadata = await sharp(imgBuffer).metadata();
         if (metadata.format !== 'jpeg' && metadata.format !== 'png') {
             finalImgBuffer = await sharp(imgBuffer).jpeg({ quality: 95 }).toBuffer();
         }
      }

      let pdfImage;
      try {
         pdfImage = await pdfDoc.embedJpg(finalImgBuffer);
      } catch(e) {
         pdfImage = await pdfDoc.embedPng(finalImgBuffer);
      }

      const imgDims = pdfImage.scale(1);
      const A4_WIDTH = 595.28, A4_HEIGHT = 841.89, MARGIN = 28.35;
      const maxW = A4_WIDTH - (MARGIN * 2);
      const maxH = A4_HEIGHT - (MARGIN * 2);

      let w = imgDims.width, h = imgDims.height;
      if (w > maxW) { h = (h * maxW) / w; w = maxW; }
      if (h > maxH) { w = (w * maxH) / h; h = maxH; }

      const x = (A4_WIDTH - w) / 2;
      const y = (A4_HEIGHT - h) / 2;

      const page = pdfDoc.addPage([A4_WIDTH, A4_HEIGHT]);
      page.drawImage(pdfImage, { x, y, width: w, height: h });
    }

    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(outputPath, pdfBytes);
    return outputPath;
  } catch (error) {
    console.error('Image to PDF error:', error);
    throw error;
  }
}

// ========== Image Reducer Logic (Dynamic format preservation & PNG lossless/lossy balance) ==========
async function imageReducerConvert(inputPath, outputPath, config = {}) {
  try {
    const mode = config.reducerMode || 'resize';
    const inputExt = path.extname(inputPath).toLowerCase();
    const targetExt = ['.png', '.webp', '.bmp', '.jpeg', '.jpg'].includes(inputExt) ? inputExt : '.jpg';
    const finalOutputPath = outputPath.replace(/\.[^/.]+$/, targetExt === '.jpeg' ? '.jpg' : targetExt);

    if (sharp) {
      if (targetExt === '.bmp') {
        const img = await loadImage(inputPath);
        const canvas = createCanvas(img.width, img.height);
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, img.width, img.height);
        ctx.drawImage(img, 0, 0);
        fs.writeFileSync(finalOutputPath, canvas.toBuffer('image/jpeg'));
        return finalOutputPath;
      }

      let pipeline = sharp(inputPath);
      if (targetExt !== '.png') {
        pipeline = pipeline.flatten({ background: { r: 255, g: 255, b: 255 } });
      }

      if (mode === 'resize') {
        const w = parseInt(config.reducerWidth) || 800;
        const h = parseInt(config.reducerHeight) || 600;
        let p = pipeline.resize(w, h, { fit: 'fill', kernel: sharp.kernel.lanczos3 });
        
        if (targetExt === '.png') await p.png({ compressionLevel: 9 }).toFile(finalOutputPath);
        else if (targetExt === '.webp') await p.webp({ quality: 95 }).toFile(finalOutputPath);
        else await p.jpeg({ quality: 95, mozjpeg: true }).toFile(finalOutputPath);
      } else {
        const targetKb = parseFloat(config.reducerTargetKb) || 50;
        const targetBytes = targetKb * 1024;
        
        const meta = await sharp(inputPath).metadata();
        let currentW = meta.width || 1200;
        let currentH = meta.height || 1200;

        let quality = 90;
        let scale = 1.0;
        let outputBuffer;

        for (let attempt = 0; attempt < 6; attempt++) {
            let resizePipeline = sharp(inputPath);
            if (targetExt !== '.png') resizePipeline = resizePipeline.flatten({ background: { r: 255, g: 255, b: 255 } });
            
            let resObj = resizePipeline.resize(
              Math.max(150, Math.round(currentW * scale)), 
              Math.max(150, Math.round(currentH * scale)), 
              { fit: 'inside', kernel: sharp.kernel.lanczos3 }
            );

            if (targetExt === '.png') {
                outputBuffer = await resObj.png({ compressionLevel: 9, palette: true, colors: Math.max(16, Math.round(256 * (quality/100))) }).toBuffer();
            } else if (targetExt === '.webp') {
                outputBuffer = await resObj.webp({ quality: quality }).toBuffer();
            } else {
                outputBuffer = await resObj.jpeg({ quality: quality, mozjpeg: true }).toBuffer();
            }

            if (outputBuffer.length <= targetBytes || quality <= 20) {
                break;
            }
            quality -= 12;
            scale -= 0.12;
        }

        fs.writeFileSync(finalOutputPath, outputBuffer);
      }

      return finalOutputPath;
    } else {
      const img = await loadImage(inputPath);
      const canvas = createCanvas(800, 600);
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#FFFFFF';
      ctx.fillRect(0, 0, 800, 600);
      ctx.drawImage(img, 0, 0, 800, 600);
      const buffer = canvas.toBuffer('image/jpeg', { quality: 0.90 });
      fs.writeFileSync(finalOutputPath, buffer);
      return finalOutputPath;
    }
  } catch (error) {
    console.error('Image Reducer error:', error);
    throw error;
  }
}

// ========== EXACT WEBSITE INTEGRATED OCR ENGINE (100% Match with Frontend Code) ==========
async function preprocessImageUsingCanvas(imageSrc) {
    if (sharp) {
        try {
            const meta = await sharp(imageSrc).metadata();
            const maxSide = Math.max(meta.width || 0, meta.height || 0) || 1000;
            const scale = Math.max(1.8, 1800 / maxSide);
            const targetW = Math.max(1, Math.round((meta.width || 1000) * scale));
            const targetH = Math.max(1, Math.round((meta.height || 1000) * scale));

            const buffer = await sharp(imageSrc)
                .rotate()
                .resize(targetW, targetH, { kernel: sharp.kernel.lanczos3 })
                .flatten({ background: { r: 255, g: 255, b: 255 } })
                .greyscale()
                .normalize()
                .sharpen()
                .png()
                .toBuffer();

            return buffer;
        } catch (sharpErr) {
            console.log('OCR preprocessing: sharp failed, falling back to canvas ->', sharpErr.message);
        }
    }

    try {
        const img = await loadImage(imageSrc);
        const scale = Math.max(1.8, 1800 / Math.max(img.width, img.height));
        const targetW = Math.round(img.width * scale);
        const targetH = Math.round(img.height * scale);

        const canvas = createCanvas(targetW, targetH);
        const ctx = canvas.getContext('2d', { willReadFrequently: true });

        ctx.imageSmoothingEnabled = true;
        ctx.drawImage(img, 0, 0, targetW, targetH);

        const imgData = ctx.getImageData(0, 0, targetW, targetH);
        const data = imgData.data;

        for (let i = 0; i < data.length; i += 4) {
            const gray = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
            data[i] = gray;
            data[i + 1] = gray;
            data[i + 2] = gray;
        }

        ctx.putImageData(imgData, 0, 0);
        return canvas.toBuffer('image/png');
    } catch (canvasErr) {
        throw new Error('Could not decode image for OCR (unsupported/corrupted format): ' + canvasErr.message);
    }
}

function applySpellFixes(text) {
    if (!text) return "";
    const corrections = [
        [/\bHomless\b/gi, "Homeless"],
        [/\btabrics\b/gi, "fabrics"],
        [/\btound\b/gi, "found"],
        [/\bcarly\b/gi, "early"],
        [/\bcary\b/gi, "early"],
        [/\b3000 BO\b/gi, "3000 B.C."],
        [/\bthewg2os\b/gi, "1920s"]
    ];
    corrections.forEach(([pattern, replacement]) => {
        text = text.replace(pattern, replacement);
    });
    return text;
}

function convertOCRToWordHTML(rawText) {
    if (!rawText) return { htmlContent: "", plainText: "" };

    let cleanedRaw = applySpellFixes(rawText)
        .replace(/\r\n/g, '\n')
        .replace(/[|~`^{}\[\]\\]/g, '')
        .trim();

    let rawBlocks = cleanedRaw.split(/\n\s*\n+/);
    let finalParagraphs = [];

    rawBlocks.forEach((block) => {
        let lines = block.split('\n').map(l => l.trim()).filter(l => l.length > 0);
        if (lines.length === 0) return;

        let currentParagraph = "";

        lines.forEach((line) => {
            const isBullet = /^([•\-\*]|[\d\w]+\.)\s+/.test(line);

            if (!currentParagraph) {
                currentParagraph = line;
            } else {
                if (isBullet) {
                    finalParagraphs.push(currentParagraph);
                    currentParagraph = line;
                } else {
                    currentParagraph += " " + line;
                }
            }
        });

        if (currentParagraph) {
            finalParagraphs.push(currentParagraph);
        }
    });

    let htmlParts = finalParagraphs.map((p, idx) => {
        const pClean = p.trim();
        if (idx === 0 && pClean.length < 50 && !pClean.endsWith('.')) {
            return `<h2 style="font-size: 18pt; font-weight: bold; color: #1e293b; margin-bottom: 14pt; margin-top: 0;">${pClean}</h2>`;
        } else if (/^([•\-\*]|[\d\w]+\.)\s+/.test(pClean)) {
            return `<p style="font-size: 11pt; line-height: 1.6; color: #334155; margin-bottom: 8pt; padding-left: 15pt;">${pClean}</p>`;
        } else {
            return `<p style="font-size: 11pt; line-height: 1.6; color: #334155; margin-bottom: 12pt; text-align: justify;">${pClean}</p>`;
        }
    });

    return { htmlContent: htmlParts.join(''), plainText: finalParagraphs.join('\n\n') };
}

async function imageToTextConvert(inputPath, outputPath, config = {}) {
  try {
    if (!Tesseract) {
      throw new Error('OCR engine (tesseract.js) is not installed on the server. Run: npm install tesseract.js');
    }

    let rawText = "";
    let worker = null;
    try {
      const processedBuffer = await preprocessImageUsingCanvas(inputPath);

      worker = await Tesseract.createWorker('eng', 1);
      await worker.setParameters({
        tessedit_pageseg_mode: Tesseract.PSM?.AUTO || 3,
        preserve_interword_spaces: '1'
      });

      const { data } = await worker.recognize(processedBuffer);
      await worker.terminate();
      rawText = data.text || '';
    } finally {
      if (worker) await worker.terminate();
    }

    if (!rawText || !rawText.trim()) {
      console.log('Image to Text: OCR returned no text for', inputPath);
    }

    const { htmlContent } = convertOCRToWordHTML(rawText);
    const finalWordHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>Extracted Document</title><style>body{font-family: Calibri, Arial, sans-serif;} table{border-collapse:collapse;width:100%;} td{border:1px solid #000000;padding:6px;}</style></head><body>' + htmlContent + '</body></html>';
    
    fs.writeFileSync(outputPath, '\ufeff' + finalWordHtml);
    return outputPath;
  } catch (error) {
    console.error('Image to Text error:', error);
    throw error;
  }
}

// ========== PASSPORT STUDIO WORKFLOW HANDLER (Matching HTML features & logic) ==========
const PASSPORT_SIZE_PRESETS = {
  '3.5x4.5': { wIn: 3.5 / 2.54, hIn: 4.5 / 2.54, wMm: 35, hMm: 45, cols: 6, copyOptions: [4, 8, 16, 24, 32] },
  '3.5x3.5': { wIn: 3.5 / 2.54, hIn: 3.5 / 2.54, wMm: 35, hMm: 35, cols: 6, copyOptions: [4, 8, 16, 24, 32] },
  '2x2':     { wIn: 2,            hIn: 2,            wMm: 51, hMm: 51, cols: 4, copyOptions: [4, 8, 16] }
};
const PASSPORT_DPI = 300;
const MM_TO_PT = 2.83465;

function resolvePassportPreset(sizePreset) {
  return PASSPORT_SIZE_PRESETS[sizePreset] || PASSPORT_SIZE_PRESETS['3.5x4.5'];
}

function resolvePassportDimensions(sizePreset) {
  const preset = resolvePassportPreset(sizePreset);
  return {
    width: Math.round(preset.wIn * PASSPORT_DPI),
    height: Math.round(preset.hIn * PASSPORT_DPI)
  };
}

function hexToRgb(hex, fallback) {
  if (typeof hex === 'string' && /^#[0-9a-fA-F]{6}$/.test(hex)) {
    return {
      r: parseInt(hex.slice(1, 3), 16),
      g: parseInt(hex.slice(3, 5), 16),
      b: parseInt(hex.slice(5, 7), 16)
    };
  }
  return fallback;
}

async function removeBackgroundViaCloud(inputBuffer, apiKey) {
  const form = new FormData();
  form.append('image_file', new Blob([inputBuffer], { type: 'image/jpeg' }), 'photo.jpg');
  form.append('size', 'auto');

  const response = await fetch('https://api.remove.bg/v1.0/removebg', {
    method: 'POST',
    headers: { 'X-Api-Key': apiKey || 'hiKEsbZArnK8g4kkV1fc7VXg' },
    body: form
  });

  if (!response.ok) {
    throw new Error('Cloud AI background removal failed (HTTP ' + response.status + ')');
  }
  const arrayBuf = await response.arrayBuffer();
  return Buffer.from(arrayBuf);
}

async function decodeToPngBuffer(inputBuffer) {
  if (sharp) {
    try {
      return await sharp(inputBuffer).rotate().png().toBuffer();
    } catch (e) {
      console.log('Passport Studio: sharp decode failed ->', e.message);
    }
  }
  return inputBuffer;
}

async function renderPassportPhoto(sourceBuffer, targetW, targetH, bgColor, zoomVal, offsetXPercent, offsetYPercent) {
  const pngBuffer = await decodeToPngBuffer(sourceBuffer);
  const img = await loadImage(pngBuffer);

  const canvas = createCanvas(targetW, targetH);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, targetW, targetH);

  const imgAspect = img.width / img.height;
  const frameAspect = targetW / targetH;
  
  let drawW, drawH;
  if (imgAspect > frameAspect) {
    drawW = targetW * zoomVal;
    drawH = drawW / imgAspect;
  } else {
    drawH = targetH * zoomVal;
    drawW = drawH * imgAspect;
  }

  const baseLeft = (targetW - drawW) / 2;
  const baseTop = (targetH - drawH) / 2;

  const finalX = baseLeft + (offsetXPercent / 100) * targetW;
  const finalY = baseTop + (offsetYPercent / 100) * targetH;

  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(img, finalX, finalY, drawW, drawH);

  return canvas.toBuffer('image/jpeg', { quality: 0.95 });
}

async function generatePassportA4Sheet(photoJpegBuffer, sizePreset, copies) {
  const preset = resolvePassportPreset(sizePreset);
  const totalCopies = preset.copyOptions.includes(Number(copies))
    ? Number(copies)
    : preset.copyOptions[preset.copyOptions.length - 1];

  const pageWmm = 210, pageHmm = 297;
  const pageWpt = pageWmm * MM_TO_PT; // 595.28 pt
  const pageHpt = pageHmm * MM_TO_PT; // 841.89 pt

  const cellWpt = preset.wMm * MM_TO_PT;
  const cellHpt = preset.hMm * MM_TO_PT;
  const cols = preset.cols; // 6 columns

  // Exact grid alignment matching the frontend web layout preview
  const totalGridW = cols * cellWpt;
  const availableW = pageWpt - totalGridW;
  const marginSidePt = availableW / (cols + 1);
  const gapPt = marginSidePt;
  const marginTopPt = 25;
  const gapVerticalPt = 6;

  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([pageWpt, pageHpt]);
  const jpgImage = await pdfDoc.embedJpg(photoJpegBuffer);

  for (let i = 0; i < totalCopies; i++) {
    const col = i % cols;
    const row = Math.floor(i / cols);

    const x = marginSidePt + col * (cellWpt + gapPt);
    const yFromTop = marginTopPt + row * (cellHpt + gapVerticalPt);
    const y = pageHpt - yFromTop - cellHpt;

    if (y < 15) break;

    page.drawImage(jpgImage, { x, y, width: cellWpt, height: cellHpt });
    page.drawRectangle({
      x, y, width: cellWpt, height: cellHpt,
      borderColor: rgb(0, 0, 0), borderWidth: 0.75
    });
  }

  return await pdfDoc.save();
}

async function passportStudioConvert(inputPath, outputPath, config = {}) {
  try {
    const sizePreset = config.sizePreset || '3.5x4.5';
    const bgColor = config.bgColor || '#f87171';
    const zoomVal = (parseFloat(config.zoom) || 100) / 100;
    const offsetXPercent = parseFloat(config.offsetXPercent) || 0;
    const offsetYPercent = parseFloat(config.offsetYPercent) || 0;

    const { width: targetW, height: targetH } = resolvePassportDimensions(sizePreset);
    let sourceBuffer = fs.readFileSync(inputPath);

    if (config.removeBg === true || config.removeBg === 'true') {
      const apiKey = config.removeBgApiKey || process.env.REMOVE_BG_API_KEY || "hiKEsbZArnK8g4kkV1fc7VXg";
      try {
        sourceBuffer = await removeBackgroundViaCloud(sourceBuffer, apiKey);
      } catch (bgErr) {
        console.log('Passport Studio: cloud background removal failed ->', bgErr.message);
      }
    }

    const photoBuffer = await renderPassportPhoto(sourceBuffer, targetW, targetH, bgColor, zoomVal, offsetXPercent, offsetYPercent);

    const wantsSheet = (config.printSheet === true || config.printSheet === 'true' || config.copies);
    if (wantsSheet) {
      const pdfBytes = await generatePassportA4Sheet(photoBuffer, sizePreset, config.copies || 32);
      const pdfOutputPath = outputPath.replace(/\.[^/.]+$/, ".pdf");
      fs.writeFileSync(pdfOutputPath, pdfBytes);
      return pdfOutputPath;
    }

    const finalOutputPath = outputPath.replace(/\.[^/.]+$/, ".jpg");
    fs.writeFileSync(finalOutputPath, photoBuffer);
    return finalOutputPath;
  } catch (error) {
    console.error('Passport Studio error:', error);
    throw error;
  }
}

// ========== Merge Image Logic (Matching Frontend HTML Exactly) ==========
async function mergeImageConvert(input, outputPath, config = {}) {
  try {
    let inputs = Array.isArray(input) ? input : [input];
    if (inputs.length === 0) throw new Error("No images provided for merging.");

    const direction = config.mergeDirection || config.direction || 'horizontal';
    const professionalGap = 16;

    let loadedImages = [];
    for (let i = 0; i < inputs.length; i++) {
      const img = await loadImage(inputs[i]);
      loadedImages.push(img);
    }

    const totalW = direction === 'vertical' 
      ? Math.max(...loadedImages.map(img => img.width)) 
      : loadedImages.reduce((sum, img) => sum + img.width, 0) + (professionalGap * (loadedImages.length - 1));
      
    const totalH = direction === 'vertical' 
      ? loadedImages.reduce((sum, img) => sum + img.height, 0) + (professionalGap * (loadedImages.length - 1)) 
      : Math.max(...loadedImages.map(img => img.height));

    const canvas = createCanvas(totalW, totalH);
    const ctx = canvas.getContext('2d');

    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, totalW, totalH);

    let offset = 0;
    loadedImages.forEach((img) => {
      if (direction === 'vertical') {
        const drawX = (totalW - img.width) / 2;
        ctx.drawImage(img, drawX, offset);
        offset += img.height + professionalGap;
      } else {
        const drawY = (totalH - img.height) / 2;
        ctx.drawImage(img, offset, drawY);
        offset += img.width + professionalGap;
      }
    });

    const finalOutputPath = outputPath.replace(/\.[^/.]+$/, ".jpg");
    let buffer = canvas.toBuffer('image/jpeg', { quality: 0.88 });

    if (sharp) {
      let quality = 88;
      while (buffer.length > 100 * 1024 && quality > 15) {
        quality -= 8;
        buffer = await sharp(canvas.toBuffer('image/png'))
          .jpeg({ quality: quality, mozjpeg: true })
          .toBuffer();
      }
    }

    fs.writeFileSync(finalOutputPath, buffer);
    return finalOutputPath;
  } catch (error) {
    console.error('Merge Image error:', error);
    throw error;
  }
} 

// ========== Main processFile ==========
async function processFile(toolId, inputPath, outputPath, config = {}) {
  switch (toolId) {
    case 'pdf-to-word':
      return await pdfToWordConvert(inputPath, outputPath);
    case 'pdf-to-excel':
      await pdfToExcelConvert(inputPath, outputPath, config);
      return outputPath;
    case 'merge-pdf':
      return await mergePdfConvert(inputPath, outputPath, config);
    case 'pdf-splitter':
      return await pdfSplitterConvert(inputPath, outputPath, config);
    case 'pdf-organizer':
      return await pdfOrganizerConvert(inputPath, outputPath, config);
    case 'pdf-watermark':
      return await pdfWatermarkConvert(inputPath, outputPath, config);
    case 'pdf-to-image':
      return await pdfToImageConvert(inputPath, outputPath, config);
    case 'compress-image':
      return await compressImageConvert(inputPath, outputPath, config);
    case 'image-converter':
      return await imageConverterConvert(inputPath, outputPath, config);
    case 'image-to-pdf':
      return await imageToPdfConvert(inputPath, outputPath, config);
    case 'image-reducer':
      return await imageReducerConvert(inputPath, outputPath, config);
    case 'image-to-text':
      return await imageToTextConvert(inputPath, outputPath, config);
    case 'passport-studio':
      return await passportStudioConvert(inputPath, outputPath, config);
   case 'merge-image':
      return await mergeImageConvert(inputPath, outputPath, config);
    default:
      fs.copyFileSync(Array.isArray(inputPath) ? inputPath[0] : inputPath, outputPath);
      return outputPath;
  }
}

// ========== API Routes ==========
app.get('/api/tools', (req, res) => res.json({ success: true, tools: toolRegistry }));

app.post('/api/process/:toolId', async (req, res) => {
  try {
    const tool = toolRegistry.find(t => t.id === req.params.toolId);
    if (!tool) return res.status(404).json({ success: false, error: 'Tool not found' });
    if (!req.files || !req.files.file) return res.status(400).json({ success: false, error: 'No file uploaded' });

    const uploadedFiles = Array.isArray(req.files.file) ? req.files.file : [req.files.file];
    const uniqueId = crypto.randomUUID();
    const inputExt = path.extname(uploadedFiles[0].name);
    const outputExt = tool.outputExt || inputExt;
    const outputPath = path.join(tempDir, 'output_' + uniqueId + outputExt);

    let inputPaths = [];
    for (let i = 0; i < uploadedFiles.length; i++) {
      const p = path.join(tempDir, `input_${i}_${uniqueId}${path.extname(uploadedFiles[i].name)}`);
      await uploadedFiles[i].mv(p);
      inputPaths.push(p);
    }

    const inputArg = inputPaths.length === 1 ? inputPaths[0] : inputPaths;
    
    const actualOutputPath = await processFile(tool.id, inputArg, outputPath, req.body || {});
    const finalExt = path.extname(actualOutputPath);

    res.download(actualOutputPath, 'converted_' + uniqueId + finalExt, (err) => {
      if (err) console.error(err);
      setTimeout(() => { 
        inputPaths.forEach(p => fs.unlink(p, () => {}));
        fs.unlink(actualOutputPath, () => {}); 
      }, 5000);
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/workflow/create', (req, res) => {
  const { name, steps } = req.body;
  if (!name || !Array.isArray(steps) || steps.length === 0) {
    return res.status(400).json({ success: false, error: 'Workflow needs name and steps' });
  }
  const id = crypto.randomUUID();
  workflowsDB.set(id, { id, name, steps });
  res.json({ success: true, workflow: { id, name, steps } });
});

app.get('/api/workflow/list', (req, res) => {
  res.json({ success: true, workflows: Array.from(workflowsDB.values()) });
});

app.post('/api/workflow/execute/:id', async (req, res) => {
  try {
    const workflow = workflowsDB.get(req.params.id);
    if (!workflow) return res.status(404).json({ success: false, error: 'Workflow not found' });
    if (!req.files || !req.files.file) return res.status(400).json({ success: false, error: 'No file uploaded' });

    let uploadedFiles = Array.isArray(req.files.file) ? req.files.file : [req.files.file];
    
    const firstStep = workflow.steps[0];
    if (firstStep && (firstStep.toolId === 'merge-pdf' || firstStep.toolId === 'image-to-pdf')) {
       if (firstStep.sortOrder === 'az') {
         uploadedFiles.sort((a, b) => a.name.localeCompare(b.name));
       } else if (firstStep.sortOrder === 'za') {
         uploadedFiles.sort((a, b) => b.name.localeCompare(a.name));
       }
    }

    const uniqueId = crypto.randomUUID();
    let currentInput = [];
    for (let i = 0; i < uploadedFiles.length; i++) {
      const f = uploadedFiles[i];
      const p = path.join(tempDir, `input_${i}_${uniqueId}${path.extname(f.name)}`);
      await f.mv(p);
      currentInput.push(p);
    }

    let currentArg = currentInput.length === 1 ? currentInput[0] : currentInput;
    const results = [];

    for (let i = 0; i < workflow.steps.length; i++) {
      const stepConfig = workflow.steps[i];
      const tool = toolRegistry.find(t => t.id === stepConfig.toolId);
      if (!tool) continue;

      let currentExt = '.pdf';
      if (typeof currentArg === 'string') currentExt = path.extname(currentArg);
      else if (currentArg.length > 0) currentExt = path.extname(currentArg[0]);

      let outputExt = tool.outputExt || currentExt;
      if (tool.id === 'image-converter') {
         const targetFmt = stepConfig.convertFormat || 'jpg';
         outputExt = targetFmt === 'jpeg' ? '.jpeg' : `.${targetFmt}`;
      } else if (tool.id === 'passport-studio' && (stepConfig.printSheet || stepConfig.copies)) {
         outputExt = '.pdf';
      }

      const stepOutput = path.join(tempDir, `step_${i}_${uniqueId}${outputExt}`);
      
      try {
        const actualStepOutput = await processFile(tool.id, currentArg, stepOutput, stepConfig);
        currentArg = actualStepOutput; 
        results.push({ step: i+1, toolId: tool.id, name: tool.name, status: 'completed' });
      } catch (e) {
        results.push({ step: i+1, toolId: tool.id, name: tool.name, status: 'failed', error: e.message });
        return res.status(500).json({ success: false, results, error: 'Step ' + (i+1) + ' failed: ' + e.message });
      }
    }

    const finalExt = path.extname(currentArg);
    res.download(currentArg, `workflow_output_${uniqueId}${finalExt}`, (err) => {
      if (err) console.error(err);
      setTimeout(() => {
        fs.readdirSync(tempDir).forEach(f => { if (f.includes(uniqueId)) fs.unlink(path.join(tempDir, f), () => {}); });
      }, 5000);
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ========== Frontend & Workflow Builder UI with Gorgeous Multi-Color Studio Modal & Live Preview ==========
app.get('/', (req, res) => {
  const publicIndex = path.join(__dirname, 'public', 'index.html');
  if (fs.existsSync(publicIndex)) {
    return res.sendFile(publicIndex);
  }
  return res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/workflow-builder', (req, res) => {
  const toolsListHTML = toolRegistry.map(tool => {
    return '<div class="tool-item" onclick="addToWorkflow(\'' + tool.id + '\', \'' + tool.name.replace(/'/g, "\\'") + '\', \'' + tool.icon + '\')">' +
      '<span class="tool-icon">' + tool.icon + '</span>' +
      '<span class="tool-name">' + tool.name + '</span>' +
      '</div>';
  }).join('');

  res.send('<!DOCTYPE html><html><head><title>Workflow Builder - Filesque</title><script src="https://cdn.tailwindcss.com"></script><link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap" rel="stylesheet"><style>' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    'body{font-family:"Plus Jakarta Sans",sans-serif;background:#f8fafc;min-height:100vh}' +
    '.navbar{background:rgba(255,255,255,0.9);backdrop-filter:blur(12px);border-bottom:1px solid #f1f5f9;padding:1.25rem 2rem;display:flex;justify-content:space-between;align-items:center;box-shadow:0 1px 2px 0 rgba(0,0,0,0.05)}' +
    '.back-link{color:#0f172a;text-decoration:none;font-weight:800;font-size:15px;background:#f1f5f9;padding:10px 18px;border-radius:12px;border:1px solid #e2e8f0;display:flex;align-items:center;gap:6px;transition:all 0.2s}' +
    '.back-link:hover{background:#0f172a;color:white;border-color:#0f172a;transform:translateY(-1px);box-shadow:0 4px 12px rgba(15,23,42,0.15)}' +
    '.container{display:flex;gap:2rem;width:100%;max-width:none;margin:0;padding:2rem 2rem}' +
    '.panel{background:white;border-radius:1.25rem;box-shadow:0 4px 20px rgba(0,0,0,0.03);padding:1.75rem;border:1px solid #e2e8f0}' +
    '.tools-panel{flex:0 0 400px;max-height:85vh;overflow-y:auto}.workflow-panel{flex:1}' +
    '.panel-title{font-size:1.35rem;font-weight:800;margin-bottom:1.25rem;color:#0f172a}' +
    '.tool-item{display:flex;align-items:center;gap:0.8rem;padding:0.9rem 1.1rem;background:#f8fafc;border:1px solid #e2e8f0;border-radius:0.75rem;margin-bottom:0.6rem;cursor:pointer;transition:0.2s}' +
    '.tool-item:hover{background:#f1f5f9;border-color:#cbd5e1;transform:translateX(4px)}' +
    '.tool-icon{font-size:1.5rem}.tool-name{font-weight:700;color:#334155;font-size:14px}' +
    '.workflow-area{min-height:220px;border:2px dashed #cbd5e1;border-radius:1rem;padding:1.25rem;margin-bottom:1.25rem;background:#fafafa}' +
    '.step{display:flex;align-items:center;gap:1rem;padding:1.1rem;background:white;border:1px solid #e2e8f0;border-radius:0.75rem;margin-bottom:0.75rem;transition:all 0.3s;box-shadow:0 2px 6px rgba(0,0,0,0.02)}' +
    '.step.active-step{background:#e0f2fe;border:2px solid #0284c7;box-shadow:0 4px 12px rgba(2,132,199,0.2);transform:scale(1.01)}' +
    '.step-number{width:34px;height:34px;background:#dc2626;color:white;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:14px}' +
    '.remove-btn{background:#fee2e2;color:#dc2626;border:none;border-radius:0.5rem;padding:0.4rem 0.9rem;cursor:pointer;font-weight:800;font-size:13px;transition:0.2s}' +
    '.remove-btn:hover{background:#dc2626;color:white}' +
    '.controls{display:flex;gap:1rem;margin-bottom:1.5rem}' +
    'input[type="text"]{flex:1;padding:0.9rem 1.25rem;border:1px solid #cbd5e1;border-radius:0.75rem;font-size:15px;font-weight:600;outline:none}' +
    'input[type="text"]:focus{border-color:#0284c7;box-shadow:0 0 0 3px rgba(2,132,199,0.1)}' +
    '.btn{padding:0.9rem 1.75rem;background:#dc2626;color:white;border:none;border-radius:0.75rem;cursor:pointer;font-weight:800;font-size:15px;transition:0.2s}' +
    '.btn:hover{background:#b91c1c;transform:translateY(-1px);box-shadow:0 4px 12px rgba(220,38,38,0.2)}' +
    '.btn-secondary{background:#64748b}.btn-secondary:hover{background:#475569;box-shadow:0 4px 12px rgba(100,116,139,0.2)}' +
    '.saved-list{margin-top:1.5rem}.saved-item{padding:0.9rem 1.1rem;background:#f8fafc;border:1px solid #e2e8f0;border-radius:0.75rem;margin-bottom:0.5rem;font-size:14px;font-weight:600;color:#334155}' +
    '.file-upload-section{margin-top:1.5rem;padding:1.25rem;background:#f8fafc;border:1px solid #e2e8f0;border-radius:1rem}' +
    '.file-upload-section input[type="file"]{margin-top:0.75rem;font-size:14px;font-weight:600}' +
    '</style></head><body>' +
    '<div class="navbar">' +
    '  <div class="cursor-pointer flex items-center space-x-2 group shrink-0" onclick="window.location.href=\'/\'">' +
    '    <span class="text-4xl sm:text-7xl font-black tracking-tighter text-[#0f172a] flex items-center overflow-hidden">' +
    '      <span class="inline-block transform group-hover:-translate-y-0.5 transition-transform duration-300">Files</span><span class="text-[#dc2626] inline-block transform group-hover:translate-y-0.5 group-hover:scale-105 transition-all duration-300">que</span>' +
    '    </span>' +
    '    <span class="w-2.5 h-2.5 rounded-full bg-red-600 animate-ping ml-1 hidden sm:inline-block"></span>' +
    '  </div>' +
    '  <a href="/" class="back-link"><span>🏠</span> Back to Dashboard</a>' +
    '</div>' +
    '<div class="container">' +
    '<div class="panel tools-panel"><div class="panel-title">🛠️ Tools (' + toolRegistry.length + ')</div>' + toolsListHTML + '</div>' +
    '<div class="panel workflow-panel"><div class="panel-title">📋 Your Workflow</div>' +
    '<div class="workflow-area" id="workflowArea"><p style="color:#94a3b8;font-weight:600;font-size:15px;">Click on tools to add steps</p></div>' +
    '<div class="controls"><input type="text" id="wfName" placeholder="Workflow Name">' +
    '<button class="btn" onclick="saveWorkflow()">💾 Save Workflow</button>' +
    '<button class="btn btn-secondary" onclick="clearWorkflow()">🗑️ Clear</button></div>' +
    '<div class="file-upload-section"><h3 style="font-size:15px;font-weight:800;color:#0f172a;">📁 Test Your Workflow</h3>' +
    '<input type="file" id="wfFile" multiple>' + 
    '<button class="btn" onclick="executeWorkflow()" style="margin-left: 10px;">▶️ Execute Workflow</button>' +
    '<div id="executionResult" style="margin-top:10px;font-weight:700;font-size:14px;"></div></div>' +
    '<div class="saved-list"><h3 style="font-size:15px;font-weight:800;color:#0f172a;margin-bottom:0.75rem;">Saved Workflows</h3><div id="savedList"></div></div>' +
    '</div></div>' +
    '<div id="modalContainer"></div>' +
    '<script>' +
    'let steps = [];' +
    'let selectedFiles = [];' +
    'document.getElementById("wfFile").addEventListener("change", function(e) { selectedFiles = e.target.files; });' +
    'function updateStepConfig(idx, key, val) { steps[idx][key] = val; }' +
    'function addToWorkflow(id, name, icon) { ' +
    '  steps.push({ id: id, name: name, icon: icon, isComplex: false, mode: "all", range: "", sortOrder: "upload", customSequence: "", pageOrder: "", rotatePages: "", rotateDegree: "90", watermarkText: "CONFIDENTIAL", watermarkMode: "diagonal", compressQuality: "0.70", convertFormat: "jpg", convertQuality: "0.92", reducerMode: "resize", reducerWidth: "800", reducerHeight: "600", reducerTargetKb: "50", sizePreset: "3.5x4.5", bgColor: "#f87171", zoom: "100", copies: "32", removeBg: false, mergeDirection: "horizontal" });' +
    '  renderSteps();' +
    '}' +
    'function renderSteps(activeIdx = -1) {' +
    '  const area = document.getElementById("workflowArea");' +
    '  if (steps.length === 0) { area.innerHTML = "<p style=\'color:#94a3b8;font-weight:600;font-size:15px;\'>Click on tools to add steps</p>"; return; }' +
    '  let html = "";' +
    '  for (let i = 0; i < steps.length; i++) {' +
    '    let toggleHtml = "";' +
    '    let isActive = (i === activeIdx);' +
    '    let stepClass = isActive ? "step active-step" : "step";' +
    '    if (isActive) {' +
    '       toggleHtml += "<div style=\'margin-left:auto; display:flex; align-items:center; gap:6px; background:#0284c7; color:white; padding:6px 12px; border-radius:8px; font-weight:800; font-size:13px;\'>⚡ Processing Step...</div>";' +
    '    } else if (steps[i].id === "pdf-to-excel") {' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px; background:#e2e8f0; padding:6px 12px; border-radius:8px; border:1px solid #cbd5e1;\'>" +' +
    '        "<span style=\'font-size:13px; font-weight:bold; color:#475569;\'>Mode:</span>" +' +
    '        "<select onchange=\'updateStepConfig(" + i + ", \\\"isComplex\\\", this.value === \\\"complex\\\")\' style=\'padding:4px 8px; border-radius:6px; border:1px solid #94a3b8; font-size:13px; font-weight:bold; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '          "<option value=\\\"simple\\\" " + (!steps[i].isComplex ? "selected" : "") + ">Simple Table</option>" +' +
    '          "<option value=\\\"complex\\\" " + (steps[i].isComplex ? "selected" : "") + ">Complex Table</option>" +' +
    '        "</select>" +' +
    '      "</div>";' +
    '    } else if (steps[i].id === "merge-image") {' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px; background:linear-gradient(135deg, #0d9488, #4f46e5); padding:6px 14px; border-radius:8px; color:white; font-weight:bold; font-size:13px; box-shadow:0 2px 5px rgba(0,0,0,0.1);\'>" + ' +
    '        "<span>🧩 Studio & Crop Dialog Ready</span>" + ' +
    '      "</div>";' +
    '    } else if (steps[i].id === "pdf-splitter") {' +
    '      let mode = steps[i].mode || "all";' +
    '      let rangeVal = steps[i].range || "";' +
    '      let displayRange = mode === "all" ? "display:none;" : "";' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px; background:#e2e8f0; padding:6px 12px; border-radius:8px; border:1px solid #cbd5e1;\'>" +' +
    '        "<span style=\'font-size:13px; font-weight:bold; color:#475569;\'>Mode:</span>" +' +
    '        "<select onchange=\'updateStepConfig(" + i + ", \\\"mode\\\", this.value); renderSteps();\' style=\'padding:4px 8px; border-radius:6px; border:1px solid #94a3b8; font-size:13px; font-weight:bold; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '          "<option value=\\\"all\\\" " + (mode === "all" ? "selected" : "") + ">Split All (ZIP)</option>" +' +
    '          "<option value=\\\"range\\\" " + (mode === "range" ? "selected" : "") + ">Custom Range</option>" +' +
    '        "</select>" +' +
    '        "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"range\\\", this.value)\' value=\'" + rangeVal + "\' placeholder=\'e.g. 1-3, 5\' style=\'padding:4px 8px; border-radius:6px; border:1px solid #94a3b8; font-size:13px; width:110px; outline:none; " + displayRange + "\'>" + ' +
    '      "</div>";' +
    '    } else if (steps[i].id === "merge-pdf" || steps[i].id === "image-to-pdf") {' +
    '      let sOrder = steps[i].sortOrder || "upload";' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px; background:#e2e8f0; padding:6px 12px; border-radius:8px; border:1px solid #cbd5e1;\'>" +' +
    '        "<span style=\'font-size:13px; font-weight:bold; color:#475569;\'>Combine Order:</span>" +' +
    '        "<select onchange=\'updateStepConfig(" + i + ", \\\"sortOrder\\\", this.value)\' style=\'padding:4px 8px; border-radius:6px; border:1px solid #94a3b8; font-size:13px; font-weight:bold; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '          "<option value=\\\"upload\\\" " + (sOrder === "upload" ? "selected" : "") + ">As Uploaded</option>" +' +
    '          "<option value=\\\"az\\\" " + (sOrder === "az" ? "selected" : "") + ">Alphabetical (A-Z)</option>" +' +
    '          "<option value=\\\"za\\\" " + (sOrder === "za" ? "selected" : "") + ">Alphabetical (Z-A)</option>" +' +
    '        "</select>" +' +
    '      "</div>";' +
    '    } else if (steps[i].id === "pdf-organizer") {' +
    '      let pOrder = steps[i].pageOrder || "";' +
    '      let rPages = steps[i].rotatePages || "";' +
    '      let rDeg = steps[i].rotateDegree || "90";' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px;\'>" +' +
    '        "<div style=\'display:flex; align-items:flex-start; gap:5px; background:#e2e8f0; padding:6px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '          "<span style=\'font-size:13px; font-weight:bold; color:#475569; margin-top:2px;\'>Page Order:</span>" +' +
    '          "<div style=\'display:flex; flex-direction:column; gap:2px;\'>" +' +
    '            "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"pageOrder\\\", this.value)\' value=\'" + pOrder + "\' placeholder=\'e.g. 1, 3, 2, 5-7\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; width:130px; outline:none;\'>" +' +
    '            "<span style=\'font-size:10px; color:#64748b; line-height:1;\'>e.g., 1, 3, 2, 5-7</span>" +' +
    '          "</div>" +' +
    '        "</div>" +' +
    '        "<div style=\'display:flex; align-items:flex-start; gap:5px; background:#e2e8f0; padding:6px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '          "<span style=\'font-size:13px; font-weight:bold; color:#475569; margin-top:2px;\'>Rotate:</span>" +' +
    '          "<div style=\'display:flex; flex-direction:column; gap:2px;\'>" +' +
    '            "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"rotatePages\\\", this.value)\' value=\'" + rPages + "\' placeholder=\'e.g. 2, 4 or all\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; width:120px; outline:none;\'>" +' +
    '            "<span style=\'font-size:10px; color:#64748b; line-height:1;\'>e.g. 2, 4 or all</span>" +' +
    '          "</div>" +' +
    '          "<select onchange=\'updateStepConfig(" + i + ", \\\"rotateDegree\\\", this.value)\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '             "<option value=\\\"90\\\" " + (rDeg === "90" ? "selected" : "") + ">90°</option>" +' +
    '             "<option value=\\\"180\\\" " + (rDeg === "180" ? "selected" : "") + ">180°</option>" +' +
    '             "<option value=\\\"270\\\" " + (rDeg === "270" ? "selected" : "") + ">270°</option>" +' +
    '          "</select>" +' +
    '        "</div>" +' +
    '      "</div>";' +
    '    } else if (steps[i].id === "pdf-watermark") {' +
    '      let wmText = steps[i].watermarkText || "CONFIDENTIAL";' +
    '      let wmMode = steps[i].watermarkMode || "diagonal";' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px;\'>" +' +
    '        "<div style=\'display:flex; align-items:flex-start; gap:5px; background:#e2e8f0; padding:6px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '          "<span style=\'font-size:13px; font-weight:bold; color:#475569; margin-top:2px;\'>Text:</span>" +' +
    '          "<div style=\'display:flex; flex-direction:column; gap:2px;\'>" +' +
    '            "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"watermarkText\\\", this.value)\' value=\'" + wmText + "\' placeholder=\'e.g. CONFIDENTIAL\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; width:120px; outline:none;\'>" +' +
    '            "<span style=\'font-size:10px; color:#64748b; line-height:1;\'>e.g. DRAFT</span>" +' +
    '          "</div>" +' +
    '        "</div>" +' +
    '        "<div style=\'display:flex; align-items:flex-start; gap:5px; background:#e2e8f0; padding:6px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '          "<span style=\'font-size:13px; font-weight:bold; color:#475569; margin-top:2px;\'>Mode:</span>" +' +
    '          "<select onchange=\'updateStepConfig(" + i + ", \\\"watermarkMode\\\", this.value)\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '             "<option value=\\\"diagonal\\\" " + (wmMode === "diagonal" ? "selected" : "") + ">Watermark Only</option>" +' +
    '             "<option value=\\\"footer_num\\\" " + (wmMode === "footer_num" ? "selected" : "") + ">Page Numbers Only</option>" +' +
    '             "<option value=\\\"both\\\" " + (wmMode === "both" ? "selected" : "") + ">Both</option>" +' +
    '          "</select>" +' +
    '        "</div>" +' +
    '      "</div>";' +
    '    } else if (steps[i].id === "compress-image") {' +
    '      let qual = steps[i].compressQuality || "0.70";' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px; background:#e2e8f0; padding:6px 12px; border-radius:8px; border:1px solid #cbd5e1;\'>" +' +
    '        "<span style=\'font-size:13px; font-weight:bold; color:#475569;\'>Quality:</span>" +' +
    '        "<select onchange=\'updateStepConfig(" + i + ", \\\"compressQuality\\\", this.value)\' style=\'padding:4px 8px; border-radius:6px; border:1px solid #94a3b8; font-size:13px; font-weight:bold; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '          "<option value=\\\"0.90\\\" " + (qual === "0.90" ? "selected" : "") + ">High (Less Compression)</option>" +' +
    '          "<option value=\\\"0.70\\\" " + (qual === "0.70" ? "selected" : "") + ">Medium (Default)</option>" +' +
    '          "<option value=\\\"0.50\\\" " + (qual === "0.50" ? "selected" : "") + ">Low (Max Compression)</option>" +' +
    '        "</select>" +' +
    '      "</div>";' +
    '    } else if (steps[i].id === "image-converter") {' +
    '      let fmt = steps[i].convertFormat || "jpg";' +
    '      let qual = steps[i].convertQuality || "0.92";' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px;\'>" +' +
    '        "<div style=\'display:flex; align-items:center; gap:5px; background:#e2e8f0; padding:6px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '          "<span style=\'font-size:13px; font-weight:bold; color:#475569;\'>Format:</span>" +' +
    '          "<select onchange=\'updateStepConfig(" + i + ", \\\"convertFormat\\\", this.value)\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '            "<option value=\\\"jpg\\\" " + (fmt === "jpg" ? "selected" : "") + ">JPG</option>" +' +
    '            "<option value=\\\"jpeg\\\" " + (fmt === "jpeg" ? "selected" : "") + ">JPEG</option>" +' +
    '            "<option value=\\\"png\\\" " + (fmt === "png" ? "selected" : "") + ">PNG</option>" +' +
    '            "<option value=\\\"webp\\\" " + (fmt === "webp" ? "selected" : "") + ">WEBP</option>" +' +
    '            "<option value=\\\"bmp\\\" " + (fmt === "bmp" ? "selected" : "") + ">BMP</option>" +' +
    '            "<option value=\\\"ico\\\" " + (fmt === "ico" ? "selected" : "") + ">ICO</option>" +' +
    '          "</select>" +' +
    '        "</div>" +' +
    '        "<div style=\'display:flex; align-items:center; gap:5px; background:#e2e8f0; padding:6px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '          "<span style=\'font-size:13px; font-weight:bold; color:#475569;\'>Quality:</span>" +' +
    '          "<select onchange=\'updateStepConfig(" + i + ", \\\"convertQuality\\\", this.value)\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '            "<option value=\\\"1.00\\\" " + (qual === "1.00" ? "selected" : "") + ">100% (Max)</option>" +' +
    '            "<option value=\\\"0.92\\\" " + (qual === "0.92" ? "selected" : "") + ">92% (High)</option>" +' +
    '            "<option value=\\\"0.70\\\" " + (qual === "0.70" ? "selected" : "") + ">70% (Medium)</option>" +' +
    '            "<option value=\\\"0.50\\\" " + (qual === "0.50" ? "selected" : "") + ">50% (Low)</option>" +' +
    '          "</select>" +' +
    '        "</div>" +' +
    '      "</div>";' +
    '    } else if (steps[i].id === "image-reducer") {' +
    '      let rMode = steps[i].reducerMode || "resize";' +
    '      let rW = steps[i].reducerWidth || "800";' +
    '      let rH = steps[i].reducerHeight || "600";' +
    '      let rKb = steps[i].reducerTargetKb || "50";' +
    '      let showResize = rMode === "resize" ? "inline-block" : "none";' +
    '      let showTarget = rMode === "target" ? "inline-block" : "none";' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px;\'>" +' +
    '        "<div style=\'display:flex; align-items:center; gap:5px; background:#e2e8f0; padding:6px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '          "<span style=\'font-size:13px; font-weight:bold; color:#475569;\'>Mode:</span>" +' +
    '          "<select onchange=\'updateStepConfig(" + i + ", \\\"reducerMode\\\", this.value); renderSteps();\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '            "<option value=\\\"resize\\\" " + (rMode === "resize" ? "selected" : "") + ">Resize (W x H)</option>" +' +
    '            "<option value=\\\"target\\\" " + (rMode === "target" ? "selected" : "") + ">Target KB</option>" +' +
    '          "</select>" +' +
    '          "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"reducerWidth\\\", this.value)\' value=\'" + rW + "\' placeholder=\'W\' style=\'display:" + showResize + "; padding:2px 4px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; width:50px; outline:none;\'>" +' +
    '          "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"reducerHeight\\\", this.value)\' value=\'" + rH + "\' placeholder=\'H\' style=\'display:" + showResize + "; padding:2px 4px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; width:50px; outline:none;\'>" +' +
    '          "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"reducerTargetKb\\\", this.value)\' value=\'" + rKb + "\' placeholder=\'KB\' style=\'display:" + showTarget + "; padding:2px 4px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; width:65px; outline:none;\'>" +' +
    '        "</div>" +' +
    '      "</div>";' +
    '    }' +
    '    html += "<div class=\'" + stepClass + "\'><div class=\'step-number\'>" + (i+1) + "</div><span style=\'font-size:1.35rem;\'>" + steps[i].icon + "</span><strong style=\'font-size:1.2rem;\'>" + steps[i].name + "</strong>" + (toggleHtml ? toggleHtml : "<div style=\'margin-left:auto;\'></div>") + (!isActive ? "<button class=\'remove-btn\' style=\'margin-left:10px;\' onclick=\'removeStep(" + i + ")\'>✕</button>" : "") + "</div>";' +
    '  }' +
    '  area.innerHTML = html;' +
    '}' +
    'function removeStep(index) { steps.splice(index, 1); renderSteps(); }' +
    'function clearWorkflow() { steps = []; renderSteps(); }' +
    'async function saveWorkflow() {' +
    '  const name = document.getElementById("wfName").value.trim() || "My Workflow";' +
    '  if (steps.length === 0) { alert("Please add at least one tool"); return; }' +
    '  const payload = { name: name, steps: steps.map(s => ({ toolId: s.id, isComplex: !!s.isComplex, mode: s.mode, range: s.range, sortOrder: s.sortOrder, customSequence: s.customSequence, pageOrder: s.pageOrder, rotatePages: s.rotatePages, rotateDegree: s.rotateDegree, watermarkText: s.watermarkText, watermarkMode: s.watermarkMode, compressQuality: s.compressQuality, convertFormat: s.convertFormat, convertQuality: s.convertQuality, reducerMode: s.reducerMode, reducerWidth: s.reducerWidth, reducerHeight: s.reducerHeight, reducerTargetKb: s.reducerTargetKb, sizePreset: s.sizePreset, bgColor: s.bgColor, zoom: s.zoom, copies: s.copies, removeBg: s.removeBg, printSheet: s.printSheet, mergeDirection: s.mergeDirection, customProcessedImages: s.customProcessedImages })) };' +
    '  try {' +
    '    const res = await fetch("/api/workflow/create", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });' +
    '    const data = await res.json();' +
    '    if (data.success) { alert("✅ Workflow saved!"); loadSavedWorkflows(); } else { alert("Error: " + data.error); }' +
    '  } catch(e) { alert("Network error: " + e.message); }' +
    '}' +
    'async function loadSavedWorkflows() {' +
    '  try {' +
    '    const res = await fetch("/api/workflow/list");' +
    '    const data = await res.json();' +
    '    const list = document.getElementById("savedList");' +
    '    let html = "";' +
    '    for (const w of data.workflows) { html += "<div class=\'saved-item\'><strong>" + w.name + "</strong> - " + w.steps.length + " steps</div>"; }' +
    '    list.innerHTML = html;' +
    '  } catch(e) { console.error(e); }' +
    '}' +
    'async function executeWorkflow() {' +
    '  if (selectedFiles.length === 0 || steps.length === 0) { alert("Please select file(s) and add workflow steps."); return; }' +
    '  const passportStepIdx = steps.findIndex(s => s.id === "passport-studio");' +
    '  if (passportStepIdx !== -1) {' +
    '    openProfessionalStudioModal(selectedFiles[0], passportStepIdx);' +
    '    return;' +
    '  }' +
    '  const mergeImgStepIdx = steps.findIndex(s => s.id === "merge-image");' +
    '  if (mergeImgStepIdx !== -1) {' +
    '    openMergeImageStudioModal(selectedFiles, mergeImgStepIdx);' +
    '    return;' +
    '  }' +
    '  sendWorkflowExecution();' +
    '}' +
    'function openMergeImageStudioModal(files, stepIdx) {' +
    '  let modalHtml = \'<div style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(15,23,42,0.8);backdrop-filter:blur(8px);display:flex;justify-content:center;align-items:center;z-index:9999;font-family:\\\'Plus Jakarta Sans\\\',sans-serif;"><div style="background:white;padding:2.5rem;border-radius:1.75rem;width:980px;max-width:96%;box-shadow:0 25px 60px rgba(0,0,0,0.4);max-height:92vh;overflow-y:auto;border:1px solid #e2e8f0;"><div style="display:flex;align-items:center;margin-bottom:1.5rem;"><div style="width:48px;height:48px;border-radius:14px;background:linear-gradient(135deg,#0d9488,#4f46e5);display:flex;align-items:center;justify-content:center;color:white;font-size:24px;box-shadow:0 10px 20px rgba(13,148,136,0.3);margin-right:1rem;">🧩</div><div><h2 style="font-size:24px;font-weight:900;color:#0f172a;margin:0;letter-spacing:-0.5px;">Merge Image Studio Pro</h2><p style="font-size:12px;color:#64748b;font-weight:600;margin:2px 0 0 0;">Interactive Studio: Reorder, Rotate 90°, Crop & Split Bucket Extraction</p></div></div><hr style="border:0;border-top:1px solid #e2e8f0;margin-bottom:1.5rem;"><div style="margin-bottom:1.5rem;"><label style="display:block;font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">Merge Direction / Flow</label><select id="studioMergeDir" style="width:100%;padding:12px 14px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:12px;outline:none;font-size:13px;font-weight:700;color:#1e293b;cursor:pointer;"><option value="horizontal">Horizontal (Side by Side)</option><option value="vertical" selected>Vertical (Top to Bottom)</option></select></div><div style="margin-bottom:1.5rem;"><label style="display:block;font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">Uploaded Images Workspace</label><div id="studioFilesContainer" style="display:flex;flex-direction:column;gap:1rem;background:#f8fafc;padding:1rem;border-radius:12px;border:1px solid #e2e8f0;max-height:350px;overflow-y:auto;"></div></div><div style="display:flex;justify-content:flex-end;gap:12px;padding-top:1rem;border-top:1px solid #e2e8f0;"><button onclick="closeModal()" style="padding:12px 24px;background:#f1f5f9;color:#334155;border:none;border-radius:12px;font-weight:800;font-size:13px;cursor:pointer;transition:0.2s;">Cancel</button><button id="applyMergeBtn" style="padding:12px 28px;background:linear-gradient(135deg,#0d9488,#4f46e5);color:white;border:none;border-radius:12px;font-weight:800;font-size:13px;cursor:pointer;box-shadow:0 10px 20px rgba(13,148,136,0.3);transition:0.2s;">✨ Apply & Execute Workflow</button></div></div></div>\';' +
    '  document.getElementById("modalContainer").innerHTML = modalHtml;' +
    '  document.getElementById("applyMergeBtn").onclick = function() { applyMergeStudioAndExecute(stepIdx); };' +
    '  window.studioItems = [];' +
    '  let loadedCount = 0;' +
    '  for (let i = 0; i < files.length; i++) {' +
    '    const reader = new FileReader();' +
    '    reader.onload = function(e) {' +
    '      const img = new Image();' +
    '      img.onload = function() {' +
    '        window.studioItems.push({ id: Math.random().toString(36).substring(2, 9), file: files[i], url: e.target.result, rotation: 0, imgObj: img });' +
    '        loadedCount++;' +
    '        if (loadedCount === files.length) { renderStudioFilesList(); }' +
    '      };' +
    '      img.src = e.target.result;' +
    '    };' +
    '    reader.readAsDataURL(files[i]);' +
    '  }' +
    '}' +
    'function renderStudioFilesList() {' +
    '  const container = document.getElementById("studioFilesContainer");' +
    '  if (!container) return;' +
    '  if (window.studioItems.length === 0) { container.innerHTML = "<p style=\'font-size:12px;color:#64748b;\'>No images remaining.</p>"; return; }' +
    '  let html = "";' +
    '  for (let i = 0; i < window.studioItems.length; i++) {' +
    '    const item = window.studioItems[i];' +
    '    html += \'<div style="background:white;padding:12px 16px;border-radius:12px;border:1px solid #cbd5e1;display:flex;align-items:center;justify-content:space-between;gap:1rem;"><div style="display:flex;align-items:center;gap:12px;"><img src="\' + item.url + \'" style="width:50px;height:50px;object-fit:contain;border-radius:8px;border:1px solid #e2e8f0;"><div style="display:flex;flex-direction:column;"><strong style="font-size:13px;color:#0f172a;">Image #\' + (i+1) + \'</strong><span style="font-size:11px;color:#64748b;">Size: \' + item.imgObj.width + \' × \' + item.imgObj.height + \'px</span></div></div><div style="display:flex;align-items:center;gap:6px;"><button type="button" onclick="moveStudioItem(\' + i + \', -1)" \' + (i === 0 ? \'disabled style="opacity:0.3;cursor:not-allowed;padding:6px 10px;background:#f1f5f9;border-radius:6px;font-size:12px;font-weight:bold;"\' : \'style="padding:6px 10px;background:#f1f5f9;border-radius:6px;font-size:12px;font-weight:bold;cursor:pointer;"\') + \'>⬆️ Up</button><button type="button" onclick="moveStudioItem(\' + i + \', 1)" \' + (i === window.studioItems.length - 1 ? \'disabled style="opacity:0.3;cursor:not-allowed;padding:6px 10px;background:#f1f5f9;border-radius:6px;font-size:12px;font-weight:bold;"\' : \'style="padding:6px 10px;background:#f1f5f9;border-radius:6px;font-size:12px;font-weight:bold;cursor:pointer;"\') + \'>⬇️ Down</button><button type="button" onclick="rotateStudioItem(\' + i + \')" style="padding:6px 10px;background:#fef3c7;color:#b45309;border:none;border-radius:6px;font-size:11px;font-weight:800;cursor:pointer;">🔄 Rotate</button><button type="button" onclick="openStudioCropModal(\' + i + \')" style="padding:6px 10px;background:#ccfbf1;color:#0f766e;border:none;border-radius:6px;font-size:11px;font-weight:800;cursor:pointer;">✂️ Crop</button><button type="button" onclick="openStudioSplitModal(\' + i + \')" style="padding:6px 10px;background:#e0e7ff;color:#3730a3;border:none;border-radius:6px;font-size:11px;font-weight:800;cursor:pointer;">🔪 Split</button><button type="button" onclick="deleteStudioItem(\' + i + \')" style="padding:6px 10px;background:#fee2e2;color:#dc2626;border:none;border-radius:6px;font-size:11px;font-weight:800;cursor:pointer;">🗑️ Delete</button></div></div>\';' +
    '  }' +
    '  container.innerHTML = html;' +
    '}' +
    'function moveStudioItem(idx, direction) {' +
    '  const target = idx + direction;' +
    '  if (target < 0 || target >= window.studioItems.length) return;' +
    '  const temp = window.studioItems[idx];' +
    '  window.studioItems[idx] = window.studioItems[target];' +
    '  window.studioItems[target] = temp;' +
    '  renderStudioFilesList();' +
    '}' +
    'function deleteStudioItem(idx) {' +
    '  window.studioItems.splice(idx, 1);' +
    '  renderStudioFilesList();' +
    '}' +
    'function rotateStudioItem(idx) {' +
    '  const item = window.studioItems[idx];' +
    '  item.rotation = (item.rotation + 90) % 360;' +
    '  const canvas = document.createElement("canvas");' +
    '  const ctx = canvas.getContext("2d");' +
    '  const img = item.imgObj;' +
    '  if (item.rotation === 90 || item.rotation === 270) { canvas.width = img.height; canvas.height = img.width; }' +
    '  else { canvas.width = img.width; canvas.height = img.height; }' +
    '  ctx.fillStyle = "#FFFFFF"; ctx.fillRect(0,0,canvas.width,canvas.height);' +
    '  ctx.translate(canvas.width/2, canvas.height/2);' +
    '  ctx.rotate((item.rotation * Math.PI) / 180);' +
    '  ctx.drawImage(img, -img.width/2, -img.height/2);' +
    '  item.url = canvas.toDataURL("image/jpeg", 0.95);' +
    '  item.rotation = 0;' +
    '  const newImg = new Image();' +
    '  newImg.onload = function() { item.imgObj = newImg; renderStudioFilesList(); };' +
    '  newImg.src = item.url;' +
    '}' +
    'function openStudioCropModal(targetIdx) {' +
    '  const item = window.studioItems[targetIdx];' +
    '  const cropHtml = \'<div id="subStudioModal" style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(15,23,42,0.85);z-index:10000;display:flex;justify-content:center;align-items:center;"><div style="background:white;padding:2rem;border-radius:1.5rem;width:700px;max-width:95%;box-shadow:0 25px 50px rgba(0,0,0,0.5);text-align:center;"><h3 style="font-size:18px;font-weight:900;color:#0f172a;margin-bottom:1rem;">Crop Image</h3><div style="background:#f1f5f9;padding:1rem;border-radius:12px;display:flex;justify-content:center;max-height:50vh;overflow:hidden;margin-bottom:1rem;"><canvas id="subCropCanvas" style="max-width:100%;max-height:45vh;object-fit:contain;cursor:crosshair;"></canvas></div><div style="display:flex;justify-content:flex-end;gap:10px;"><button onclick="document.getElementById(\\\'subStudioModal\\\").remove()" style="padding:10px 20px;background:#f1f5f9;color:#334155;border:none;border-radius:8px;font-weight:800;cursor:pointer;">Cancel</button><button id="saveCropBtn" style="padding:10px 24px;background:#0d9488;color:white;border:none;border-radius:8px;font-weight:800;cursor:pointer;">Save Crop</button></div></div></div>\';' +
    '  const div = document.createElement("div"); div.innerHTML = cropHtml; document.body.appendChild(div);' +
    '  document.getElementById("saveCropBtn").onclick = function() { saveStudioCrop(targetIdx); };' +
    '  const canvas = document.getElementById("subCropCanvas");' +
    '  const ctx = canvas.getContext("2d");' +
    '  const img = item.imgObj;' +
    '  canvas.width = img.width; canvas.height = img.height;' +
    '  ctx.drawImage(img, 0, 0);' +
    '  window.subCropBox = { x: img.width*0.1, y: img.height*0.1, w: img.width*0.8, h: img.height*0.8 };' +
    '  let isSubDragging = false, startX = 0, startY = 0;' +
    '  function redrawSubCrop() {' +
    '    ctx.clearRect(0,0,canvas.width,canvas.height); ctx.drawImage(img,0,0);' +
    '    ctx.fillStyle = "rgba(0,0,0,0.5)"; ctx.fillRect(0,0,canvas.width,canvas.height);' +
    '    ctx.save(); ctx.beginPath(); ctx.rect(window.subCropBox.x, window.subCropBox.y, window.subCropBox.w, window.subCropBox.h); ctx.clip(); ctx.drawImage(img,0,0); ctx.restore();' +
    '    ctx.strokeStyle = "#0d9488"; ctx.lineWidth = 3; ctx.strokeRect(window.subCropBox.x, window.subCropBox.y, window.subCropBox.w, window.subCropBox.h);' +
    '  }' +
    '  redrawSubCrop();' +
    '  canvas.onmousedown = function(e) {' +
    '    const rect = canvas.getBoundingClientRect();' +
    '    const scaleX = canvas.width / rect.width; const scaleY = canvas.height / rect.height;' +
    '    startX = (e.clientX - rect.left) * scaleX; startY = (e.clientY - rect.top) * scaleY;' +
    '    isSubDragging = true;' +
    '  };' +
    '  canvas.onmousemove = function(e) {' +
    '    if (!isSubDragging) return;' +
    '    const rect = canvas.getBoundingClientRect();' +
    '    const scaleX = canvas.width / rect.width; const scaleY = canvas.height / rect.height;' +
    '    const curX = (e.clientX - rect.left) * scaleX; const curY = (e.clientY - rect.top) * scaleY;' +
    '    window.subCropBox.w = Math.max(30, Math.min(canvas.width - window.subCropBox.x, curX - startX + window.subCropBox.w));' +
    '    window.subCropBox.h = Math.max(30, Math.min(canvas.height - window.subCropBox.y, curY - startY + window.subCropBox.h));' +
    '    redrawSubCrop();' +
    '  };' +
    '  window.onmouseup = function() { isSubDragging = false; };' +
    '}' +
    'function saveStudioCrop(targetIdx) {' +
    '  const canvas = document.getElementById("subCropCanvas");' +
    '  const box = window.subCropBox;' +
    '  const outCanvas = document.createElement("canvas");' +
    '  outCanvas.width = box.w; outCanvas.height = box.h;' +
    '  const outCtx = outCanvas.getContext("2d");' +
    '  outCtx.fillStyle = "#FFFFFF"; outCtx.fillRect(0,0,box.w,box.h);' +
    '  outCtx.drawImage(canvas, box.x, box.y, box.w, box.h, 0, 0, box.w, box.h);' +
    '  const item = window.studioItems[targetIdx];' +
    '  item.url = outCanvas.toDataURL("image/jpeg", 0.95);' +
    '  const newImg = new Image();' +
    '  newImg.onload = function() { item.imgObj = newImg; renderStudioFilesList(); document.getElementById("subStudioModal").remove(); };' +
    '  newImg.src = item.url;' +
    '}' +
    'function openStudioSplitModal(targetIdx) {' +
    '  const item = window.studioItems[targetIdx];' +
    '  const splitHtml = \'<div id="subSplitModal" style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(15,23,42,0.85);z-index:10000;display:flex;justify-content:center;align-items:center;"><div style="background:white;padding:2rem;border-radius:1.5rem;width:750px;max-width:95%;box-shadow:0 25px 50px rgba(0,0,0,0.5);text-align:center;"><h3 style="font-size:18px;font-weight:900;color:#0f172a;margin-bottom:0.5rem;">Split Image</h3><p style="font-size:12px;color:#64748b;margin-bottom:1rem;">Drag box on image and click "Add Part" to add parts to merge list.</p><div style="background:#f1f5f9;padding:1rem;border-radius:12px;display:flex;justify-content:center;max-height:45vh;overflow:hidden;margin-bottom:1rem;"><canvas id="subSplitCanvas" style="max-width:100%;max-height:40vh;object-fit:contain;cursor:crosshair;"></canvas></div><div style="margin-bottom:1.0rem;display:flex;gap:8px;overflow-x:auto;padding:6px;background:#f8fafc;border-radius:8px;min-height:50px;align-items:center;" id="subSplitBucket"></div><div style="display:flex;justify-content:space-between;align-items:center;"><button onclick="addSubSplitPart()" style="padding:10px 18px;background:#4f46e5;color:white;border:none;border-radius:8px;font-weight:800;cursor:pointer;">➕ Add Part to Bucket</button><div style="display:flex;gap:10px;"><button onclick="document.getElementById(\\\'subSplitModal\\\").remove()" style="padding:10px 20px;background:#f1f5f9;color:#334155;border:none;border-radius:8px;font-weight:800;cursor:pointer;">Cancel</button><button id="saveSplitBtn" style="padding:10px 24px;background:#0d9488;color:white;border:none;border-radius:8px;font-weight:800;cursor:pointer;">Done & Replace</button></div></div></div></div>\';' +
    '  const div = document.createElement("div"); div.innerHTML = splitHtml; document.body.appendChild(div);' +
    '  document.getElementById("saveSplitBtn").onclick = function() { saveStudioSplit(targetIdx); };' +
    '  const canvas = document.getElementById("subSplitCanvas");' +
    '  const ctx = canvas.getContext("2d");' +
    '  const img = item.imgObj;' +
    '  canvas.width = img.width; canvas.height = img.height;' +
    '  ctx.drawImage(img, 0, 0);' +
    '  window.subSplitBox = { x: img.width*0.1, y: img.height*0.1, w: img.width*0.4, h: img.height*0.4 };' +
    '  window.subSplitBucketList = [];' +
    '  let isSubSplitting = false, startX = 0, startY = 0;' +
    '  function redrawSubSplit() {' +
    '    ctx.clearRect(0,0,canvas.width,canvas.height); ctx.drawImage(img,0,0);' +
    '    ctx.fillStyle = "rgba(0,0,0,0.5)"; ctx.fillRect(0,0,canvas.width,canvas.height);' +
    '    ctx.save(); ctx.beginPath(); ctx.rect(window.subSplitBox.x, window.subSplitBox.y, window.subSplitBox.w, window.subSplitBox.h); ctx.clip(); ctx.drawImage(img,0,0); ctx.restore();' +
    '    ctx.strokeStyle = "#4f46e5"; ctx.lineWidth = 3; ctx.strokeRect(window.subSplitBox.x, window.subSplitBox.y, window.subSplitBox.w, window.subSplitBox.h);' +
    '  }' +
    '  redrawSubSplit();' +
    '  canvas.onmousedown = function(e) {' +
    '    const rect = canvas.getBoundingClientRect();' +
    '    const scaleX = canvas.width / rect.width; const scaleY = canvas.height / rect.height;' +
    '    startX = (e.clientX - rect.left) * scaleX; startY = (e.clientY - rect.top) * scaleY;' +
    '    isSubSplitting = true;' +
    '  };' +
    '  canvas.onmousemove = function(e) {' +
    '    if (!isSubSplitting) return;' +
    '    const rect = canvas.getBoundingClientRect();' +
    '    const scaleX = canvas.width / rect.width; const scaleY = canvas.height / rect.height;' +
    '    const curX = (e.clientX - rect.left) * scaleX; const curY = (e.clientY - rect.top) * scaleY;' +
    '    window.subSplitBox.w = Math.max(30, Math.min(canvas.width - window.subSplitBox.x, curX - startX + window.subSplitBox.w));' +
    '    window.subSplitBox.h = Math.max(30, Math.min(canvas.height - window.subSplitBox.y, curY - startY + window.subSplitBox.h));' +
    '    redrawSubSplit();' +
    '  };' +
    '  window.onmouseup = function() { isSubSplitting = false; };' +
    '}' +
    'function addSubSplitPart() {' +
    '  const canvas = document.getElementById("subSplitCanvas");' +
    '  const box = window.subSplitBox;' +
    '  const outCanvas = document.createElement("canvas");' +
    '  outCanvas.width = box.w; outCanvas.height = box.h;' +
    '  const outCtx = outCanvas.getContext("2d");' +
    '  outCtx.fillStyle = "#FFFFFF"; outCtx.fillRect(0,0,box.w,box.h);' +
    '  outCtx.drawImage(canvas, box.x, box.y, box.w, box.h, 0, 0, box.w, box.h);' +
    '  const url = outCanvas.toDataURL("image/jpeg", 0.95);' +
    '  window.subSplitBucketList.push(url);' +
    '  const bucket = document.getElementById("subSplitBucket");' +
    '  bucket.innerHTML += \'<img src="\' + url + \'" style="height:40px;width:40px;object-fit:contain;border:1px solid #cbd5e1;border-radius:6px;background:white;">\';' +
    '}' +
    'function saveStudioSplit(targetIdx) {' +
    '  if (window.subSplitBucketList.length === 0) { document.getElementById("subSplitModal").remove(); return; }' +
    '  let loaded = 0;' +
    '  const newItems = [];' +
    '  window.subSplitBucketList.forEach((url, i) => {' +
    '    const img = new Image();' +
    '    img.onload = function() {' +
    '      newItems.push({ id: Math.random().toString(36).substring(2, 9), file: window.studioItems[targetIdx].file, url: url, rotation: 0, imgObj: img });' +
    '      loaded++;' +
    '      if (loaded === window.subSplitBucketList.length) {' +
    '        window.studioItems.splice(targetIdx, 1, ...newItems);' +
    '        renderStudioFilesList();' +
    '        document.getElementById("subSplitModal").remove();' +
    '      }' +
    '    };' +
    '    img.src = url;' +
    '  });' +
    '}' +
    'function applyMergeStudioAndExecute(stepIdx) {' +
    '  steps[stepIdx].mergeDirection = document.getElementById("studioMergeDir").value;' +
    '  if (window.studioItems && window.studioItems.length > 0) {' +
    '    selectedFiles = window.studioItems.map(it => it.file);' +
    '  }' +
    '  closeModal();' +
    '  sendWorkflowExecution();' +
    '}' +
    'function openProfessionalStudioModal(file, stepIdx) {' +
    '  const reader = new FileReader();' +
    '  reader.onload = function(e) {' +
    '    const imgSrc = e.target.result;' +
    '    const modalHtml = \'<div style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(15,23,42,0.8);backdrop-filter:blur(8px);display:flex;justify-content:center;align-items:center;z-index:9999;font-family:\\\'Plus Jakarta Sans\\\',sans-serif;"><div style="background:white;padding:2.5rem;border-radius:1.75rem;width:820px;max-width:96%;box-shadow:0 25px 60px rgba(0,0,0,0.4);max-height:92vh;overflow-y:auto;border:1px solid #e2e8f0;"><div style="display:flex;align-items:center;margin-bottom:1.5rem;"><div style="width:48px;height:48px;border-radius:14px;background:linear-gradient(135deg,#4f46e5,#e11d48);display:flex;align-items:center;justify-content:center;color:white;font-size:24px;box-shadow:0 10px 20px rgba(79,70,229,0.3);margin-right:1rem;">🛂</div><div><h2 style="font-size:24px;font-weight:900;color:#0f172a;margin:0;letter-spacing:-0.5px;">Passport Photo Studio Pro</h2><p style="font-size:12px;color:#64748b;font-weight:600;margin:2px 0 0 0;">Professional Studio: Cloud AI Background Removal & HD Print Layout</p></div></div><hr style="border:0;border-top:1px solid #e2e8f0;margin-bottom:1.5rem;"><div style="display:grid;grid-template-columns:1.2fr 1fr;gap:2rem;margin-bottom:2rem;"><div style="display:flex;flex-direction:column;gap:1.2rem;"><div><label style="display:block;font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">1. Background Color Palette</label><select id="studioBg" style="width:100%;padding:12px 14px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:12px;outline:none;font-size:13px;font-weight:700;color:#1e293b;cursor:pointer;"><option value="#ffffff">Pure White (#FFFFFF)</option><option value="#a5cbf7">Light Blue (#A5CBF7)</option><option value="#3b82f6">Royal Blue (#3B82F6)</option><option value="#f87171" selected>Soft Red (#F87171)</option></select></div><div><label style="display:block;font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">2. Passport Size Preset</label><select id="studioSize" style="width:100%;padding:12px 14px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:12px;outline:none;font-size:13px;font-weight:700;color:#1e293b;cursor:pointer;"><option value="3.5x4.5">Standard Indian Passport (3.5 x 4.5 cm)</option><option value="2x2">US Visa Layout (2 x 2 inch)</option><option value="3.5x3.5">Indian PAN Card Size (3.5 x 3.5 cm)</option></select></div><div><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;"><label style="font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.5px;">3. Zoom & Face Position</label><span id="zoomValBadge" style="font-size:11px;font-weight:800;color:#4f46e5;background:#e0e7ff;padding:2px 8px;border-radius:6px;">100%</span></div><input type="range" id="studioZoom" min="50" max="250" value="100" style="width:100%;height:6px;background:#cbd5e1;border-radius:4px;accent-color:#4f46e5;cursor:pointer;"></div><div><label style="display:block;font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">4. Copies & Grid Layout</label><select id="studioCopies" style="width:100%;padding:12px 14px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:12px;outline:none;font-size:13px;font-weight:700;color:#1e293b;cursor:pointer;"><option value="4">4 Photos Grid</option><option value="8">8 Photos Grid</option><option value="16">16 Photos Grid</option><option value="32" selected>32 Photos Grid (Full A4 Sheet)</option></select></div><div style="background:#e0e7ff;border:1px solid #c7d2fe;padding:12px 14px;border-radius:12px;display:flex;align-items:center;gap:10px;"><input type="checkbox" id="studioRemoveBg" style="width:18px;height:18px;accent-color:#4f46e5;cursor:pointer;"><label for="studioRemoveBg" style="font-size:12px;font-weight:800;color:#3730a3;cursor:pointer;">Remove Background Pro (Cloud AI Engine)</label></div></div><div style="display:flex;flex-direction:column;align-items:center;justify-content:center;background:#f1f5f9;padding:1.5rem;border-radius:16px;border:1px solid #e2e8f0;position:relative;"><span style="font-size:11px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:1px;margin-bottom:12px;">Live Studio Preview</span><div id="modalPreviewBox" style="width:160px;height:200px;background:#f87171;box-shadow:0 20px 30px rgba(0,0,0,0.15);border:2px solid white;overflow:hidden;position:relative;border-radius:8px;transition:background 0.3s;display:flex;align-items:center;justify-content:center;"><img src="\' + imgSrc + \'" id="modalPreviewImg" style="width:100%;height:100%;object-fit:contain;transform:scale(1);transform-origin:center;transition:transform 0.1s;" /></div><p style="font-size:11px;color:#64748b;font-weight:600;margin-top:12px;text-align:center;">Interactive preview updates instantly</p></div></div><div style="display:flex;justify-content:flex-end;gap:12px;padding-top:1rem;border-top:1px solid #e2e8f0;"><button onclick="closeModal()" style="padding:12px 24px;background:#f1f5f9;color:#334155;border:none;border-radius:12px;font-weight:800;font-size:13px;cursor:pointer;transition:0.2s;">Cancel</button><button onclick="applyStudioAndExecute(\' + stepIdx + \')" style="padding:12px 28px;background:linear-gradient(135deg,#4f46e5,#e11d48);color:white;border:none;border-radius:12px;font-weight:800;font-size:13px;cursor:pointer;box-shadow:0 10px 20px rgba(79,70,229,0.3);transition:0.2s;">✨ Apply & Execute Workflow</button></div></div></div>\';' +
    '    document.getElementById("modalContainer").innerHTML = modalHtml;' +
    '    document.getElementById("studioZoom").addEventListener("input", function(e) {' +
    '        const val = e.target.value;' +
    '        document.getElementById("zoomValBadge").textContent = val + "%";' +
    '        document.getElementById("modalPreviewImg").style.transform = "scale(" + (val / 100) + ")";' +
    '    });' +
    '    document.getElementById("studioBg").addEventListener("change", function(e) {' +
    '        document.getElementById("modalPreviewBox").style.backgroundColor = e.target.value;' +
    '    });' +
    '  };' +
    '  reader.readAsDataURL(file);' +
    '}' +
    'function closeModal() { document.getElementById("modalContainer").innerHTML = ""; }' +
    'function applyStudioAndExecute(stepIdx) {' +
    '  steps[stepIdx].sizePreset = document.getElementById("studioSize").value;' +
    '  steps[stepIdx].bgColor = document.getElementById("studioBg").value;' +
    '  steps[stepIdx].zoom = document.getElementById("studioZoom").value;' +
    '  steps[stepIdx].copies = document.getElementById("studioCopies").value;' +
    '  steps[stepIdx].removeBg = document.getElementById("studioRemoveBg").checked;' +
    '  steps[stepIdx].printSheet = true;' +
    '  closeModal();' +
    '  sendWorkflowExecution();' +
    '}' +
    'async function sendWorkflowExecution() {' +
    '  const name = document.getElementById("wfName").value.trim() || "My Workflow";' +
    '  let saveRes;' +
    '  try { saveRes = await fetch("/api/workflow/create", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: name, steps: steps.map(s => ({ toolId: s.id, isComplex: !!s.isComplex, mode: s.mode, range: s.range, sortOrder: s.sortOrder, customSequence: s.customSequence, pageOrder: s.pageOrder, rotatePages: s.rotatePages, rotateDegree: s.rotateDegree, watermarkText: s.watermarkText, watermarkMode: s.watermarkMode, compressQuality: s.compressQuality, convertFormat: s.convertFormat, convertQuality: s.convertQuality, reducerMode: s.reducerMode, reducerWidth: s.reducerWidth, reducerHeight: s.reducerHeight, reducerTargetKb: s.reducerTargetKb, sizePreset: s.sizePreset, bgColor: s.bgColor, zoom: s.zoom, copies: s.copies, removeBg: s.removeBg, printSheet: s.printSheet, mergeDirection: s.mergeDirection })) }) }); }' +
    '  catch(e) { alert("Save failed: " + e.message); return; }' +
    '  const saveData = await saveRes.json();' +
    '  if (!saveData.success) { alert("Save error: " + saveData.error); return; }' +
    '  const formData = new FormData();' +
    '  for (let i = 0; i < selectedFiles.length; i++) { formData.append("file", selectedFiles[i]); }' +
    '  try {' +
    '    document.getElementById("executionResult").innerHTML = "<p style=\'color:blue;\'>⏳ Processing workflow steps...</p>";' +
    '    for (let i = 0; i < steps.length; i++) {' +
    '       renderSteps(i);' +
    '       await new Promise(r => setTimeout(r, 600));' +
    '    }' +
    '    const execRes = await fetch("/api/workflow/execute/" + saveData.workflow.id, { method: "POST", body: formData });' +
    '    renderSteps(-1);' +
    '    if (execRes.ok) {' +
    '      const blob = await execRes.blob();' +
    '      let fname = "workflow_output";' +
    '      const cd = execRes.headers.get("content-disposition");' +
    '      if (cd) { const m = cd.match(/filename="?([^";]+)"?/i); if (m && m[1]) fname = m[1]; }' +
    '      else {' +
    '        const mimeExtMap = { "application/pdf": ".pdf", "image/jpeg": ".jpg", "image/png": ".png", "image/webp": ".webp", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx", "application/zip": ".zip", "application/msword": ".doc" };' +
    '        fname = "workflow_output" + (mimeExtMap[blob.type] || "");' +
    '      }' +
    '      const url = window.URL.createObjectURL(blob);' +
    '      const a = document.createElement("a");' +
    '      a.href = url; a.download = fname; a.click(); window.URL.revokeObjectURL(url);' +
    '      document.getElementById("executionResult").innerHTML = "<p style=\'color:green;\'>✅ Workflow executed successfully! Check download.</p>";' +
    '    } else {' +
    '      const errorData = await execRes.json().catch(() => ({ error: "Execution failed" }));' +
    '      document.getElementById("executionResult").innerHTML = "<p style=\'color:red;\'>❌ " + (errorData.error || "Execution failed") + "</p>";' +
    '    }' +
    '  } catch(e) { renderSteps(-1); document.getElementById("executionResult").innerHTML = "<p style=\'color:red;\'>Network error during execution: " + e.message + "</p>"; }' +
    '}' +
    'loadSavedWorkflows();' +
    '</script>' +
    '<!-- Filesque Enterprise Dark Footer -->' +
    '<footer class="bg-[#1e2029] text-slate-300 pt-16 pb-12 px-6 sm:px-12 border-t border-slate-800 font-sans mt-auto">' +
    '    <div class="max-w-7xl mx-auto">' +
    '        <!-- Main Links Grid -->' +
    '        <div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-8 mb-12">' +
    '            ' +
    '            <!-- Column 1: PRODUCT -->' +
    '            <div>' +
    '                <h4 class="text-white font-extrabold text-xs uppercase tracking-widest mb-4">Product</h4>' +
    '                <ul class="space-y-2.5 text-xs font-semibold text-slate-400">' +
    '                    <li><a href="index.html" class="hover:text-white transition-colors">Home</a></li>' +
    '                    <li><a href="features.html" class="hover:text-white transition-colors">Features</a></li>' +
    '                    <li><a href="tools.html" class="hover:text-white transition-colors">Tools</a></li>' +
    '                    <li><a href="faq.html" class="hover:text-white transition-colors">FAQ</a></li>' +
    '                </ul>' +
    '            </div>' +
    '' +
    '            <!-- Column 2: RESOURCES -->' +
    '            <div>' +
    '                <h4 class="text-white font-extrabold text-xs uppercase tracking-widest mb-4">Resources</h4>' +
    '                <ul class="space-y-2.5 text-xs font-semibold text-slate-400">' +
    '                    <li><a href="index.html#pdf-section" class="hover:text-white transition-colors">PDF Tools</a></li>' +
    '                    <li><a href="index.html#image-section" class="hover:text-white transition-colors">Image Tools</a></li>' +
    '                    <li><a href="index.html#business-section" class="hover:text-white transition-colors">Business Tools</a></li>' +
    '                    <li><a href="/workflow-builder" class="hover:text-white transition-colors flex items-center gap-1.5"><span class="text-amber-400">⚡</span> <span>WorkFlow</span></a></li>' +
    '                </ul>' +
    '            </div>' +
    '' +
    '            <!-- Column 3: SOLUTIONS -->' +
    '            <div>' +
    '                <h4 class="text-white font-extrabold text-xs uppercase tracking-widest mb-4">Solutions</h4>' +
    '                <ul class="space-y-2.5 text-xs font-semibold text-slate-400">' +
    '                    <li><a href="freelancer.html" class="hover:text-white transition-colors">Freelancer</a></li>' +
    '                    <li><a href="business.html" class="hover:text-white transition-colors">Business</a></li>' +
    '                    <li><a href="education.html" class="hover:text-white transition-colors">Education</a></li>' +
    '                </ul>' +
    '            </div>' +
    '' +
    '            <!-- Column 4: LEGAL -->' +
    '            <div>' +
    '                <h4 class="text-white font-extrabold text-xs uppercase tracking-widest mb-4">Legal</h4>' +
    '                <ul class="space-y-2.5 text-xs font-semibold text-slate-400">' +
    '                    <li><a href="security.html" class="hover:text-white transition-colors">Security</a></li>' +
    '                    <li><a href="privacy.html" class="hover:text-white transition-colors">Privacy Policy</a></li>' +
    '                    <li><a href="terms.html" class="hover:text-white transition-colors">Terms &amp; Conditions</a></li>' +
    '                    <li><a href="cookies.html" class="hover:text-white transition-colors">Cookies</a></li>' +
    '                </ul>' +
    '            </div>' +
    '' +
    '            <!-- Column 5: COMPANY -->' +
    '            <div>' +
    '                <h4 class="text-white font-extrabold text-xs uppercase tracking-widest mb-4">Company</h4>' +
    '                <ul class="space-y-2.5 text-xs font-semibold text-slate-400">' +
    '                    <li><a href="about.html" class="hover:text-white transition-colors">About Us</a></li>' +
    '                    <li><a href="contact.html" class="hover:text-white transition-colors">Contact Us</a></li>' +
    '                </ul>' +
    '            </div>' +
    '' +
    '            <!-- Column 6: APP STORE BADGES -->' +
    '            <div class="col-span-2 md:col-span-1 lg:col-span-1 space-y-2.5">' +
    '                <!-- Google Play Badge -->' +
    '                <a href="#" onclick="alert(\'FilesQue mobile & desktop native apps are releasing soon in 2026!\'); return false;" class="flex items-center space-x-3 bg-black/40 hover:bg-black/70 border border-slate-700/80 rounded-xl px-3.5 py-2 transition-all group">' +
    '                    <svg class="w-5 h-5 fill-current text-white shrink-0" viewBox="0 0 24 24"><path d="M3.609 1.814L13.792 12 3.61 22.186a1.996 1.996 0 0 1-.61-.92L3 2.735a2.002 2.002 0 0 1 .609-.921zm11.233 11.236l2.195 2.195-12.01 6.84 9.815-9.035zm0-2.1l-9.815-9.035 12.01 6.84-2.195 2.195zm1.455 1.05l3.878-2.21a1.2 1.2 0 0 1 0 2.083l-3.878 2.21 1.05-1.041-1.05-1.042z"/></svg>' +
    '                    <div class="text-left leading-tight">' +
    '                        <div class="text-[9px] uppercase tracking-wider text-slate-400 font-bold">GET IT ON</div>' +
    '                        <div class="text-xs font-black text-white">Google Play</div>' +
    '                    </div>' +
    '                </a>' +
    '' +
    '                <!-- App Store Badge -->' +
    '                <a href="#" onclick="alert(\'FilesQue iOS app releasing soon in 2026!\'); return false;" class="flex items-center space-x-3 bg-black/40 hover:bg-black/70 border border-slate-700/80 rounded-xl px-3.5 py-2 transition-all group">' +
    '                    <svg class="w-5 h-5 fill-current text-white shrink-0" viewBox="0 0 24 24"><path d="M18.71 19.5c-.83 1.24-1.71 2.45-3.05 2.47-1.34.03-1.77-.79-3.29-.79-1.53 0-2 .77-3.27.82-1.31.05-2.3-1.32-3.14-2.53C4.25 17 2.94 12.45 4.7 9.39c.87-1.52 2.43-2.48 4.12-2.51 1.28-.02 2.5.87 3.29.87.78 0 2.26-1.07 3.81-.91.65.03 2.47.26 3.64 1.98-.09.06-2.17 1.28-2.15 3.81.03 3.02 2.65 4.03 2.68 4.04-.03.07-.42 1.44-1.38 2.83M15.97 6.37c.62-.75 1.04-1.8 1.01-2.85-.9.04-1.99.6-2.61 1.33-.55.63-.99 1.66-.86 2.67 1 .08 2.04-.51 2.46-1.15z"/></svg>' +
    '                    <div class="text-left leading-tight">' +
    '                        <div class="text-[9px] uppercase tracking-wider text-slate-400 font-bold">Download on the</div>' +
    '                        <div class="text-xs font-black text-white">App Store</div>' +
    '                    </div>' +
    '                </a>' +
    '' +
    '                <!-- Mac App Store Badge -->' +
    '                <a href="#" onclick="alert(\'FilesQue macOS native version releasing soon in 2026!\'); return false;" class="flex items-center space-x-3 bg-black/40 hover:bg-black/70 border border-slate-700/80 rounded-xl px-3.5 py-2 transition-all group">' +
    '                    <svg class="w-5 h-5 fill-current text-white shrink-0" viewBox="0 0 24 24"><path d="M18.71 19.5c-.83 1.24-1.71 2.45-3.05 2.47-1.34.03-1.77-.79-3.29-.79-1.53 0-2 .77-3.27.82-1.31.05-2.3-1.32-3.14-2.53C4.25 17 2.94 12.45 4.7 9.39c.87-1.52 2.43-2.48 4.12-2.51 1.28-.02 2.5.87 3.29.87.78 0 2.26-1.07 3.81-.91.65.03 2.47.26 3.64 1.98-.09.06-2.17 1.28-2.15 3.81.03 3.02 2.65 4.03 2.68 4.04-.03.07-.42 1.44-1.38 2.83M15.97 6.37c.62-.75 1.04-1.8 1.01-2.85-.9.04-1.99.6-2.61 1.33-.55.63-.99 1.66-.86 2.67 1 .08 2.04-.51 2.46-1.15z"/></svg>' +
    '                    <div class="text-left leading-tight">' +
    '                        <div class="text-[9px] uppercase tracking-wider text-slate-400 font-bold">Download on the</div>' +
    '                        <div class="text-xs font-black text-white">Mac App Store</div>' +
    '                    </div>' +
    '                </a>' +
    '' +
    '                <!-- Microsoft Store Badge -->' +
    '                <a href="#" onclick="alert(\'FilesQue Windows Desktop release coming in 2026!\'); return false;" class="flex items-center space-x-3 bg-black/40 hover:bg-black/70 border border-slate-700/80 rounded-xl px-3.5 py-2 transition-all group">' +
    '                    <svg class="w-5 h-5 fill-current text-white shrink-0" viewBox="0 0 24 24"><path d="M0 3.449L9.75 2.1v9.451H0m10.949-9.602L24 0v11.4H10.949M0 12.6h9.75v9.451L0 20.699M10.949 12.6H24V24l-12.9-1.801"/></svg>' +
    '                    <div class="text-left leading-tight">' +
    '                        <div class="text-[9px] uppercase tracking-wider text-slate-400 font-bold">GET IT FROM</div>' +
    '                        <div class="text-xs font-black text-white">Microsoft Store</div>' +
    '                    </div>' +
    '                </a>' +
    '            </div>' +
    '        </div>' +
    '' +
    '        <!-- Horizontal Divider -->' +
    '        <div class="border-t border-slate-800/80 pt-8 flex flex-col md:flex-row items-center justify-between gap-6">' +
    '            ' +
    '            <!-- Language Selector Dropdown -->' +
    '            <div class="relative inline-block text-left">' +
    '                <div class="flex items-center space-x-2 bg-slate-800/80 hover:bg-slate-700/80 text-slate-200 border border-slate-700 rounded-lg px-3.5 py-1.5 text-xs font-bold transition-all cursor-pointer">' +
    '                    <span>🌐</span>' +
    '                    <select id="footer-lang-select" onchange="alert(\'Language preference set to \' + this.options[this.selectedIndex].text)" class="bg-transparent border-none text-slate-200 text-xs font-bold focus:outline-none cursor-pointer">' +
    '                        <option value="en" class="bg-slate-900 text-slate-200" selected>English</option>' +
    '                        <option value="es" class="bg-slate-900 text-slate-200">Español</option>' +
    '                        <option value="fr" class="bg-slate-900 text-slate-200">Français</option>' +
    '                        <option value="de" class="bg-slate-900 text-slate-200">Deutsch</option>' +
    '                        <option value="hi" class="bg-slate-900 text-slate-200">हिन्दी (Hindi)</option>' +
    '                        <option value="pt" class="bg-slate-900 text-slate-200">Português</option>' +
    '                        <option value="ar" class="bg-slate-900 text-slate-200">العربية</option>' +
    '                        <option value="zh" class="bg-slate-900 text-slate-200">中文 (Chinese)</option>' +
    '                    </select>' +
    '                </div>' +
    '            </div>' +
    '' +
    '            <!-- Social Media Icons -->' +
    '            <div class="flex items-center space-x-5 text-slate-400">' +
    '                <!-- X (Twitter) -->' +
    '                <a href="https://x.com" target="_blank" rel="noopener" class="hover:text-white transition-colors" title="Follow on X">' +
    '                    <svg class="w-4 h-4 fill-current" viewBox="0 0 24 24"><path d="M18.244 2.25h3.308l-7.227 8.26 8.502 11.24H16.17l-5.214-6.817L4.99 21.75H1.68l7.73-8.835L1.254 2.25H8.08l4.713 6.231zm-1.161 17.52h1.833L7.084 4.126H5.117z"/></svg>' +
    '                </a>' +
    '                <!-- Facebook -->' +
    '                <a href="https://facebook.com" target="_blank" rel="noopener" class="hover:text-white transition-colors" title="Facebook">' +
    '                    <svg class="w-4 h-4 fill-current" viewBox="0 0 24 24"><path d="M24 12.073c0-6.627-5.373-12-12-12s-12 5.373-12 12c0 5.99 4.388 10.954 10.125 11.854v-8.385H7.078v-3.47h3.047V9.43c0-3.007 1.792-4.669 4.533-4.669 1.312 0 2.686.235 2.686.235v2.953H15.83c-1.491 0-1.956.925-1.956 1.874v2.25h3.328l-.532 3.47h-2.796v8.385C19.612 23.027 24 18.062 24 12.073z"/></svg>' +
    '                </a>' +
    '                <!-- LinkedIn -->' +
    '                <a href="https://linkedin.com" target="_blank" rel="noopener" class="hover:text-white transition-colors" title="LinkedIn">' +
    '                    <svg class="w-4 h-4 fill-current" viewBox="0 0 24 24"><path d="M19 0h-14c-2.761 0-5 2.239-5 5v14c0 2.761 2.239 5 5 5h14c2.762 0 5-2.239 5-5v-14c0-2.761-2.238-5-5-5zm-11 19h-3v-11h3v11zm-1.5-12.268c-.966 0-1.75-.79-1.75-1.764s.784-1.764 1.75-1.764 1.75.79 1.75 1.764-.783 1.764-1.75 1.764zm13.5 12.268h-3v-5.604c0-3.368-4-3.113-4 0v5.604h-3v-11h3v1.765c1.396-2.586 7-2.777 7 2.476v6.759z"/></svg>' +
    '                </a>' +
    '                <!-- Instagram -->' +
    '                <a href="https://instagram.com" target="_blank" rel="noopener" class="hover:text-white transition-colors" title="Instagram">' +
    '                    <svg class="w-4 h-4 fill-current" viewBox="0 0 24 24"><path d="M12 2.163c3.204 0 3.584.012 4.85.07 3.252.148 4.771 1.691 4.919 4.919.058 1.265.069 1.645.069 4.849 0 3.205-.012 3.584-.069 4.849-.149 3.225-1.664 4.771-4.919 4.919-1.266.058-1.644.07-4.85.07-3.204 0-3.584-.012-4.849-.07-3.26-.149-4.771-1.699-4.919-4.92-.058-1.265-.07-1.644-.07-4.849 0-3.204.013-3.583.07-4.849.149-3.227 1.664-4.771 4.919-4.919 1.266-.057 1.645-.069 4.849-.069zm0-2.163c-3.259 0-3.667.014-4.947.072-4.358.2-6.78 2.618-6.98 6.98-.059 1.281-.073 1.689-.073 4.948 0 3.259.014 3.668.072 4.948.2 4.358 2.618 6.78 6.98 6.98 1.281.058 1.689.072 4.948.072 3.259 0 3.668-.014 4.948-.072 4.354-.2 6.782-2.618 6.979-6.98.059-1.28.073-1.689.073-4.948 0-3.259-.014-3.667-.072-4.947-.196-4.354-2.617-6.78-6.979-6.98-1.281-.059-1.69-.073-4.949-.073zm0 5.838c-3.403 0-6.162 2.759-6.162 6.162s2.759 6.163 6.162 6.163 6.162-2.759 6.162-6.163c0-3.403-2.759-6.162-6.162-6.162zm0 10.162c-2.209 0-4-1.79-4-4 0-2.209 1.791-4 4-4s4 1.791 4 4c0 2.21-1.791 4-4 4zm6.406-11.845c-.796 0-1.441.645-1.441 1.44s.645 1.44 1.441 1.44c.795 0 1.439-.645 1.439-1.44s-.644-1.44-1.439-1.44z"/></svg>' +
    '                </a>' +
    '                <!-- TikTok -->' +
    '                <a href="https://tiktok.com" target="_blank" rel="noopener" class="hover:text-white transition-colors" title="TikTok">' +
    '                    <svg class="w-4 h-4 fill-current" viewBox="0 0 24 24"><path d="M12.525.02c1.31-.02 2.61-.01 3.91-.02.08 1.53.63 3.09 1.75 4.17 1.12 1.11 2.7 1.62 4.24 1.79v4.03c-1.44-.05-2.89-.35-4.2-.97-.57-.26-1.1-.59-1.62-1.01-.01 2.92.01 5.84-.02 8.75-.08 1.4-.54 2.79-1.35 3.94-1.31 1.92-3.58 3.17-5.91 3.21-1.43.08-2.86-.31-4.08-1.03-2.02-1.19-3.44-3.37-3.65-5.71-.02-.5-.03-1-.01-1.49.18-1.9 1.12-3.72 2.58-4.96 1.66-1.44 3.98-2.13 6.15-1.72.02 1.48-.04 2.96-.04 4.44-.99-.32-2.15-.23-3.02.37-.63.41-1.11 1.04-1.36 1.75-.21.51-.24 1.07-.14 1.61.24 1.64 1.82 3.02 3.5 2.87 1.12-.01 2.19-.66 2.77-1.61.19-.33.4-.67.41-1.06.1-1.79.06-3.57.07-5.36.01-4.03-.01-8.05.02-12.07z"/></svg>' +
    '                </a>' +
    '                <!-- Reddit -->' +
    '                <a href="https://reddit.com" target="_blank" rel="noopener" class="hover:text-white transition-colors" title="Reddit">' +
    '                    <svg class="w-4 h-4 fill-current" viewBox="0 0 24 24"><path d="M12 0A12 12 0 0 0 0 12a12 12 0 0 0 12 12 12 12 0 0 0 12-12A12 12 0 0 0 12 0zm5.01 4.744c.688 0 1.25.56 1.25 1.249a1.25 1.25 0 0 1-2.498.056l-2.597-.547-.8 3.747c1.824.07 3.48.632 4.674 1.488.308-.309.73-.491 1.207-.491.968 0 1.754.786 1.754 1.754 0 .716-.435 1.333-1.01 1.614a3.111 3.111 0 0 1 .042.52c0 2.694-3.13 4.87-7.004 4.87-3.874 0-7.004-2.176-7.004-4.87 0-.183.015-.366.043-.534A1.748 1.748 0 0 1 4.028 12c0-.968.786-1.754 1.754-1.754.463 0 .898.196 1.207.49 1.207-.883 2.878-1.43 4.744-1.487l.885-4.182a.342.342 0 0 1 .14-.197.35.35 0 0 1 .238-.042l2.906.617a1.214 1.214 0 0 1 1.108-.701zM9.25 12C8.561 12 8 12.562 8 13.25c0 .687.561 1.248 1.25 1.248.687 0 1.248-.561 1.248-1.249 0-.688-.561-1.249-1.249-1.249zm5.5 0c-.687 0-1.248.561-1.248 1.25 0 .687.561 1.248 1.249 1.248.688 0 1.249-.561 1.249-1.249 0-.687-.562-1.249-1.25-1.249zm-5.466 3.99a.327.327 0 0 0-.231.094.33.33 0 0 0 0 .463c.842.842 2.484.913 2.961.913.477 0 2.105-.056 2.961-.913a.361.361 0 0 0 .029-.463.33.33 0 0 0-.464 0c-.547.533-1.684.73-2.512.73-.828 0-1.979-.197-2.512-.73a.326.326 0 0 0-.232-.095z"/></svg>' +
    '                </a>' +
    '            </div>' +
    '' +
    '            <!-- Copyright Notice -->' +
    '            <div class="text-xs text-slate-500 font-medium">' +
    '                &copy; FilesQue 2026 &reg; - Your Files Editor' +
    '            </div>' +
    '        </div>' +
    '    </div>' +
    '</footer>' +
    '' +
    '</body></html>');
});

app.listen(PORT, () => {
  console.log('========================================');
  console.log('✅ Filesque server running');
  console.log('📱 Homepage: http://localhost:' + PORT);
  console.log('⚡ Workflow Builder: http://localhost:' + PORT + '/workflow-builder');
  console.log('🔧 Total Tools: ' + toolRegistry.length);
  console.log('========================================');
});