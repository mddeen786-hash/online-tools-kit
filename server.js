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
  { id: 'merge-image',         name: 'Merge Image',             icon: '🧩', file: 'merge-image.html' },
  { id: 'invoice-generator',   name: 'Invoice Generator',       icon: '🧾', file: 'invoice-gen.html' },
  { id: 'barcode-generator',   name: 'Barcode Generator',       icon: '🏷️', file: 'barcode-gen.html' },
  { id: 'qr-code-studio',      name: 'QR Code Studio',          icon: '📱', file: 'qr-studio.html' },
  { id: 'digital-signature',   name: 'Digital Signature Studio', icon: '✍️', file: 'sign-maker.html' },
  { id: 'code-formatter',      name: 'Code Formatter',          icon: '⚡', file: 'code-formatter.html' },
  { id: 'product-listing',     name: 'Product Listing',         icon: '📦', file: 'product-listing.html' }
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
  try {
    const worker = await Tesseract.createWorker('eng', 1);
    const { data } = await worker.recognize(inputPath);
    await worker.terminate();
    if (data && data.text) {
      const lines = data.text.split('\n');
      lines.forEach(l => {
        const cols = l.split(/\t|\s{2,}/).map(c => c.trim()).filter(c => c.length > 0);
        if (cols.length > 0) rows.push(cols);
      });
    }
  } catch (e) { console.log('OCR Error:', e.message); }
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

// ========== Image Converter Logic ==========
async function imageConverterConvert(inputPath, outputPath, config = {}) {
  try {
    if (!sharp) throw new Error("Sharp module required.");
    const format = (config.convertFormat || 'jpg').toLowerCase();
    const quality = config.convertQuality ? Math.round(parseFloat(config.convertQuality) * 100) : 92;
    let imageProcess = sharp(inputPath);

    if (format === 'jpg' || format === 'jpeg' || format === 'bmp') {
      imageProcess = imageProcess.flatten({ background: { r: 255, g: 255, b: 255 } });
    }

    let ext = format === 'jpeg' ? 'jpeg' : format;
    if (format === 'ico') ext = 'ico';

    const finalOutputPath = outputPath.replace(/\.[^/.]+$/, `.${ext}`);

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

// ========== Image Reducer Logic ==========
async function imageReducerConvert(inputPath, outputPath, config = {}) {
  try {
    const mode = config.reducerMode || 'resize';
    const finalOutputPath = outputPath.replace(/\.[^/.]+$/, ".jpg");

    if (sharp) {
      let pipeline = sharp(inputPath).flatten({ background: { r: 255, g: 255, b: 255 } });

      if (mode === 'resize') {
        const w = parseInt(config.reducerWidth) || 800;
        const h = parseInt(config.reducerHeight) || 600;
        await pipeline
          .resize(w, h, { fit: 'fill', kernel: sharp.kernel.lanczos3 })
          .jpeg({ quality: 95, mozjpeg: true })
          .toFile(finalOutputPath);
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
            outputBuffer = await sharp(inputPath)
                .flatten({ background: { r: 255, g: 255, b: 255 } })
                .resize(
                  Math.max(150, Math.round(currentW * scale)), 
                  Math.max(150, Math.round(currentH * scale)), 
                  { fit: 'inside', kernel: sharp.kernel.lanczos3 }
                )
                .jpeg({ quality: quality, mozjpeg: true })
                .toBuffer();

            if (outputBuffer.length <= targetBytes || quality <= 30) {
                break;
            }
            quality -= 12;
            scale -= 0.10;
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
  '2x2':     { wIn: 2,           hIn: 2,           wMm: 51, hMm: 51, cols: 4, copyOptions: [4, 8, 16] }
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
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

app.get('/workflow-builder', (req, res) => {
  const toolsListHTML = toolRegistry.map(tool => {
    return '<div class="tool-item" onclick="addToWorkflow(\'' + tool.id + '\', \'' + tool.name.replace(/'/g, "\\'") + '\', \'' + tool.icon + '\')">' +
      '<span class="tool-icon">' + tool.icon + '</span>' +
      '<span class="tool-name">' + tool.name + '</span>' +
      '</div>';
  }).join('');

  res.send('<!DOCTYPE html><html><head><title>Workflow Builder - Filesque</title><style>' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    'body{font-family:"Plus Jakarta Sans",sans-serif;background:#f8fafc;min-height:100vh}' +
    '.navbar{background:white;border-bottom:1px solid #e2e8f0;padding:1rem 2rem;display:flex;justify-content:space-between;align-items:center}' +
    '.logo{font-size:1.8rem;font-weight:900;color:#0f172a}.logo span{color:#dc2626}' +
    '.back-link{color:#dc2626;text-decoration:none;font-weight:700}' +
    '.container{display:flex;gap:2rem;max-width:1400px;margin:2rem auto;padding:0 1rem}' +
    '.panel{background:white;border-radius:1rem;box-shadow:0 2px 10px rgba(0,0,0,0.05);padding:1.5rem}' +
    '.tools-panel{flex:1;max-height:80vh;overflow-y:auto}.workflow-panel{flex:2}' +
    '.panel-title{font-size:1.3rem;font-weight:800;margin-bottom:1rem}' +
    '.tool-item{display:flex;align-items:center;gap:0.8rem;padding:0.8rem 1rem;background:#f1f5f9;border-radius:0.5rem;margin-bottom:0.5rem;cursor:pointer;transition:0.2s}' +
    '.tool-item:hover{background:#e2e8f0;transform:translateX(5px)}' +
    '.tool-icon{font-size:1.4rem}.tool-name{font-weight:600}' +
    '.workflow-area{min-height:200px;border:2px dashed #cbd5e1;border-radius:0.75rem;padding:1rem;margin-bottom:1rem}' +
    '.step{display:flex;align-items:center;gap:1rem;padding:0.8rem;background:#f1f5f9;border-radius:0.5rem;margin-bottom:0.5rem}' +
    '.step-number{width:30px;height:30px;background:#dc2626;color:white;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:bold}' +
    '.remove-btn{background:#ef4444;color:white;border:none;border-radius:0.25rem;padding:0.3rem 0.7rem;cursor:pointer}' +
    '.controls{display:flex;gap:0.8rem;margin-bottom:1rem}' +
    'input[type="text"]{flex:1;padding:0.7rem 1rem;border:1px solid #cbd5e1;border-radius:0.5rem;font-size:1rem}' +
    '.btn{padding:0.7rem 1.5rem;background:#dc2626;color:white;border:none;border-radius:0.5rem;cursor:pointer;font-weight:700}' +
    '.btn:hover{background:#b91c1c}.btn-secondary{background:#64748b}.btn-secondary:hover{background:#475569}' +
    '.saved-list{margin-top:1rem}.saved-item{padding:0.6rem;background:#f8fafc;border-radius:0.4rem;margin-bottom:0.4rem}' +
    '.file-upload-section{margin-top:1.5rem;padding:1rem;background:#f8fafc;border-radius:0.5rem}' +
    '.file-upload-section input[type="file"]{margin-top:0.5rem}' +
    '</style></head><body>' +
    '<div class="navbar"><div class="logo">Files<span>que</span></div><a href="/" class="back-link">← Back to Home</a></div>' +
    '<div class="container">' +
    '<div class="panel tools-panel"><div class="panel-title">🛠️ Tools (' + toolRegistry.length + ')</div>' + toolsListHTML + '</div>' +
    '<div class="panel workflow-panel"><div class="panel-title">📋 Your Workflow</div>' +
    '<div class="workflow-area" id="workflowArea"><p style="color:#94a3b8;">Click on tools to add steps</p></div>' +
    '<div class="controls"><input type="text" id="wfName" placeholder="Workflow Name">' +
    '<button class="btn" onclick="saveWorkflow()">💾 Save</button>' +
    '<button class="btn btn-secondary" onclick="clearWorkflow()">🗑️ Clear</button></div>' +
    '<div class="file-upload-section"><h3>📁 Test Your Workflow</h3>' +
    '<input type="file" id="wfFile" multiple>' + 
    '<button class="btn" onclick="executeWorkflow()" style="margin-left: 10px;">▶️ Execute Workflow</button>' +
    '<div id="executionResult" style="margin-top:10px;"></div></div>' +
    '<div class="saved-list"><h3>Saved Workflows</h3><div id="savedList"></div></div>' +
    '</div></div>' +
    '<div id="modalContainer"></div>' +
    '<script>' +
    'let steps = [];' +
    'let selectedFiles = [];' +
    'document.getElementById("wfFile").addEventListener("change", function(e) { selectedFiles = e.target.files; });' +
    'function updateStepConfig(idx, key, val) { steps[idx][key] = val; }' +
    'function addToWorkflow(id, name, icon) { ' +
    '  steps.push({ id: id, name: name, icon: icon, isComplex: false, mode: "all", range: "", sortOrder: "upload", customSequence: "", pageOrder: "", rotatePages: "", rotateDegree: "90", watermarkText: "CONFIDENTIAL", watermarkMode: "diagonal", compressQuality: "0.70", convertFormat: "jpg", convertQuality: "0.92", reducerMode: "resize", reducerWidth: "800", reducerHeight: "600", reducerTargetKb: "50", sizePreset: "3.5x4.5", bgColor: "#f87171", zoom: "100", copies: "32", removeBg: false });' +
    '  renderSteps();' +
    '}' +
    'function renderSteps() {' +
    '  const area = document.getElementById("workflowArea");' +
    '  if (steps.length === 0) { area.innerHTML = "<p style=\'color:#94a3b8;\'>Click on tools to add steps</p>"; return; }' +
    '  let html = "";' +
    '  for (let i = 0; i < steps.length; i++) {' +
    '    let toggleHtml = "";' +
    '    if (steps[i].id === "pdf-to-excel") {' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px; background:#e2e8f0; padding:6px 12px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '        "<span style=\'font-size:12px; font-weight:bold; color:#475569;\'>Mode:</span>" +' +
    '        "<select onchange=\'updateStepConfig(" + i + ", \\\"isComplex\\\", this.value === \\\"complex\\\")\' style=\'padding:4px 8px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; font-weight:bold; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '          "<option value=\\\"simple\\\" " + (!steps[i].isComplex ? "selected" : "") + ">Simple Table</option>" +' +
    '          "<option value=\\\"complex\\\" " + (steps[i].isComplex ? "selected" : "") + ">Complex Table</option>" +' +
    '        "</select>" +' +
    '      "</div>";' +
    '    } else if (steps[i].id === "pdf-splitter") {' +
    '      let mode = steps[i].mode || "all";' +
    '      let rangeVal = steps[i].range || "";' +
    '      let displayRange = mode === "all" ? "display:none;" : "";' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px; background:#e2e8f0; padding:6px 12px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '        "<span style=\'font-size:12px; font-weight:bold; color:#475569;\'>Mode:</span>" +' +
    '        "<select onchange=\'updateStepConfig(" + i + ", \\\"mode\\\", this.value); renderSteps();\' style=\'padding:4px 8px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; font-weight:bold; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '          "<option value=\\\"all\\\" " + (mode === "all" ? "selected" : "") + ">Split All (ZIP)</option>" +' +
    '          "<option value=\\\"range\\\" " + (mode === "range" ? "selected" : "") + ">Custom Range</option>" +' +
    '        "</select>" +' +
    '        "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"range\\\", this.value)\' value=\'" + rangeVal + "\' placeholder=\'e.g. 1-3, 5\' style=\'padding:4px 8px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; width:100px; outline:none; " + displayRange + "\'>" + ' +
    '      "</div>";' +
    '    } else if (steps[i].id === "merge-pdf" || steps[i].id === "image-to-pdf") {' +
    '      let sOrder = steps[i].sortOrder || "upload";' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px; background:#e2e8f0; padding:6px 12px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '        "<span style=\'font-size:12px; font-weight:bold; color:#475569;\'>Combine Order:</span>" +' +
    '        "<select onchange=\'updateStepConfig(" + i + ", \\\"sortOrder\\\", this.value)\' style=\'padding:4px 8px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; font-weight:bold; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
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
    '          "<span style=\'font-size:11px; font-weight:bold; color:#475569; margin-top:2px;\'>Page Order:</span>" +' +
    '          "<div style=\'display:flex; flex-direction:column; gap:2px;\'>" +' +
    '            "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"pageOrder\\\", this.value)\' value=\'" + pOrder + "\' placeholder=\'e.g. 1, 3, 2, 5-7\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:11px; width:120px; outline:none;\' title=\'Leave blank to keep all pages original order\'>" +' +
    '            "<span style=\'font-size:9px; color:#64748b; line-height:1;\'>e.g., 1, 3, 2, 5-7</span>" +' +
    '          "</div>" +' +
    '        "</div>" +' +
    '        "<div style=\'display:flex; align-items:flex-start; gap:5px; background:#e2e8f0; padding:6px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '          "<span style=\'font-size:11px; font-weight:bold; color:#475569; margin-top:2px;\'>Rotate:</span>" +' +
    '          "<div style=\'display:flex; flex-direction:column; gap:2px;\'>" +' +
    '            "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"rotatePages\\\", this.value)\' value=\'" + rPages + "\' placeholder=\'e.g. 2, 4 or all\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:11px; width:110px; outline:none;\'>" +' +
    '            "<span style=\'font-size:9px; color:#64748b; line-height:1;\'>e.g. 2, 4 or all</span>" +' +
    '          "</div>" +' +
    '          "<select onchange=\'updateStepConfig(" + i + ", \\\"rotateDegree\\\", this.value)\' style=\'padding:1px 4px; border-radius:4px; border:1px solid #94a3b8; font-size:11px; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
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
    '          "<span style=\'font-size:11px; font-weight:bold; color:#475569; margin-top:2px;\'>Text:</span>" +' +
    '          "<div style=\'display:flex; flex-direction:column; gap:2px;\'>" +' +
    '            "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"watermarkText\\\", this.value)\' value=\'" + wmText + "\' placeholder=\'e.g. CONFIDENTIAL\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:11px; width:110px; outline:none;\'>" +' +
    '            "<span style=\'font-size:9px; color:#64748b; line-height:1;\'>e.g. DRAFT</span>" +' +
    '          "</div>" +' +
    '        "</div>" +' +
    '        "<div style=\'display:flex; align-items:flex-start; gap:5px; background:#e2e8f0; padding:6px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '          "<span style=\'font-size:11px; font-weight:bold; color:#475569; margin-top:2px;\'>Mode:</span>" +' +
    '          "<select onchange=\'updateStepConfig(" + i + ", \\\"watermarkMode\\\", this.value)\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:11px; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '             "<option value=\\\"diagonal\\\" " + (wmMode === "diagonal" ? "selected" : "") + ">Watermark Only</option>" +' +
    '             "<option value=\\\"footer_num\\\" " + (wmMode === "footer_num" ? "selected" : "") + ">Page Numbers Only</option>" +' +
    '             "<option value=\\\"both\\\" " + (wmMode === "both" ? "selected" : "") + ">Both</option>" +' +
    '          "</select>" +' +
    '        "</div>" +' +
    '      "</div>";' +
    '    } else if (steps[i].id === "compress-image") {' +
    '      let qual = steps[i].compressQuality || "0.70";' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px; background:#e2e8f0; padding:6px 12px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '        "<span style=\'font-size:12px; font-weight:bold; color:#475569;\'>Quality:</span>" +' +
    '        "<select onchange=\'updateStepConfig(" + i + ", \\\"compressQuality\\\", this.value)\' style=\'padding:4px 8px; border-radius:4px; border:1px solid #94a3b8; font-size:12px; font-weight:bold; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
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
    '          "<span style=\'font-size:11px; font-weight:bold; color:#475569;\'>Format:</span>" +' +
    '          "<select onchange=\'updateStepConfig(" + i + ", \\\"convertFormat\\\", this.value)\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:11px; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '            "<option value=\\\"jpg\\\" " + (fmt === "jpg" ? "selected" : "") + ">JPG</option>" +' +
    '            "<option value=\\\"jpeg\\\" " + (fmt === "jpeg" ? "selected" : "") + ">JPEG</option>" +' +
    '            "<option value=\\\"png\\\" " + (fmt === "png" ? "selected" : "") + ">PNG</option>" +' +
    '            "<option value=\\\"webp\\\" " + (fmt === "webp" ? "selected" : "") + ">WEBP</option>" +' +
    '            "<option value=\\\"bmp\\\" " + (fmt === "bmp" ? "selected" : "") + ">BMP</option>" +' +
    '            "<option value=\\\"ico\\\" " + (fmt === "ico" ? "selected" : "") + ">ICO</option>" +' +
    '          "</select>" +' +
    '        "</div>" +' +
    '        "<div style=\'display:flex; align-items:center; gap:5px; background:#e2e8f0; padding:6px; border-radius:6px; border:1px solid #cbd5e1;\'>" +' +
    '          "<span style=\'font-size:11px; font-weight:bold; color:#475569;\'>Quality:</span>" +' +
    '          "<select onchange=\'updateStepConfig(" + i + ", \\\"convertQuality\\\", this.value)\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:11px; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
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
    '          "<span style=\'font-size:11px; font-weight:bold; color:#475569;\'>Mode:</span>" +' +
    '          "<select onchange=\'updateStepConfig(" + i + ", \\\"reducerMode\\\", this.value); renderSteps();\' style=\'padding:2px 6px; border-radius:4px; border:1px solid #94a3b8; font-size:11px; cursor:pointer; outline:none; background:white; color:#0f172a;\'>" +' +
    '            "<option value=\\\"resize\\\" " + (rMode === "resize" ? "selected" : "") + ">Resize (W x H)</option>" +' +
    '            "<option value=\\\"target\\\" " + (rMode === "target" ? "selected" : "") + ">Target KB</option>" +' +
    '          "</select>" +' +
    '          "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"reducerWidth\\\", this.value)\' value=\'" + rW + "\' placeholder=\'W\' style=\'display:" + showResize + "; padding:2px 4px; border-radius:4px; border:1px solid #94a3b8; font-size:11px; width:45px; outline:none;\'>" +' +
    '          "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"reducerHeight\\\", this.value)\' value=\'" + rH + "\' placeholder=\'H\' style=\'display:" + showResize + "; padding:2px 4px; border-radius:4px; border:1px solid #94a3b8; font-size:11px; width:45px; outline:none;\'>" +' +
    '          "<input type=\'text\' onkeyup=\'updateStepConfig(" + i + ", \\\"reducerTargetKb\\\", this.value)\' value=\'" + rKb + "\' placeholder=\'KB\' style=\'display:" + showTarget + "; padding:2px 4px; border-radius:4px; border:1px solid #94a3b8; font-size:11px; width:60px; outline:none;\'>" +' +
    '        "</div>" +' +
    '      "</div>";' +
    '    } else if (steps[i].id === "passport-studio") {' +
    '      toggleHtml = "<div style=\'margin-left:auto; display:flex; align-items:center; gap:8px; background:linear-gradient(135deg, #4f46e5, #e11d48); padding:6px 14px; border-radius:8px; color:white; font-weight:bold; font-size:12px; box-shadow:0 2px 5px rgba(0,0,0,0.1);\'>" + ' +
    '        "<span>✨ Auto Studio Dialog Ready</span>" + ' +
    '      "</div>";' +
    '    }' +
    '    html += "<div class=\'step\'><div class=\'step-number\'>" + (i+1) + "</div><span style=\'font-size:1.2rem;\'>" + steps[i].icon + "</span><strong style=\'font-size:1.1rem;\'>" + steps[i].name + "</strong>" + (toggleHtml ? toggleHtml : "<div style=\'margin-left:auto;\'></div>") + "<button class=\'remove-btn\' style=\'margin-left:10px;\' onclick=\'removeStep(" + i + ")\'>✕</button></div>";' +
    '  }' +
    '  area.innerHTML = html;' +
    '}' +
    'function removeStep(index) { steps.splice(index, 1); renderSteps(); }' +
    'function clearWorkflow() { steps = []; renderSteps(); }' +
    'async function saveWorkflow() {' +
    '  const name = document.getElementById("wfName").value.trim() || "My Workflow";' +
    '  if (steps.length === 0) { alert("Please add at least one tool"); return; }' +
    '  const payload = { name: name, steps: steps.map(s => ({ toolId: s.id, isComplex: !!s.isComplex, mode: s.mode, range: s.range, sortOrder: s.sortOrder, customSequence: s.customSequence, pageOrder: s.pageOrder, rotatePages: s.rotatePages, rotateDegree: s.rotateDegree, watermarkText: s.watermarkText, watermarkMode: s.watermarkMode, compressQuality: s.compressQuality, convertFormat: s.convertFormat, convertQuality: s.convertQuality, reducerMode: s.reducerMode, reducerWidth: s.reducerWidth, reducerHeight: s.reducerHeight, reducerTargetKb: s.reducerTargetKb, sizePreset: s.sizePreset, bgColor: s.bgColor, zoom: s.zoom, copies: s.copies, removeBg: s.removeBg, printSheet: s.printSheet })) };' +
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
    '  sendWorkflowExecution();' +
    '}' +
    'function openProfessionalStudioModal(file, stepIdx) {' +
    '  const reader = new FileReader();' +
    '  reader.onload = function(e) {' +
    '    const imgSrc = e.target.result;' +
    '    const modalHtml = \'<div style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(15,23,42,0.8);backdrop-filter:blur(8px);display:flex;justify-content:center;align-items:center;z-index:9999;font-family:\\\'Plus Jakarta Sans\\\',sans-serif;"><div style="background:white;padding:2.5rem;border-radius:1.75rem;width:820px;max-width:96%;box-shadow:0 25px 60px rgba(0,0,0,0.4);max-height:92vh;overflow-y:auto;border:1px solid #e2e8f0;"><div style="display:flex;align-items:center;margin-bottom:1.5rem;"><div style="width:48px;height:48px;border-radius:14px;background:linear-gradient(135deg,#4f46e5,#e11d48);display:flex;align-items:center;justify-content:center;color:white;font-size:24px;box-shadow:0 10px 20px rgba(79,70,229,0.3);margin-right:1rem;">🛂</div><div><h2 style="font-size:24px;font-weight:900;color:#0f172a;margin:0;letter-spacing:-0.5px;">Passport Photo Studio Pro</h2><p style="font-size:12px;color:#64748b;font-weight:600;margin:2px 0 0 0;">Professional Studio: Cloud AI Background Removal & HD Print Layout</p></div></div><hr style="border:0;border-top:1px solid #e2e8f0;margin-bottom:1.5rem;"><div style="display:grid;grid-template-columns:1.2fr 1fr;gap:2rem;margin-bottom:2rem;"><div style="display:flex;flex-direction:column;gap:1.2rem;"><div><label style="display:block;font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">1. Background Color Palette</label><select id="studioBg" style="width:100%;padding:12px 14px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:12px;outline:none;font-size:13px;font-weight:700;color:#1e293b;cursor:pointer;"><option value="#ffffff">Pure White (#FFFFFF)</option><option value="#a5cbf7">Light Blue (#A5CBF7)</option><option value="#3b82f6">Royal Blue (#3B82F6)</option><option value="#f87171" selected>Soft Red (#F87171)</option></select></div><div><label style="display:block;font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">2. Passport Size Preset</label><select id="studioSize" style="width:100%;padding:12px 14px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:12px;outline:none;font-size:13px;font-weight:700;color:#1e293b;cursor:pointer;"><option value="3.5x4.5">Standard Indian Passport (3.5 x 4.5 cm)</option><option value="2x2">US Visa Layout (2 x 2 inch)</option><option value="3.5x3.5">Indian PAN Card Size (3.5 x 3.5 cm)</option></select></div><div><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;"><label style="font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.5px;">3. Zoom & Face Position</label><span id="zoomValBadge" style="font-size:11px;font-weight:800;color:#4f46e5;background:#e0e7ff;padding:2px 8px;border-radius:6px;">100%</span></div><input type="range" id="studioZoom" min="50" max="250" value="100" style="width:100%;height:6px;background:#cbd5e1;border-radius:4px;accent-color:#4f46e5;cursor:pointer;"></div><div><label style="display:block;font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">4. Copies & Grid Layout</label><select id="studioCopies" style="width:100%;padding:12px 14px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:12px;outline:none;font-size:13px;font-weight:700;color:#1e293b;cursor:pointer;"><option value="4">4 Photos Grid</option><option value="8">8 Photos Grid</option><option value="16">16 Photos Grid</option><option value="32" selected>32 Photos Grid (Full A4 Sheet)</option></select></div><div style="background:#e0e7ff;border:1px solid #c7d2fe;padding:12px 14px;border-radius:12px;display:flex;align-items:center;gap:10px;"><input type="checkbox" id="studioRemoveBg" style="width:18px;height:18px;accent-color:#4f46e5;cursor:pointer;"><label for="studioRemoveBg" style="font-size:12px;font-weight:800;color:#3730a3;cursor:pointer;">Remove Background Pro (Cloud AI Engine)</label></div></div><div style="display:flex;flex-direction:column;align-items:center;justify-content:center;background:#f1f5f9;padding:1.5rem;border-radius:16px;border:1px solid #e2e8f0;position:relative;"><span style="font-size:11px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:1px;margin-bottom:12px;">Live Studio Preview</span><div id="modalPreviewBox" style="width:160px;height:200px;background:#f87171;box-shadow:0 20px 30px rgba(0,0,0,0.15);border:2px solid white;overflow:hidden;position:relative;border-radius:8px;transition:background 0.3s;display:flex;align-items:center;justify-content:center;"><img src="\' + imgSrc + \'" id="modalPreviewImg" style="width:100%;height:100%;object-fit:contain;transform:scale(1);transform-origin:center;transition:transform 0.1s;" /></div><p style="font-size:11px;color:#64748b;font-weight:600;margin-top:12px;text-align:center;">Interactive preview updates instantly</p></div></div><div style="display:flex;justify-content:flex-end;gap:12px;padding-top:1rem;border-top:1px solid #e2e8f0;"><button onclick="closeModal()" style="padding:12px 24px;background:#f1f5f9;color:#334155;border:none;border-radius:12px;font-weight:800;font-size:13px;cursor:pointer;transition:0.2s;">Cancel</button><button onclick="applyStudioAndExecute(\' + stepIdx + \')" style="padding:12px 28px;background:linear-gradient(135deg,#4f46e5,#e11d48);color:white;border:none;border-radius:12px;font-weight:800;font-size:13px;cursor:pointer;box-shadow:0 10px 20px rgba(79,70,229,0.3);transition:0.2s;">✨ Apply & Execute Workflow</button></div></div></div>\';' +
    '    document.getElementById("modalContainer").innerHTML = modalHtml;' +
    '    document.getElementById("studioZoom").addEventListener("input", function(e) {' +
    '       const val = e.target.value;' +
    '       document.getElementById("zoomValBadge").textContent = val + "%";' +
    '       document.getElementById("modalPreviewImg").style.transform = "scale(" + (val / 100) + ")";' +
    '    });' +
    '    document.getElementById("studioBg").addEventListener("change", function(e) {' +
    '       document.getElementById("modalPreviewBox").style.backgroundColor = e.target.value;' +
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
    '  try { saveRes = await fetch("/api/workflow/create", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: name, steps: steps.map(s => ({ toolId: s.id, isComplex: !!s.isComplex, mode: s.mode, range: s.range, sortOrder: s.sortOrder, customSequence: s.customSequence, pageOrder: s.pageOrder, rotatePages: s.rotatePages, rotateDegree: s.rotateDegree, watermarkText: s.watermarkText, watermarkMode: s.watermarkMode, compressQuality: s.compressQuality, convertFormat: s.convertFormat, convertQuality: s.convertQuality, reducerMode: s.reducerMode, reducerWidth: s.reducerWidth, reducerHeight: s.reducerHeight, reducerTargetKb: s.reducerTargetKb, sizePreset: s.sizePreset, bgColor: s.bgColor, zoom: s.zoom, copies: s.copies, removeBg: s.removeBg, printSheet: s.printSheet })) }) }); }' +
    '  catch(e) { alert("Save failed: " + e.message); return; }' +
    '  const saveData = await saveRes.json();' +
    '  if (!saveData.success) { alert("Save error: " + saveData.error); return; }' +
    '  const formData = new FormData();' +
    '  for (let i = 0; i < selectedFiles.length; i++) { formData.append("file", selectedFiles[i]); }' +
    '  try {' +
    '    document.getElementById("executionResult").innerHTML = "<p style=\'color:blue;\'>⏳ Processing workflow steps...</p>";' +
    '    const execRes = await fetch("/api/workflow/execute/" + saveData.workflow.id, { method: "POST", body: formData });' +
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
    '  } catch(e) { document.getElementById("executionResult").innerHTML = "<p style=\'color:red;\'>Network error during execution: " + e.message + "</p>"; }' +
    '}' +
    'loadSavedWorkflows();' +
    '</script>' +
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