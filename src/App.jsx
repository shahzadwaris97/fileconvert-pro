import html2canvas from "html2canvas";
import { useRef, useState } from "react";
import { PDFDocument, degrees, rgb, StandardFonts } from "pdf-lib";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import html2pdf from "html2pdf.js";
import JSZip from "jszip";
import * as pdfjsLib from "pdfjs-dist";
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import "./App.css";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

const tools = [
  { id: "merge", name: "Merge PDF", desc: "Combine PDFs in the order you choose.", icon: "▤", accept: ".pdf", multiple: true, category: "Organize PDF" },
  { id: "split", name: "Split PDF", desc: "Save every page as a separate PDF.", icon: "✂", accept: ".pdf", category: "Organize PDF" },
  { id: "extract", name: "Extract Pages", desc: "Export selected pages into a new PDF.", icon: "↗", accept: ".pdf", category: "Organize PDF" },
  { id: "delete", name: "Delete Pages", desc: "Remove selected pages from a PDF.", icon: "⌫", accept: ".pdf", category: "Organize PDF" },
  { id: "organize", name: "Reorder Pages", desc: "Reorder pages using a page-number sequence.", icon: "☷", accept: ".pdf", category: "Organize PDF" },
  { id: "rotate", name: "Rotate PDF", desc: "Rotate every page by 90, 180 or 270 degrees.", icon: "⟳", accept: ".pdf", category: "Organize PDF" },
  { id: "compress", name: "Compress PDF", desc: "Reduce size by rendering pages as optimized images.", icon: "⇩", accept: ".pdf", category: "Optimize PDF" },
  { id: "watermark", name: "Watermark", desc: "Stamp custom text on every page.", icon: "T", accept: ".pdf", category: "Edit PDF" },
  { id: "numbers", name: "Page Numbers", desc: "Add page numbers to the bottom of pages.", icon: "№", accept: ".pdf", category: "Edit PDF" },
  { id: "images-pdf", name: "JPG / PNG to PDF", desc: "Turn images into a single PDF.", icon: "▧", accept: "image/jpeg,image/png,image/webp", multiple: true, category: "Convert PDF" },
  { id: "pdf-images", name: "PDF to JPG", desc: "Export every PDF page as a JPG in a ZIP.", icon: "▣", accept: ".pdf", category: "Convert PDF" },
  { id: "word-pdf", name: "Word to PDF", desc: "Convert DOCX text and basic formatting to PDF.", icon: "W", accept: ".docx", category: "Convert PDF" },
  { id: "excel-pdf", name: "Excel to PDF", desc: "Convert spreadsheet sheets to PDF tables.", icon: "X", accept: ".xlsx,.xls,.csv", category: "Convert PDF" },
  { id: "pdf-text", name: "PDF to Text", desc: "Extract selectable text into a TXT file.", icon: "Tt", accept: ".pdf", category: "Convert PDF" },
  { id: "text-pdf", name: "Text to PDF", desc: "Turn a TXT file into a simple PDF.", icon: "¶", accept: ".txt", category: "Convert PDF" },
];

const categories = ["All Tools", "Workflows", "Organize PDF", "Optimize PDF", "Convert PDF", "Edit PDF", "PDF Security", "PDF Intelligence"];

const roadmapTools = [
  { id: "pdf-word", name: "PDF to Word", desc: "Convert PDF pages into an editable Word document.", icon: "W", accept: ".pdf", category: "Convert PDF", comingSoon: true },
  { id: "pdf-excel", name: "PDF to Excel", desc: "Extract tables from PDF files into spreadsheets.", icon: "X", accept: ".pdf", category: "Convert PDF", comingSoon: true },
  { id: "pdf-powerpoint", name: "PDF to PowerPoint", desc: "Create editable presentation slides from PDFs.", icon: "P", accept: ".pdf", category: "Convert PDF", comingSoon: true },
  { id: "powerpoint-pdf", name: "PowerPoint to PDF", desc: "Convert PPT and PPTX presentations into PDF.", icon: "P", accept: ".ppt,.pptx", category: "Convert PDF", comingSoon: true },
  { id: "html-pdf", name: "HTML to PDF", desc: "Create PDFs from web pages or HTML documents.", icon: "⌘", accept: ".html,.htm", category: "Convert PDF", comingSoon: true },
  { id: "edit-pdf", name: "Edit PDF", desc: "Add and edit text, images and shapes on PDF pages.", icon: "✎", accept: ".pdf", category: "Edit PDF", comingSoon: true },
  { id: "crop-pdf", name: "Crop PDF", desc: "Crop page margins and content areas.", icon: "□", accept: ".pdf", category: "Edit PDF", comingSoon: true },
  { id: "sign-pdf", name: "Sign PDF", desc: "Add a signature to your PDF document.", icon: "✍", accept: ".pdf", category: "PDF Security", comingSoon: true },
  { id: "protect-pdf", name: "Protect PDF", desc: "Encrypt a PDF with a password.", icon: "🔒", accept: ".pdf", category: "PDF Security", comingSoon: true },
  { id: "unlock-pdf", name: "Unlock PDF", desc: "Remove PDF restrictions when you have authorization.", icon: "🔓", accept: ".pdf", category: "PDF Security", comingSoon: true },
  { id: "redact-pdf", name: "Redact PDF", desc: "Permanently remove sensitive information from a PDF.", icon: "▰", accept: ".pdf", category: "PDF Security", comingSoon: true },
  { id: "repair-pdf", name: "Repair PDF", desc: "Attempt recovery of damaged PDF documents.", icon: "🛠", accept: ".pdf", category: "Optimize PDF", comingSoon: true },
  { id: "pdfa", name: "PDF to PDF/A", desc: "Prepare a PDF for long-term archival workflows.", icon: "A", accept: ".pdf", category: "Optimize PDF", comingSoon: true },
  { id: "ocr-pdf", name: "OCR PDF", desc: "Make scanned documents searchable with OCR.", icon: "Aa", accept: ".pdf", category: "PDF Intelligence", comingSoon: true },
  { id: "summarize-pdf", name: "AI Summarizer", desc: "Summarize document text into key points.", icon: "✦", accept: ".pdf", category: "PDF Intelligence", comingSoon: true },
  { id: "translate-pdf", name: "Translate PDF", desc: "Translate extracted document text into another language.", icon: "文", accept: ".pdf", category: "PDF Intelligence", comingSoon: true },
  { id: "compare-pdf", name: "Compare PDF", desc: "Compare two PDF files to identify differences.", icon: "⇄", accept: ".pdf", category: "PDF Intelligence", comingSoon: true },
  { id: "markdown-pdf", name: "PDF to Markdown", desc: "Extract document content into Markdown format.", icon: "M↓", accept: ".pdf", category: "PDF Intelligence", comingSoon: true },
];

const formatBytes = (bytes) => bytes < 1024 * 1024 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / 1024 / 1024).toFixed(2)} MB`;

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}
function baseName(filename = "document") {
  return filename.replace(/\.[^/.]+$/, "").replace(/[<>:"/\\|?*]/g, "_");
}
function parsePageRanges(input, pageCount) {
  const pages = new Set();
  for (const part of input.split(",")) {
    const value = part.trim();
    if (!value) continue;
    const range = value.match(/^(\d+)\s*-\s*(\d+)$/);
    if (range) {
      const start = Number(range[1]), end = Number(range[2]);
      if (start < 1 || end > pageCount || start > end) throw new Error(`Invalid page range "${value}". This PDF has ${pageCount} pages.`);
      for (let p = start; p <= end; p++) pages.add(p - 1);
    } else if (/^\d+$/.test(value)) {
      const p = Number(value);
      if (p < 1 || p > pageCount) throw new Error(`Page ${p} does not exist. This PDF has ${pageCount} pages.`);
      pages.add(p - 1);
    } else throw new Error(`Invalid page range "${value}". Example: 1-3,5`);
  }
  return [...pages].sort((a, b) => a - b);
}
async function readPdf(file) {
  return PDFDocument.load(await file.arrayBuffer());
}
async function pdfBlob(doc) {
  return new Blob([await doc.save()], { type: "application/pdf" });
}
async function renderPdfPage(page, scale = 1.5) {
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Could not create a canvas.");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  await page.render({ canvasContext: context, viewport }).promise;
  return { canvas, viewport };
}

export default function App() {
  const [activeCategory, setActiveCategory] = useState("All Tools");
  const [activeTool, setActiveTool] = useState(null);
  const [files, setFiles] = useState([]);
  const [pageRanges, setPageRanges] = useState("1-2");
  const [pageOrder, setPageOrder] = useState("1,2,3");
  const [watermarkText, setWatermarkText] = useState("YOURLOVEPDF");
  const [rotation, setRotation] = useState("90");
  const [compressionLevel, setCompressionLevel] = useState("recommended");
  const [search, setSearch] = useState("");
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [splitResults, setSplitResults] = useState([]);
  const [splitZipBlob, setSplitZipBlob] = useState(null);
  const [freeConversions, setFreeConversions] = useState(() => {
    try { return Math.min(20, Math.max(0, Number(localStorage.getItem("yourlovepdf-free-demo-count") || 0))); }
    catch { return 0; }
  });
  const fileInput = useRef(null);

  function recordSuccessfulConversion() {
    setFreeConversions((previous) => {
      const next = Math.min(20, previous + 1);
      try { localStorage.setItem("yourlovepdf-free-demo-count", String(next)); } catch { /* storage may be unavailable */ }
      return next;
    });
  }

  const visibleTools = [...tools, ...roadmapTools].filter((tool) =>
    (activeCategory === "All Tools" || activeCategory === "Workflows" || tool.category === activeCategory) &&
    `${tool.name} ${tool.desc}`.toLowerCase().includes(search.toLowerCase())
  );

  function resetWorkspace() {
    setActiveTool(null); setFiles([]); setStatus(""); setError("");
    setSplitResults([]); setSplitZipBlob(null);
    if (fileInput.current) fileInput.current.value = "";
  }
  function chooseTool(tool) {
    if (tool.comingSoon) {
      setActiveTool(null); setFiles([]); setError("");
      setStatus(`${tool.name} is listed in the roadmap but is not enabled yet. We will activate it after its conversion workflow is implemented and tested.`);
      return;
    }
    setActiveTool(tool); setFiles([]); setPageRanges("1-2"); setPageOrder("1,2,3");
    setWatermarkText("YOURLOVEPDF"); setRotation("90"); setCompressionLevel("recommended");
    setStatus(""); setError(""); setSplitResults([]); setSplitZipBlob(null);
    if (fileInput.current) fileInput.current.value = "";
    document.getElementById("tools")?.scrollIntoView({ behavior: "smooth", block: "start" });
  }
  function handleFiles(list) {
    const chosen = Array.from(list || []);
    if (!activeTool) return;
    const accepted = chosen.filter((file) => {
      if (activeTool.id === "images-pdf") return /^image\/(jpeg|png|webp)$/i.test(file.type) || /\.(jpe?g|png|webp)$/i.test(file.name);
      return file.name.toLowerCase().endsWith(activeTool.accept.split(",")[0].replace(".", "")) ||
        activeTool.accept.split(",").some((ext) => ext.startsWith(".") && file.name.toLowerCase().endsWith(ext));
    });
    if (chosen.length && !accepted.length) {
      setError(`Please choose a supported file type: ${activeTool.accept}`);
      return;
    }
    setFiles(accepted); setStatus(""); setError(""); setSplitResults([]); setSplitZipBlob(null);
  }

  async function mergePdfs() {
    if (files.length < 2) throw new Error("Choose at least two PDF files to merge.");
    const output = await PDFDocument.create();
    for (const file of files) {
      const input = await readPdf(file);
      const pages = await output.copyPages(input, input.getPageIndices());
      pages.forEach((page) => output.addPage(page));
    }
    downloadBlob(await pdfBlob(output), "yourlovepdf-merged.pdf");
  }
  async function splitPdf() {
    const source = await readPdf(files[0]), zip = new JSZip(), results = [];
    if (!source.getPageCount()) throw new Error("This PDF has no pages.");
    for (let i = 0; i < source.getPageCount(); i++) {
      setStatus(`Splitting page ${i + 1} of ${source.getPageCount()}…`);
      const one = await PDFDocument.create();
      const [page] = await one.copyPages(source, [i]); one.addPage(page);
      const bytes = await one.save(), name = `${baseName(files[0].name)}-page-${i + 1}.pdf`;
      const blob = new Blob([bytes], { type: "application/pdf" });
      results.push({ name, blob }); zip.file(name, bytes);
    }
    setSplitResults(results); setSplitZipBlob(await zip.generateAsync({ type: "blob" }));
    setStatus(`Done — ${results.length} PDF pages created.`);
  }
  async function extractPages() {
    const src = await readPdf(files[0]), indexes = parsePageRanges(pageRanges, src.getPageCount());
    if (!indexes.length) throw new Error("Enter page numbers, for example 1-3,5.");
    const out = await PDFDocument.create(), pages = await out.copyPages(src, indexes);
    pages.forEach((p) => out.addPage(p));
    downloadBlob(await pdfBlob(out), `${baseName(files[0].name)}-extracted.pdf`);
  }
  async function deletePages() {
    const src = await readPdf(files[0]), indexes = parsePageRanges(pageRanges, src.getPageCount());
    if (!indexes.length) throw new Error("Enter the page numbers to delete.");
    if (indexes.length >= src.getPageCount()) throw new Error("You must keep at least one page.");
    indexes.reverse().forEach((index) => src.removePage(index));
    downloadBlob(await pdfBlob(src), `${baseName(files[0].name)}-edited.pdf`);
  }
  async function reorderPages() {
    const src = await readPdf(files[0]);
    const indexes = pageOrder.split(",").map((v) => {
      const n = Number(v.trim());
      if (!Number.isInteger(n) || n < 1 || n > src.getPageCount()) throw new Error(`Invalid page order. Use page numbers from 1 to ${src.getPageCount()}, e.g. 3,1,2.`);
      return n - 1;
    });
    if (!indexes.length) throw new Error("Enter a page order such as 3,1,2.");
    const out = await PDFDocument.create(), pages = await out.copyPages(src, indexes);
    pages.forEach((p) => out.addPage(p));
    downloadBlob(await pdfBlob(out), `${baseName(files[0].name)}-reordered.pdf`);
  }
  async function rotatePdf() {
    const pdf = await readPdf(files[0]), amount = Number(rotation);
    pdf.getPages().forEach((page) => page.setRotation(degrees((page.getRotation().angle + amount) % 360)));
    downloadBlob(await pdfBlob(pdf), `${baseName(files[0].name)}-rotated.pdf`);
  }
  async function addWatermark() {
    if (!watermarkText.trim()) throw new Error("Enter watermark text.");
    const pdf = await readPdf(files[0]), font = await pdf.embedFont(StandardFonts.HelveticaBold);
    pdf.getPages().forEach((page) => {
      const { width, height } = page.getSize();
      page.drawText(watermarkText.trim(), { x: width * 0.1, y: height * 0.48, size: Math.min(width / 12, 38), font, color: rgb(0.65, 0.68, 0.73), opacity: 0.28, rotate: degrees(30) });
    });
    downloadBlob(await pdfBlob(pdf), `${baseName(files[0].name)}-watermarked.pdf`);
  }
  async function addPageNumbers() {
    const pdf = await readPdf(files[0]), font = await pdf.embedFont(StandardFonts.Helvetica);
    pdf.getPages().forEach((page, i) => {
      const { width } = page.getSize(), label = String(i + 1), size = 11;
      page.drawText(label, { x: (width - font.widthOfTextAtSize(label, size)) / 2, y: 18, size, font, color: rgb(0.25, 0.25, 0.25) });
    });
    downloadBlob(await pdfBlob(pdf), `${baseName(files[0].name)}-numbered.pdf`);
  }
  async function imagesToPdf() {
    if (!files.length) throw new Error("Choose one or more images.");
    const out = await PDFDocument.create();
    for (const file of files) {
      const bytes = await file.arrayBuffer(), type = file.type.toLowerCase();
      let image;
      if (type === "image/jpeg" || /\.jpe?g$/i.test(file.name)) image = await out.embedJpg(bytes);
      else if (type === "image/png" || /\.png$/i.test(file.name)) image = await out.embedPng(bytes);
      else if (type === "image/webp" || /\.webp$/i.test(file.name)) {
        const bitmap = await createImageBitmap(file), canvas = document.createElement("canvas");
        canvas.width = bitmap.width; canvas.height = bitmap.height;
        const ctx = canvas.getContext("2d"); if (!ctx) throw new Error("Could not process image.");
        ctx.drawImage(bitmap, 0, 0);
        const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
        bitmap.close(); if (!blob) throw new Error(`Could not process ${file.name}.`);
        image = await out.embedPng(await blob.arrayBuffer());
      } else throw new Error(`Unsupported image: ${file.name}`);
      const page = out.addPage([image.width, image.height]);
      page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
    }
    downloadBlob(await pdfBlob(out), "yourlovepdf-images.pdf");
  }
  async function pdfToImages() {
    const task = pdfjsLib.getDocument({ data: new Uint8Array(await files[0].arrayBuffer()) });
    const doc = await task.promise, zip = new JSZip();
    for (let n = 1; n <= doc.numPages; n++) {
      setStatus(`Rendering page ${n} of ${doc.numPages}…`);
      const page = await doc.getPage(n), { canvas } = await renderPdfPage(page, 1.6);
      const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", 0.9));
      if (!blob) throw new Error(`Could not render page ${n}.`);
      zip.file(`${baseName(files[0].name)}-page-${n}.jpg`, await blob.arrayBuffer());
      canvas.width = 0; canvas.height = 0;
    }
    downloadBlob(await zip.generateAsync({ type: "blob" }), `${baseName(files[0].name)}-jpg.zip`);
  } 
 
  
async function wordToPdf() {
  const file = files[0];

  if (!file || !/\.docx$/i.test(file.name)) {
    throw new Error("Please select a valid Word .docx file.");
  }

  let container;
  let canvas;

  try {
    setStatus("Reading Word document...");

    const result = await mammoth.convertToHtml(
      { arrayBuffer: await file.arrayBuffer() },
      {
        convertImage: mammoth.images.imgElement(async (image) => {
          const base64 = await image.read("base64");
          return {
            src: `data:${image.contentType};base64,${base64}`
          };
        })
      }
    );

    if (!result.value || !result.value.trim()) {
      throw new Error("No readable text found in this Word document.");
    }

    container = document.createElement("div");
    container.className = "word-export";

    Object.assign(container.style, {
      position: "absolute",
      left: "0",
      top: "0",
      width: "700px",
      boxSizing: "border-box",
      padding: "12px",
      margin: "0",
      backgroundColor: "#ffffff",
      color: "#111111",
      fontFamily: "Arial, sans-serif",
      fontSize: "11pt",
      lineHeight: "1.4",
      visibility: "visible",
      opacity: "1",
      zIndex: "99999"
    });

    container.innerHTML = `
      <style>
        .word-export,
        .word-export * {
          box-sizing: border-box;
        }

        .word-export {
          background: #ffffff !important;
          color: #111111 !important;
          font-family: Arial, sans-serif !important;
          font-size: 11pt !important;
          line-height: 1.4 !important;
          overflow-wrap: anywhere;
        }

        .word-export p {
          margin: 0 0 10pt;
        }

        .word-export table {
          width: 100%;
          border-collapse: collapse;
          margin: 8pt 0;
        }

        .word-export td,
        .word-export th {
          border: 1px solid #888888;
          padding: 5px;
          vertical-align: top;
        }

        .word-export img {
          max-width: 100%;
          height: auto;
        }

        .word-export h1,
        .word-export h2,
        .word-export h3 {
          page-break-after: avoid;
        }
      </style>
      ${result.value}
    `;

    document.body.appendChild(container);

    setStatus("Preparing Word document...");

    if (document.fonts?.ready) {
      await document.fonts.ready;
    }

    await Promise.all(
      Array.from(container.querySelectorAll("img")).map(async (img) => {
        if (img.complete) return;

        try {
          await img.decode();
        } catch {
          // Continue if an image cannot be decoded.
        }
      })
    );

    await new Promise((resolve) =>
      requestAnimationFrame(() =>
        requestAnimationFrame(resolve)
      )
    );

    setStatus("Rendering Word pages...");

    canvas = await html2canvas(container, {
      scale: 1.5,
      backgroundColor: "#ffffff",
      useCORS: true,
      logging: false,
      windowWidth: 900,
      scrollX: 0,
      scrollY: 0
    });

    if (!canvas.width || !canvas.height) {
      throw new Error(
        "Word document could not be rendered. Please try another DOCX file."
      );
    }

    const pdf = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4",
      compress: true
    });

    const margin = 12;
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const contentWidth = pageWidth - margin * 2;
    const contentHeight = pageHeight - margin * 2;

    const pixelsPerMm = canvas.width / contentWidth;
    const sourcePageHeight = Math.max(
      1,
      Math.floor(contentHeight * pixelsPerMm)
    );

    let pageNumber = 0;

    for (
      let sourceY = 0;
      sourceY < canvas.height;
      sourceY += sourcePageHeight
    ) {
      setStatus(`Creating PDF page ${pageNumber + 1}...`);

      const sliceHeight = Math.min(
        sourcePageHeight,
        canvas.height - sourceY
      );

      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = canvas.width;
      pageCanvas.height = sliceHeight;

      const context = pageCanvas.getContext("2d");

      if (!context) {
        throw new Error("Could not prepare a PDF page.");
      }

      context.fillStyle = "#ffffff";
      context.fillRect(0, 0, pageCanvas.width, pageCanvas.height);

      context.drawImage(
        canvas,
        0,
        sourceY,
        canvas.width,
        sliceHeight,
        0,
        0,
        canvas.width,
        sliceHeight
      );

      if (pageNumber > 0) {
        pdf.addPage("a4", "portrait");
      }

      const imageData = pageCanvas.toDataURL("image/jpeg", 0.95);
      const imageHeight = sliceHeight / pixelsPerMm;

      pdf.addImage(
        imageData,
        "JPEG",
        margin,
        margin,
        contentWidth,
        imageHeight,
        undefined,
        "FAST"
      );

      pageCanvas.width = 0;
      pageCanvas.height = 0;
      pageNumber++;
    }

    const blob = pdf.output("blob");

    if (!blob || blob.size < 1000 || pageNumber < 1) {
      throw new Error(
        "The generated PDF appears empty. Please try another DOCX file."
      );
    }

    downloadBlob(blob, `${baseName(file.name)}.pdf`);

    setStatus(
      `Conversion complete: ${pageNumber} PDF page(s) created. Please check the output formatting.`
    );
  } catch (err) {
    console.error("Word to PDF conversion error:", err);
    throw err;
  } finally {
    if (canvas) {
      canvas.width = 0;
      canvas.height = 0;
    }

    container?.remove();
  }
}

async function excelToPdf() {
  const file = files[0];

  if (!file) {
    throw new Error("Please select an Excel file first.");
  }

  const workbook = XLSX.read(
    await file.arrayBuffer(),
    { type: "array", cellDates: true }
  );

  const pdf = new jsPDF({
    orientation: "landscape",
    unit: "mm",
    format: "a4",
    compress: true
  });

  let hasContent = false;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];

    if (!sheet || !sheet["!ref"]) continue;

    setStatus(`Preparing sheet: ${sheetName}...`);

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const rows = [];

    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];

      for (let c = range.s.c; c <= range.e.c; c++) {
        const address = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[address];

        let value = "";

        if (cell) {
          if (cell.w !== undefined) {
            value = cell.w;
          } else if (cell.v !== undefined) {
            value = String(cell.v);
          }
        }

        row.push(String(value ?? ""));
      }

      rows.push(row);
    }

    if (!rows.some((row) => row.some((cell) => cell.trim()))) {
      continue;
    }

    if (hasContent) {
      pdf.addPage("a4", "landscape");
    }

    hasContent = true;

    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(13);
    pdf.setTextColor(30, 70, 50);
    pdf.text(String(sheetName).slice(0, 80), 9, 10);

    const usableWidth =
      pdf.internal.pageSize.getWidth() - 18;

    const columnCount = Math.max(
      1,
      ...rows.map((row) => row.length)
    );

    // Use Excel column width information when available.
    const excelColumns = sheet["!cols"] || [];

    const weights = Array.from(
      { length: columnCount },
      (_, index) => {
        const excelWidth = excelColumns[index]?.wch;

        return Math.max(
          1,
          Math.min(Number(excelWidth) || 10, 35)
        );
      }
    );

    const totalWeight = weights.reduce(
      (sum, weight) => sum + weight,
      0
    );

    const columnStyles = {};

    weights.forEach((weight, index) => {
      columnStyles[index] = {
        cellWidth: (weight / totalWeight) * usableWidth
      };
    });

    const normalized = rows.map((row) =>
      Array.from(
        { length: columnCount },
        (_, index) => String(row[index] ?? "")
      )
    );

    autoTable(pdf, {
      head: [normalized[0]],
      body: normalized.slice(1),
      startY: 15,
      margin: { top: 15, right: 9, bottom: 10, left: 9 },
      tableWidth: "auto",
      columnStyles,
      theme: "grid",
      showHead: "everyPage",
      styles: {
        font: "helvetica",
        fontSize: 7,
        cellPadding: 2,
        overflow: "linebreak",
        lineColor: [160, 170, 165],
        lineWidth: 0.2,
        textColor: [35, 35, 35],
        fillColor: [255, 255, 255],
        valign: "middle"
      },
      headStyles: {
        fillColor: [220, 240, 226],
        textColor: [20, 65, 40],
        fontStyle: "bold",
        lineColor: [100, 135, 110],
        lineWidth: 0.25
      },
      alternateRowStyles: {
        fillColor: [247, 250, 247]
      },
      horizontalPageBreak: true,
      horizontalPageBreakRepeat: 0
    });
  }

  if (!hasContent) {
    throw new Error("No spreadsheet data found.");
  }

  downloadBlob(
    pdf.output("blob"),
    `${baseName(file.name)}.pdf`
  );

  setStatus(
    "Excel conversion complete. Check the PDF layout and gridlines."
  );
}
  async function runTool() {
    if (!activeTool) return;
    if (!files.length) { setError("Please choose a file first."); return; }
    if (!activeTool.multiple && files.length !== 1) { setError("Please select one file for this tool."); return; }
    setBusy(true); setError(""); setStatus("Processing your file…"); setSplitResults([]); setSplitZipBlob(null);
    try {
      switch (activeTool.id) {
        case "merge": await mergePdfs(); break;
        case "split": await splitPdf(); break;
        case "extract": await extractPages(); break;
        case "delete": await deletePages(); break;
        case "organize": await reorderPages(); break;
        case "rotate": await rotatePdf(); break;
        case "watermark": await addWatermark(); break;
        case "numbers": await addPageNumbers(); break;
        case "images-pdf": await imagesToPdf(); break;
        case "pdf-images": await pdfToImages(); break;
        case "word-pdf": await wordToPdf(); break;
        case "excel-pdf": await excelToPdf(); break;
        case "pdf-text": await pdfToText(); break;
        case "text-pdf": await textToPdf(); break;
        case "compress": await compressPdf(); break;
        default: throw new Error("This tool is not available yet.");
      }
      recordSuccessfulConversion();
      if (!["split", "compress", "pdf-images", "pdf-text"].includes(activeTool.id)) setStatus("Done! Your download should start automatically.");
      else if (!status.startsWith("Done") && activeTool.id === "split") setStatus("Split complete. Download your pages below.");
    } catch (err) {
      console.error(err); setStatus(""); setError(err?.message || "Something went wrong. Try another file or a smaller document.");
    } finally { setBusy(false); }
  }

  return (
    <div className="app-shell" id="home">
      <header className="topbar">
        <a className="brand" href="#home" onClick={resetWorkspace}><span className="brand-icon">💙</span><span><strong>YourLovePDF</strong><small>Simple. Secure. Made for you.</small></span></a>
        <nav className="top-links" aria-label="Main navigation"><a className="nav-tab" href="#home">Home</a><a className="nav-tab" href="#tools">All Tools</a><a className="nav-tab" href="#pricing">Premium</a><a className="nav-tab" href="#business">Business</a><a className="nav-tab" href="#how-it-works">How it works</a><a className="nav-tab" href="#faq">FAQ</a><a className="nav-cta" href="#signup">Sign up free <span>↗</span></a></nav>
      </header>
      {!activeTool && status && <div className="global-notice" role="status">{status}</div>}

      <main>
        <section className="hero">
          <div className="hero-copy">
            <div className="eyebrow"><span className="live-dot" /> YOUR EVERYDAY PDF TOOLKIT · 💙</div>
            <h1>Every PDF tool.<br /><span>One simple place.</span></h1>
            <p>Merge, split, compress, convert and organize your documents with easy-to-use tools. Start with 20 free successful conversions.</p>
            <div className="hero-actions"><a className="primary-button" href="#tools">Explore all tools <span>→</span></a><a className="hero-signup-button" href="#signup">✦ Sign up free</a><span className="privacy-note">🔒 Files stay on your device for supported tools</span></div>
            <div className="hero-stats"><div><strong>{tools.length + roadmapTools.length}+</strong><span>Tools listed</span></div><div><strong>{Math.max(0, 20 - freeConversions)}/20</strong><span>Free conversions left on this browser</span></div><div><strong>💙</strong><span>Simple file tools</span></div></div>
          </div>
          <div className="hero-art" aria-label="PDF conversion illustration"><div className="art-glow" /><div className="orbit orbit-one" /><div className="orbit orbit-two" /><div className="file-card file-pdf"><div className="file-symbol red">PDF</div><div className="file-lines"><i /><i /><i /></div></div><div className="conversion-arrow">↗</div><div className="file-card file-doc"><div className="file-symbol blue">DOC</div><div className="file-lines"><i /><i /><i /></div></div><div className="floating-badge badge-top"><span>✓</span> Easy conversion</div><div className="floating-badge badge-bottom"><span>⌁</span> All in one place</div></div>
        </section>

        <section className="trust-strip"><div><span className="trust-icon">🔐</span><span><strong>Privacy-minded</strong><small>Browser-based processing for supported tools</small></span></div><div><span className="trust-icon">⚡</span><span><strong>Simple workflow</strong><small>Upload, process and download</small></span></div><div><span className="trust-icon">✦</span><span><strong>Useful file tools</strong><small>Everyday PDF and document tasks</small></span></div></section>

        <section className="tools-section" id="tools">
          <div className="section-heading"><div><span className="section-kicker">THE TOOLBOX</span><h2>What do you need to do?</h2><p>Choose a tool to get started with your files.</p></div><label className="search-box"><span>⌕</span><input type="search" placeholder="Search tools..." value={search} onChange={(e) => setSearch(e.target.value)} aria-label="Search tools" /><kbd>⌕</kbd></label></div>
          <div className="category-tabs">{categories.map((category) => <button type="button" key={category} className={activeCategory === category ? "active" : ""} onClick={() => setActiveCategory(category)}>{category}</button>)}</div>

          {activeTool && <section className="converter-panel" aria-live="polite">
            <div className="converter-header"><div className="converter-title"><div className="tool-icon selected-icon">{activeTool.icon}</div><div><span className="section-kicker">FILE WORKSPACE</span><h3>{activeTool.name}</h3><p>{activeTool.desc}</p></div></div><button type="button" className="close-tool" onClick={resetWorkspace} aria-label="Close tool">✕</button></div>
            <input ref={fileInput} className="hidden-input" type="file" accept={activeTool.accept} multiple={Boolean(activeTool.multiple)} onChange={(e) => handleFiles(e.target.files)} />
            <button type="button" className="drop-zone" onClick={() => fileInput.current?.click()} onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.preventDefault(); handleFiles(e.dataTransfer.files); }}><span className="upload-icon">↑</span><strong>Click to upload or drop files here</strong><span>{activeTool.multiple ? "Select multiple files if needed" : "Select one file to continue"}</span><small>Supported: {activeTool.accept}</small></button>
            {files.length > 0 && <div className="selected-files"><strong>Selected files ({files.length})</strong>{files.map((file, index) => <div className="file-row" key={`${file.name}-${index}`}><span className="mini-file-icon">▤</span><span className="file-name">{file.name}</span><span className="file-size">{formatBytes(file.size)}</span><button type="button" onClick={() => { setFiles((old) => old.filter((_, i) => i !== index)); setStatus(""); setError(""); }} aria-label={`Remove ${file.name}`}>×</button></div>)}</div>}
            {["extract", "delete"].includes(activeTool.id) && <label className="field-label">Page numbers or ranges<input value={pageRanges} onChange={(e) => setPageRanges(e.target.value)} placeholder="Example: 1-3,5" /><small>Separate pages with commas; use a hyphen for ranges.</small></label>}
            {activeTool.id === "organize" && <label className="field-label">New page order<input value={pageOrder} onChange={(e) => setPageOrder(e.target.value)} placeholder="Example: 3,1,2" /><small>Enter every page number in the order you want. Repeat a number to duplicate a page; omit a number to leave that page out.</small></label>}
            {activeTool.id === "rotate" && <label className="field-label">Rotation<select value={rotation} onChange={(e) => setRotation(e.target.value)}><option value="90">90° clockwise</option><option value="180">180°</option><option value="270">270° clockwise</option></select></label>}
            {activeTool.id === "watermark" && <label className="field-label">Watermark text<input value={watermarkText} onChange={(e) => setWatermarkText(e.target.value)} maxLength={80} placeholder="Enter watermark" /></label>}
            {activeTool.id === "compress" && <label className="field-label">Compression level<select value={compressionLevel} onChange={(e) => setCompressionLevel(e.target.value)}><option value="recommended">Recommended — balanced quality</option><option value="strong">Strong — smaller file</option><option value="extreme">Extreme — lowest image quality</option></select><small>Compression converts pages to images, so text may no longer be selectable. Results depend on the original PDF.</small></label>}
            {error && <div className="message error-message">{error}</div>}{status && <div className="message status-message">{busy && <span className="spinner" />}{status}</div>}
            {activeTool.id === "split" && splitResults.length > 0 && <div className="split-results"><h3>Split PDF files</h3><p>{splitResults.length} separate PDFs created.</p>{splitResults.map((item) => <div className="file-row" key={item.name}><span className="mini-file-icon">PDF</span><span className="file-name">{item.name}</span><button type="button" onClick={() => downloadBlob(item.blob, item.name)}>Download PDF</button></div>)}{splitZipBlob && <button type="button" className="primary-button" onClick={() => downloadBlob(splitZipBlob, `${baseName(files[0]?.name)}-split.zip`)}>Download all as ZIP ↓</button>}</div>}
            <button type="button" className="primary-button process-button" disabled={busy || files.length === 0} onClick={runTool}>{busy ? "Processing…" : `Process ${activeTool.name}`} {!busy && <span>→</span>}</button>
            <p className="local-note">Files are processed in your browser for these tools. Avoid using shared or untrusted devices for sensitive documents.</p>
          </section>}

          <div className="tool-grid">{visibleTools.map((tool) => <button type="button" className={`tool-card ${activeTool?.id === tool.id ? "chosen" : ""} ${tool.comingSoon ? "tool-coming-soon" : ""}`} key={tool.id} onClick={() => chooseTool(tool)}><span className={`tool-icon icon-${tool.id}`}>{tool.icon}</span><span className="tool-card-title">{tool.name}</span><span className="tool-card-desc">{tool.desc}</span><span className="tool-card-arrow">{tool.comingSoon ? "In development" : "Open tool ↗"}</span></button>)}</div>
          {visibleTools.length === 0 && <div className="empty-state">No tools match your search. Try another keyword.</div>}
        </section>

        <section className="pricing-section" id="pricing">
          <div className="section-heading centered"><span className="section-kicker">SIMPLE PLANS</span><h2>Start free. Upgrade when you need more.</h2><p>Every account is planned to include 20 successful free conversions. Paid checkout activates after secure backend and merchant setup.</p></div>
          <div className="pricing-grid">
            <article className="pricing-card" id="signup"><span className="plan-label">FREE ACCOUNT</span><h3>Sign up and get started</h3><p className="plan-price">Rs 0 <small>/ forever</small></p><ul><li>20 successful conversions</li><li>Essential browser-based tools</li><li>No payment required</li></ul><a className="pricing-button secondary signup-button" href="#tools">Start free — no account needed yet <span>→</span></a><small className="signup-note">Account registration is being prepared. You can use available tools now; secure sign-up will be enabled with the backend.</small></article>
            <article className="pricing-card featured"><span className="plan-label">PREMIUM</span><h3>Premium Monthly</h3><p className="plan-price">Price to be set <small>/ month</small></p><ul><li>Higher conversion allowance</li><li>Premium tools as they launch</li><li>Priority feature access</li></ul><button className="pricing-button" type="button" onClick={() => setStatus("Premium checkout is not active yet. Payment gateway and server-side verification must be configured first.")}>Premium coming soon</button></article>
            <article className="pricing-card"><span className="plan-label">CREDITS</span><h3>Conversion Credits</h3><p className="plan-price">Pay as you go</p><ul><li>Buy additional conversions</li><li>Use credits when free quota ends</li><li>Payment receipt and balance tracking planned</li></ul><button className="pricing-button secondary" type="button" onClick={() => setStatus("Credit purchases will be enabled after payment gateway integration.")}>Credits coming soon</button></article>
          </div>
          <p className="plan-note">Development note: this preview tracks a demo count in this browser only. It is not yet an account-level limit and can be reset. The server/database stage will enforce the real 20-conversion allowance.</p>
        </section>

        <section className="faq-section" id="faq"><div className="section-heading centered"><span className="section-kicker">HELP & SUPPORT</span><h2>Frequently asked questions</h2></div><div className="faq-grid"><details><summary>Are the tools free?</summary><p>The planned free tier includes 20 successful conversions per account. Paid plans will be enabled after backend and payment setup.</p></details><details><summary>Are my files uploaded?</summary><p>The current supported tools process files in your browser. Review the privacy note shown on each tool; future server-powered tools may handle files differently.</p></details><details><summary>Why are some tools marked as coming soon?</summary><p>We list the wider tool roadmap but only the existing implemented tools are active. Each additional tool must be coded and tested before it is enabled.</p></details><details><summary>Can I pay online now?</summary><p>Not yet. A merchant account, secure backend and verified payment callbacks must be configured before real payments can be accepted.</p></details></div></section>

        <section className="business-strip" id="business"><div><span className="section-kicker">FOR TEAMS & BUSINESSES</span><h2>PDF workflows for work, study and everyday tasks.</h2><p>YourLovePDF is being built for simple personal tools first, with account-based usage and team features planned for a later release.</p></div><a className="primary-button" href="#pricing">View plans <span>→</span></a></section>

        <section className="how-section" id="how-it-works"><div className="section-heading centered"><span className="section-kicker">THREE SIMPLE STEPS</span><h2>From file to finished.</h2><p>No complicated workflow. Just select, process and save.</p></div><div className="steps-grid"><div className="step-card"><span className="step-number">01</span><div className="step-icon">↑</div><h3>Choose your file</h3><p>Select a supported file from your device.</p></div><div className="step-card"><span className="step-number">02</span><div className="step-icon">⚙</div><h3>Run your tool</h3><p>Pick your task and process the file in your browser.</p></div><div className="step-card"><span className="step-number">03</span><div className="step-icon">↓</div><h3>Save your result</h3><p>Download the new file when processing finishes.</p></div></div></section>

        <section className="privacy-banner" id="security"><div className="privacy-shield">🔒</div><div><h2>Your files stay in your hands.</h2><p>Supported tools process files locally in your browser. This site does not upload these files to a YourLovePDF server. Conversion results can differ by file type and formatting; always review important documents.</p></div><a href="#tools">Try a tool <span>→</span></a></section>
      </main>
      <footer className="footer" id="contact"><div className="footer-main"><div className="footer-brand"><a className="brand" href="#home" onClick={resetWorkspace}><span className="brand-icon">💙</span><span><strong>YourLovePDF</strong><small>Simple. Secure. Made for you.</small></span></a><p>Helpful PDF and file tools in one simple place.</p><span className="made-pakistan">Built with care in Pakistan 🇵🇰</span></div><div className="footer-column"><h4>Product</h4><a href="#home">Home</a><a href="#tools">All Tools</a><a href="#pricing">Pricing</a><a href="#business">Business</a></div><div className="footer-column"><h4>Support</h4><a href="#faq">FAQ</a><a href="#contact">Contact</a><a href="#security">Security</a></div><div className="footer-column"><h4>Legal</h4><a href="#privacy">Privacy policy</a><a href="#terms">Terms &amp; conditions</a><a href="#cookies">Cookies</a></div><div className="footer-contact"><h3>Contact Us</h3><p><strong>Email:</strong> <a href="mailto:warisshahzad743@gmail.com">warisshahzad743@gmail.com</a></p><p><strong>Address:</strong> Syedan Colony, Dharema, Sargodha, Pakistan</p><p><strong>Phone:</strong> <a href="tel:+923091081418">+92 309 1081418</a></p></div></div><div className="footer-bottom" id="privacy"><span>© {new Date().getFullYear()} YourLovePDF 💙. All rights reserved.</span><span id="terms">Terms &amp; conditions and cookie policy should be published before launch. <a href="#cookies" id="cookies">Cookie information</a></span></div></footer>
    </div>
  );
}
