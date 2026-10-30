// PDF: lays the document out on A4 (or the document's own) pages, draws each page into a
// picture and puts the pictures into a PDF. Breaks fall between paragraphs and between table
// rows (rows joined by a merged cell stay together), and at Word's own page breaks.
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import type { WordDoc } from "./docx/document";
import { DOC_CSS } from "./docx/css";

const PX = 4 / 3; // CSS px per pt
const SCALE = 2.5; // ≈ 240 dpi on A4

export async function makePdf(doc: WordDoc): Promise<Blob> {
  const { width, height, top, bottom, left, right } = doc.section;
  const host = document.createElement("div");
  host.className = "pdf-host";
  document.body.append(host);
  try {
    const limit = (height - top - bottom) * PX + 0.5;
    const pages: HTMLElement[] = [];
    let content!: HTMLElement;
    const newPage = () => {
      const page = document.createElement("div");
      page.className = "page";
      Object.assign(page.style, { width: `${width}pt`, height: `${height}pt`, padding: `${top}pt ${right}pt ${bottom}pt ${left}pt`, overflow: "hidden" });
      content = document.createElement("div");
      content.className = "body";
      page.append(content);
      host.append(page);
      pages.push(page);
    };
    const fits = () => content.offsetHeight <= limit;

    newPage();
    for (const block of doc.render(false)) {
      if (block.dataset.break === "before" && content.childElementCount) newPage();
      content.append(block);
      if (!fits() && content.childElementCount > 1) {
        if (block instanceof HTMLTableElement) splitTable(block);
        else {
          newPage();
          content.append(block);
        }
      }
      if (block.dataset.break === "after") newPage();
    }

    function splitTable(table: HTMLTableElement) {
      const rows = Array.from(table.tBodies[0].rows);
      const units: HTMLTableRowElement[][] = [];
      let reach = -1;
      rows.forEach((tr, i) => {
        if (i > reach) units.push([]);
        units[units.length - 1].push(tr);
        for (const td of Array.from(tr.cells)) reach = Math.max(reach, i + td.rowSpan - 1);
      });
      const shell = () => {
        const t = table.cloneNode(false) as HTMLTableElement;
        t.append(table.querySelector("colgroup")!.cloneNode(true), document.createElement("tbody"));
        content.append(t);
        return t.tBodies[0];
      };
      table.remove();
      let tbody = shell();
      for (const unit of units) {
        tbody.append(...unit);
        if (fits()) continue;
        const alone = tbody.rows.length === unit.length && content.childElementCount === 1;
        if (alone) continue; // taller than a page: nothing better to do
        unit.forEach((r) => r.remove());
        if (!tbody.rows.length) tbody.parentElement!.remove();
        newPage();
        tbody = shell();
        tbody.append(...unit);
      }
    }

    const pdf = new jsPDF({ unit: "pt", format: [width, height], orientation: width > height ? "landscape" : "portrait", compress: true });
    for (let i = 0; i < pages.length; i++) {
      if (i) pdf.addPage([width, height], width > height ? "landscape" : "portrait");
      const canvas = await draw(pages[i]);
      pdf.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", 0, 0, width, height);
      canvas.width = canvas.height = 0; // frees the memory right away on iOS
    }
    return pdf.output("blob");
  } finally {
    host.remove();
  }
}

/** Draws a page with the browser's own rendering when allowed, otherwise with html2canvas. */
async function draw(page: HTMLElement): Promise<HTMLCanvasElement> {
  try {
    const canvas = await drawNative(page);
    if (canvas) return canvas;
  } catch {
    // Not allowed here (Safari taints such pictures): fall back below.
  }
  return html2canvas(page, { scale: SCALE, backgroundColor: "#ffffff", logging: false });
}

async function drawNative(page: HTMLElement) {
  const w = page.offsetWidth;
  const h = page.offsetHeight;
  const html = new XMLSerializer().serializeToString(page);
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w * SCALE}" height="${h * SCALE}" viewBox="0 0 ${w} ${h}">` +
    `<foreignObject x="0" y="0" width="${w}" height="${h}"><div xmlns="http://www.w3.org/1999/xhtml"><style>${DOC_CSS}</style>${html}</div></foreignObject></svg>`;
  const img = new Image();
  img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
  await img.decode();
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(w * SCALE);
  canvas.height = Math.round(h * SCALE);
  const ctx = canvas.getContext("2d")!;
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  // Throws when the picture is not allowed to be read back; a blank result means it did not draw.
  const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
  const hasText = (page.textContent ?? "").trim() !== "";
  if (hasText) {
    for (let i = 0; i < data.length; i += 4 * 97) if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) return canvas;
    return null;
  }
  return canvas;
}
