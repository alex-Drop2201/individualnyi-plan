// PDF: cuts the document into pages where the editor shows them (src/pages.ts), draws each
// page into a picture and puts the pictures into a PDF.
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import type { WordDoc } from "./docx/document";
import { DOC_CSS } from "./docx/css";
import { planBreaks, PX } from "./pages";

const SCALE = 2.5; // ≈ 240 dpi on A4

export async function makePdf(doc: WordDoc): Promise<Blob> {
  const { width, height, top, bottom, left, right } = doc.section;
  const host = document.createElement("div");
  host.className = "pdf-host";
  document.body.append(host);
  try {
    // Lay the document out in one column at page width, find the page breaks the editor shows…
    const flow = document.createElement("div");
    flow.className = "page";
    Object.assign(flow.style, { width: `${width}pt`, padding: `0 ${right}pt 0 ${left}pt` });
    const column = document.createElement("div");
    column.className = "body";
    const blocks = doc.render(false);
    column.append(...blocks);
    flow.append(column);
    host.append(flow);
    const starts = new Set<HTMLElement>(planBreaks(blocks, (height - top - bottom) * PX).map((b) => b.at));

    // …then cut it into pages there.
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
    newPage();
    for (const block of blocks) {
      if (starts.has(block) && content.childElementCount) newPage();
      content.append(block);
      if (!(block instanceof HTMLTableElement)) continue;
      // Rows where a page starts go into a copy of the table on the next page.
      let tbody = block.tBodies[0];
      for (const row of Array.from(tbody.rows)) {
        if (!starts.has(row)) continue;
        newPage();
        const t = block.cloneNode(false) as HTMLTableElement;
        t.append(block.querySelector("colgroup")!.cloneNode(true), (tbody = document.createElement("tbody")));
        content.append(t);
        for (let r: Element | null = row; r; ) {
          const next: Element | null = r.nextElementSibling;
          if (starts.has(r as HTMLElement) && r !== row) break;
          tbody.append(r);
          r = next;
        }
      }
    }
    flow.remove();

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
