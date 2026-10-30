/** Shares a file through the iOS share sheet, or downloads it where sharing files isn't supported. */
export async function shareFile(blob: Blob, name: string) {
  const file = new File([blob], name, { type: blob.type });
  if (navigator.canShare?.({ files: [file] }))
    try {
      await navigator.share({ files: [file] });
      return "shared";
    } catch (e) {
      if ((e as Error).name === "AbortError") return "cancelled";
    }
  downloadBlob(blob, name);
  return "downloaded";
}

export function canShareFiles() {
  return !!navigator.canShare?.({ files: [new File([""], "a.pdf", { type: "application/pdf" })] });
}

export function downloadBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
