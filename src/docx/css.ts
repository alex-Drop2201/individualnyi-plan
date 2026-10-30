// Styles of the rendered document. Shared by the editor and the PDF pages (which are drawn
// outside the app's stylesheet), so they live here as text.
export const DOC_CSS = `
.page{box-sizing:border-box;background:#fff;color:#000;text-align:left;-webkit-text-size-adjust:none;text-size-adjust:none}
.body{display:flex;flex-direction:column}
.p{white-space:pre-wrap;overflow-wrap:break-word;tab-size:4;outline:none;flex:none}
.tbl{border-collapse:collapse;table-layout:fixed;align-self:flex-start;flex:none}
.tbl td{box-sizing:border-box;overflow-wrap:anywhere;padding:0}
.tbl td.gap{border:none}
.cell{display:flex;flex-direction:column}
.page img{display:inline-block;vertical-align:bottom}
`;
