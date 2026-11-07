interface WorkbookEditorUi {
  readonly element: HTMLDivElement; readonly tools: HTMLDivElement; readonly input: HTMLTextAreaElement; readonly address: HTMLInputElement;
  readonly mode: HTMLSelectElement; readonly preview: HTMLDivElement; readonly status: HTMLParagraphElement; readonly error: HTMLParagraphElement;
}
export function createWorkbookEditorUi(container: HTMLElement, title: string): WorkbookEditorUi {
  const document = container.ownerDocument, element = document.createElement("div"), shadow = element.attachShadow({ mode: "closed" });
  const style = document.createElement("style"), heading = document.createElement("h2"), tools = document.createElement("div"), bar = document.createElement("div");
  const input = document.createElement("textarea"), address = document.createElement("input"), mode = document.createElement("select"), preview = document.createElement("div");
  const status = document.createElement("p"), error = document.createElement("p");
  style.textContent = ":host{display:block;color:#172638;font:14px/1.5 system-ui,sans-serif}*{box-sizing:border-box}h2{margin:0 0 8px;font-size:17px}.tools,.bar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:8px}button,input,textarea,select{font:inherit;border:1px solid #bdc9d5;border-radius:4px;background:#fff;color:inherit;padding:5px 8px}button{cursor:pointer}button:disabled{opacity:.5;cursor:default}textarea{flex:1;min-width:160px;resize:vertical;min-height:35px;max-height:160px}input{width:120px}p{margin:4px 0;white-space:pre-wrap}.error{color:#a12626}.preview{position:relative}";
  heading.textContent = title; tools.className = "tools"; tools.setAttribute("role", "toolbar"); tools.setAttribute("aria-label", "工作簿编辑"); bar.className = "bar";
  address.setAttribute("aria-label", "选择单元格或区域"); address.placeholder = "A1 或 A1:B10";
  input.setAttribute("aria-label", "单元格或公式输入"); input.rows = 1; input.spellcheck = false; input.disabled = true;
  mode.setAttribute("aria-label", "输入类型");
  for (const [value, label] of [["auto", "自动识别"], ["text", "文本"], ["number", "数字"]]) {
    const option = document.createElement("option"); option.value = value!; option.textContent = label!; mode.append(option);
  }
  status.setAttribute("role", "status"); status.setAttribute("aria-live", "polite"); status.textContent = "正在打开工作簿…";
  error.setAttribute("role", "alert"); error.className = "error"; error.hidden = true; preview.className = "preview";
  bar.append(address, input, mode); shadow.append(style, heading, tools, bar, status, error, preview); container.append(element);
  return { element, tools, input, address, mode, preview, status, error };
}
