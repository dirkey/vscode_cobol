import * as vscode from "vscode";
import * as fs from "fs";
import { ICOBOLSettings } from "./iconfiguration";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import {
  COBOLToken,
  COBOLTokenStyle,
  ParseState,
} from "./cobolsourcescanner";
import { ICOBOLSourceScanner } from "./icobolsourcescanner";

// ---------------- Program Window State ----------------
class ProgramWindowState {
  currentStyle: string;
  currentProgram: ICOBOLSourceScanner;

  private constructor(currentStyle: string, currentProgram: ICOBOLSourceScanner) {
    this.currentStyle = currentStyle;
    this.currentProgram = currentProgram;
  }

  private static styleMap = new Map<string, ProgramWindowState>();

  static get(url: string, program: ICOBOLSourceScanner): ProgramWindowState {
    let state = this.styleMap.get(url);
    if (!state) {
      state = new ProgramWindowState("TD", program);
      this.styleMap.set(url, state);
    }
    return state;
  }
}

// ---------------- Generate Partial Graph ----------------
function generatePartialGraph(
  lines: string[],
  clickLines: string[],
  state: ParseState,
  tokens: Map<string, COBOLToken>
) {
  for (const [name, token] of tokens) {
    const nameLower = name.toLowerCase();
    const references = state.currentSectionOutRefs.get(nameLower);

    clickLines.push(
      `click ${token.tokenNameLower} call callback("${token.tokenName}","${token.filenameAsURI}",${token.startLine},${token.startColumn}) "${token.description}"`
    );

    if (!references) continue;

    if (token.isImplicitToken) {
      lines.push(`${token.tokenNameLower}[${token.description}]`);
    }

    const tempLines: string[] = [];
    for (const ref of references) {
      if (ref.line === token.startLine && ref.column === token.startColumn) continue;

      if (ref.tokenStyle === COBOLTokenStyle.Paragraph || ref.tokenStyle === COBOLTokenStyle.Section) {
        const edge = ref.reason === "perform"
          ? `${token.tokenNameLower} --> ${ref.nameLower}`
          : `${token.tokenNameLower} -->|${ref.reason}|${ref.nameLower}`;
        tempLines.push(edge);
      }
    }

    for (const item of new Set(tempLines)) lines.push(item);
  }
}

// ---------------- DotGraphPanelView ----------------
export class DotGraphPanelView {
  public static currentPanel: DotGraphPanelView | undefined;
  private readonly panel: vscode.WebviewPanel;
  private disposables: vscode.Disposable[] = [];

  private constructor(
    context: vscode.ExtensionContext,
    url: string,
    state: ProgramWindowState,
    panel: vscode.WebviewPanel,
    lines: string[]
  ) {
    this.panel = panel;
    this.panel.webview.html = this.getWebviewContent(context, url, state, panel, lines);
    this.panel.onDidDispose(() => this.dispose(), null, this.disposables);
  }

  public static render(
    context: vscode.ExtensionContext,
    url: string,
    state: ProgramWindowState,
    lines: string[],
    programName: string
  ): vscode.WebviewPanel {
    if (DotGraphPanelView.currentPanel) {
      const current = DotGraphPanelView.currentPanel;
      current.panel.reveal(vscode.ViewColumn.Beside, true);
      current.panel.webview.html = current.getWebviewContent(context, url, state, current.panel, lines);
      return current.panel;
    }

    const resourcesDir = vscode.Uri.joinPath(context.extensionUri, "resources");
    const panel = vscode.window.createWebviewPanel(
      "cobolCallGraph",
      `Program: ${programName}`,
      { viewColumn: vscode.ViewColumn.Beside, preserveFocus: true },
      { enableScripts: true, localResourceRoots: [resourcesDir] }
    );

    DotGraphPanelView.currentPanel = new DotGraphPanelView(context, url, state, panel, lines);
    return panel;
  }

  private getWebviewContent(
    context: vscode.ExtensionContext,
    url: string,
    state: ProgramWindowState,
    panel: vscode.WebviewPanel,
    lines: string[]
  ): string {
    const style = state.currentStyle;
    let html = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src cspSource 'unsafe-inline'; script-src 'nonce-nonce' 'unsafe-inline';">
<style>
.diagram-container { width: 98%; overflow: hidden; border: 1px solid #ccc; margin: 0 auto 10px auto; display: block; }
svg { cursor: grab; }
</style>
</head>
<body class="vscode-body">
<script>
const _config = { vscode: { minimap: false, dark: 'dark', light: 'neutral' } };
</script>

<script src="panzoom.min.js"></script>
<script type="module">
import mermaid from 'mermaid.esm.min.mjs';
const vscode = acquireVsCodeApi();

window.callback = (message,filename,line,col) => {
  vscode.postMessage({ command: 'golink', text: message + "," + filename + "," + line + "," + col });
};

document.querySelector('#chart-style').addEventListener("change", function() {
  vscode.postMessage({ command: 'change-style', text: "__url__" + "," + this.value });
});

mermaid.initialize({ startOnLoad: false, useMaxWidth: true, theme: "neutral", securityLevel: 'loose' });
await mermaid.run({
  querySelector: '.mermaid',
  postRenderCallback: (id) => {
    const container = document.getElementById("diagram-container");
    const svg = container.querySelector("svg");
    const panzoomInstance = Panzoom(svg, { maxScale:5, minScale:0.5, step:0.1 });
    container.addEventListener("wheel", e => panzoomInstance.zoomWithWheel(e));
  }
});
</script>

<h3>Program: __program__</h3>
<div class="diagram-container" id="diagram-container">
  <div class="mermaid" id="mermaid1">${lines.join("\n")}</div>
</div>

<div class="vscode-select">
<p>Style:
<select name="chart-style" id="chart-style">
<option value="" selected hidden>Choose here</option>
<option value="TD">Top Down</option>
<option value="BT">Bottom-to-top</option>
<option value="RL">Right-to-left</option>
<option value="LR">Left-to-right</option>
</select>
</p>
</div>

<hr class="vscode-divider">
<h5>This is a navigation aid, not a source analysis feature.</h5>
</body>
</html>`;

    // Replace resources
    html = html.replace(
      "mermaid.esm.min.mjs",
      panel.webview.asWebviewUri(vscode.Uri.joinPath(context.extensionUri, "resources/mermaid/mermaid.esm.min.mjs")).toString()
    );
    html = html.replace(
      "panzoom.min.js",
      panel.webview.asWebviewUri(vscode.Uri.joinPath(context.extensionUri, "resources/panzoom/panzoom.min.js")).toString()
    );

    const cssPath = vscode.Uri.joinPath(context.extensionUri, "resources/vscode-elements.css");
    html = html.replace("vscode-elements.", fs.readFileSync(cssPath.fsPath).toString());

    const nonce = getNonce();
    html = html.replace(/nonce-nonce/g, `nonce-${nonce}`);
    html = html.replace(/<script /g, `<script nonce="${nonce}" `);
    html = html.replace(/<link /g, `<link nonce="${nonce}" `);
    html = html.replace("cspSource", panel.webview.cspSource);
    html = html.replace("__url__", url);
    html = html.replace("__style__", style);
    html = html.replace("__program__", getProgramName(state.currentProgram));

    return html;
  }

  public dispose() {
    DotGraphPanelView.currentPanel = undefined;
    this.panel.dispose();
    this.disposables.forEach(d => d.dispose());
    this.disposables = [];
  }
}

// ---------------- Utility Functions ----------------
function getNonce(): string {
  return Array.from({ length: 32 }, () =>
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789".charAt(Math.floor(Math.random() * 62))
  ).join("");
}

function getProgramName(program: ICOBOLSourceScanner): string {
  return program.ImplicitProgramId || program.ProgramId || "?";
}

function getCurrentProgramCallGraph(state: ProgramWindowState, asMarkdown: boolean, includeEvents: boolean): string[] {
  const current = state.currentProgram;
  const lines: string[] = [];
  const clickLines: string[] = [];
  const parseState = current.sourceReferences.state;

  if (asMarkdown) {
    lines.push(`# ${getProgramName(current)}`, "", "```mermaid");
  }

  lines.push(`flowchart ${state.currentStyle};`);
  generatePartialGraph(lines, clickLines, parseState, current.sections);
  generatePartialGraph(lines, clickLines, parseState, current.paragraphs);

  if (includeEvents) lines.push(...clickLines);
  if (asMarkdown) lines.push("```");

  return lines;
}

async function gotoLink(message: string) {
  const [_, url, lineStr, colStr] = message.split(",");
  const line = parseInt(lineStr);
  const col = parseInt(colStr);
  const pos = new vscode.Position(line, col);
  const doc = await vscode.workspace.openTextDocument(vscode.Uri.parse(url));
  const editor = await vscode.window.showTextDocument(doc);
  editor.selections = [new vscode.Selection(pos, pos)];
  editor.revealRange(new vscode.Range(pos, pos));
}

// ---------------- Exports ----------------
export async function view_dot_callgraph(context: vscode.ExtensionContext, settings: ICOBOLSettings) {
  if (!settings.enable_program_information || !vscode.window.activeTextEditor) return;

  const current = VSCOBOLSourceScanner.getCachedObject(vscode.window.activeTextEditor.document, settings);
  if (!current) return;

  const url = current.sourceHandler.getUriAsString();
  const state = ProgramWindowState.get(url, current);
  const lines = getCurrentProgramCallGraph(state, false, true);
  const webviewPanel = DotGraphPanelView.render(context, url, state, lines, getProgramName(current));

  webviewPanel.webview.onDidReceiveMessage(async message => {
    switch (message.command) {
      case "golink": await gotoLink(message.text); break;
      case "change-style": {
        const [newUrl, newStyle] = message.text.split(",");
        state.currentStyle = newStyle;
        const newLines = getCurrentProgramCallGraph(state, false, true);
        DotGraphPanelView.render(context, newUrl, state, newLines, getProgramName(current));
        break;
      }
    }
  }, undefined, context.subscriptions);

  vscode.workspace.onDidChangeTextDocument(event => {
    if (!settings.enable_program_information) return;
    if (vscode.window.activeTextEditor?.document.uri !== event.document.uri) return;

    const updated = VSCOBOLSourceScanner.getCachedObject(event.document, settings);
    if (!updated) return;

    const updatedState = ProgramWindowState.get(updated.sourceHandler.getUriAsString(), updated);
    updatedState.currentProgram = updated;
    const updatedLines = getCurrentProgramCallGraph(updatedState, false, true);
    DotGraphPanelView.render(context, updated.sourceHandler.getUriAsString(), updatedState, updatedLines, getProgramName(updated));
  });
}

export async function newFile_dot_callgraph(settings: ICOBOLSettings) {
  if (!vscode.window.activeTextEditor) return;

  const current = VSCOBOLSourceScanner.getCachedObject(vscode.window.activeTextEditor.document, settings);
  if (!current) return;

  const state = ProgramWindowState.get(current.sourceHandler.getUriAsString(), current);
  state.currentProgram = current;

  const lines = getCurrentProgramCallGraph(state, true, false);
  const doc = await vscode.workspace.openTextDocument({ language: "markdown" });
  const editor = await vscode.window.showTextDocument(doc);

  const snippet = new vscode.SnippetString(lines.join("\n"));
  await editor.insertSnippet(snippet, new vscode.Range(0, 0, lines.length + 1, 0));
}