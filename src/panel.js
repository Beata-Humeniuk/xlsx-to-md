'use strict';

// The editor for one workbook: a webview with the sheet grid, the data
// choices and the Markdown preview. The extension side reads the workbook,
// finds a saved filter's data in it, keeps the saved filters and writes the
// exported file; everything the user clicks is handled in the webview.

const vscode = require('vscode');
const path = require('path');
const { openWorkbook, sniff } = require('./workbook');
const { openCsv, isCsvName } = require('./csv');
const { isLockFile, workbookFor } = require('./lockFile');
const { resolveSource } = require('./core/resolve');
const { messagesFor, baseLanguage } = require('./nls');
const { summary, conditionsText } = require('./filters');

function nonce() {
  let s = '';
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}

// The sheet in a compact form for the webview: kinds as one string per row,
// numbers as a sparse list of [row, col, value].
function packSheet(sheet, date1904) {
  const kinds = sheet.kind.map((row) => row.map((k) => k || '.').join(''));
  const nums = [];
  sheet.num.forEach((row, r) => row.forEach((n, c) => {
    if (n !== null) nums.push(r, c, n);
  }));
  return {
    name: sheet.name,
    hidden: sheet.hidden,
    rows: sheet.rows,
    cols: sheet.cols,
    text: sheet.text,
    kinds,
    nums,
    hiddenRows: sheet.hiddenRows,
    hiddenCols: sheet.hiddenCols,
    merges: sheet.merges,
    tables: sheet.tables,
    date1904,
  };
}

// Whether a saved filter finds its data in this workbook.
function matches(workbook, recipe) {
  if (!recipe.source.ranges.length) return false;
  try {
    const hit = resolveSource(workbook, recipe);
    return hit.ranges.length > 0 && !hit.notes.some((n) => n.code === 'headersNotFound' || n.code === 'sheetMissing');
  } catch {
    return false;
  }
}

class WorkbookPanel {
  constructor({ context, store, nls, panel, uri, onDispose }) {
    this.context = context;
    this.store = store;
    this.nls = nls;
    this.panel = panel;
    this.uri = uri;
    this.workbook = null;
    this.fileName = path.basename(uri.path);
    const webview = panel.webview;
    webview.options = {
      enableScripts: true,
      localResourceRoots: [vscode.Uri.joinPath(context.extensionUri, 'out'), vscode.Uri.joinPath(context.extensionUri, 'media')],
    };
    webview.html = this.html();
    this.disposables = [
      webview.onDidReceiveMessage((m) => this.onMessage(m).catch((e) => this.fail(e))),
      store.onDidChange(() => this.postFilters()),
      panel.onDidDispose(() => {
        for (const d of this.disposables) d.dispose();
        if (onDispose) onDispose(this);
      }),
    ];
  }

  html() {
    const webview = this.panel.webview;
    const n = nonce();
    const script = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'out', 'webview.js'));
    const style = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'media', 'webview.css'));
    const lang = baseLanguage(vscode.env.language);
    return '<!DOCTYPE html>\n<html lang="' + lang + '">\n<head>\n<meta charset="UTF-8">\n'
      + '<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src ' + webview.cspSource
      + ' \'unsafe-inline\'; script-src \'nonce-' + n + '\'; font-src ' + webview.cspSource + ';">\n'
      + '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
      + '<link rel="stylesheet" href="' + style + '">\n<title>' + escapeHtml(this.fileName) + '</title>\n</head>\n'
      + '<body>\n<div id="app" class="loading">' + escapeHtml(this.nls.t('ui.loading')) + '</div>\n'
      + '<script nonce="' + n + '" src="' + script + '"></script>\n</body>\n</html>';
  }

  post(message) {
    return this.panel.webview.postMessage(message);
  }

  fail(e) {
    if (e && e.shown) return;
    vscode.window.showErrorMessage(this.nls.t('error.failed', { file: this.fileName, reason: e && e.message ? e.message : String(e) }));
  }

  // Excel's lock file "~$name" instead of the workbook: say so and offer
  // the workbook itself when it is in the same folder.
  async lockFileError() {
    const dir = vscode.Uri.joinPath(this.uri, '..');
    let names = [];
    try {
      names = (await vscode.workspace.fs.readDirectory(dir)).map(([n]) => n);
    } catch {
      names = [];
    }
    const original = workbookFor(this.fileName, names);
    this.original = original ? vscode.Uri.joinPath(dir, original) : null;
    const message = original
      ? this.nls.t('error.lockFile', { file: this.fileName, original })
      : this.nls.t('error.lockFileUnknown', { file: this.fileName });
    return Object.assign(new Error(message), {
      shown: true,
      action: original ? this.nls.t('action.openWorkbook', { file: original }) : null,
    });
  }

  async load() {
    if (isLockFile(this.fileName)) throw await this.lockFileError();
    const bytes = Buffer.from(await vscode.workspace.fs.readFile(this.uri));
    const kind = sniff(bytes);
    if (isCsvName(this.fileName) && kind !== 'xlsx') {
      this.workbook = openCsv(bytes, this.fileName);
      return;
    }
    if (kind === 'xls') throw Object.assign(new Error(this.nls.t('error.oldXls', { file: this.fileName })), { shown: true });
    if (kind !== 'xlsx') throw Object.assign(new Error(this.nls.t('error.notXlsx', { file: this.fileName })), { shown: true });
    try {
      this.workbook = openWorkbook(bytes);
    } catch (e) {
      if (e.code === 'NOT_XLSX') throw Object.assign(new Error(this.nls.t('error.notXlsx', { file: this.fileName })), { shown: true });
      if (e.code === 'NO_SHEETS') throw Object.assign(new Error(this.nls.t('error.noSheets', { file: this.fileName })), { shown: true });
      throw e;
    }
  }

  filterList() {
    return this.store.list().map((f) => ({
      id: f.id,
      name: f.name,
      summary: summary(f, this.nls),
      conditions: conditionsText(f, this.nls),
      recipe: f.recipe,
      matches: this.workbook ? matches(this.workbook, f.recipe) : false,
    }));
  }

  postFilters() {
    return this.post({ type: 'filters', filters: this.filterList() });
  }

  sheetMessage(index) {
    return { type: 'sheet', index, sheet: packSheet(this.workbook.loadSheet(index), this.workbook.date1904) };
  }

  async onMessage(m) {
    switch (m.type) {
      case 'ready': return this.onReady();
      case 'loadSheet': return this.post(this.sheetMessage(m.index));
      case 'applyFilter': return this.applyFilter(m.filterId);
      case 'saveFilter': return this.saveFilter(m);
      case 'rememberValues':
        if (vscode.workspace.getConfiguration('xlsxToMd').get('rememberValues') !== false) {
          await this.store.rememberValues(m.filterId, m.values || {});
        }
        return undefined;
      case 'openOriginal':
        if (this.original) {
          await vscode.commands.executeCommand('vscode.openWith', this.original, 'xlsxToMd.workbook');
          this.panel.dispose();
        }
        return undefined;
      case 'manageFilters': return vscode.commands.executeCommand('xlsxToMd.manageFilters');
      case 'export': return this.exportMarkdown(m.markdown, m.nameHint);
      case 'copy':
        await vscode.env.clipboard.writeText(m.markdown || '');
        vscode.window.showInformationMessage(this.nls.t('info.copied'));
        return undefined;
      case 'openUntitled': {
        const doc = await vscode.workspace.openTextDocument({ language: 'markdown', content: m.markdown || '' });
        await vscode.window.showTextDocument(doc, { preview: false });
        return undefined;
      }
      default: return undefined;
    }
  }

  async onReady() {
    if (!this.workbook) {
      try {
        await this.load();
      } catch (e) {
        this.post({
          type: 'fatal',
          message: e.shown ? e.message : this.nls.t('error.failed', { file: this.fileName, reason: e.message }),
          action: e.action || null,
        });
        throw Object.assign(e, { shown: !!e.shown });
      }
    }
    const first = Math.max(0, this.workbook.sheets.findIndex((s) => !s.hidden));
    await this.post({
      type: 'init',
      messages: messagesFor(vscode.env.language),
      lang: baseLanguage(vscode.env.language),
      fileName: this.fileName,
      sheets: this.workbook.sheets,
      firstSheet: first,
      filters: this.filterList(),
    });
    await this.post(this.sheetMessage(first));
  }

  async applyFilter(filterId) {
    const filter = this.store.get(filterId);
    if (!filter) return this.postFilters();
    const hit = resolveSource(this.workbook, filter.recipe);
    await this.post(this.sheetMessage(hit.sheetIndex));
    const remember = vscode.workspace.getConfiguration('xlsxToMd').get('rememberValues') !== false;
    return this.post({
      type: 'applied',
      filterId: filter.id,
      name: filter.name,
      recipe: filter.recipe,
      sheetIndex: hit.sheetIndex,
      ranges: hit.ranges,
      notes: hit.notes,
      lastValues: remember ? this.store.lastValues(filter.id) : {},
    });
  }

  async saveFilter(m) {
    const nls = this.nls;
    if (m.mode === 'update' && m.filterId) {
      const existing = this.store.get(m.filterId);
      if (existing) {
        const recipe = m.keepSource ? { ...m.recipe, source: existing.recipe.source } : m.recipe;
        const saved = await this.store.save({ id: existing.id, name: existing.name, recipe });
        vscode.window.showInformationMessage(nls.t('info.filterUpdated', { name: saved.name }));
        return this.post({ type: 'saved', filterId: saved.id, name: saved.name, recipe: saved.recipe });
      }
    }
    const name = await vscode.window.showInputBox({
      title: nls.t('ui.saveAsFilter').replace(/…$/, ''),
      prompt: nls.t('prompt.filterName'),
      placeHolder: nls.t('prompt.filterNamePlaceholder'),
      value: m.suggestedName || '',
      validateInput: (v) => (String(v).trim() ? null : nls.t('prompt.filterNameEmpty')),
    });
    if (!name || !name.trim()) return undefined;
    let id = null;
    const clash = this.store.findByName(name);
    if (clash) {
      const replace = nls.t('prompt.replace');
      const choice = await vscode.window.showWarningMessage(nls.t('prompt.filterExists', { name: clash.name }), { modal: true }, replace);
      if (choice !== replace) return undefined;
      id = clash.id;
    }
    const saved = await this.store.save({ id, name, recipe: m.recipe });
    if (m.values && Object.keys(m.values).length) await this.store.rememberValues(saved.id, m.values);
    vscode.window.showInformationMessage(nls.t('info.filterSaved', { name: saved.name }));
    return this.post({ type: 'saved', filterId: saved.id, name: saved.name, recipe: saved.recipe });
  }

  async exportMarkdown(markdown, nameHint) {
    const base = path.basename(this.uri.path, path.extname(this.uri.path));
    const fileName = (nameHint ? base + ' - ' + sanitize(nameHint) : base) + '.md';
    const dir = vscode.Uri.joinPath(this.uri, '..');
    const target = await vscode.window.showSaveDialog({
      defaultUri: vscode.Uri.joinPath(dir, fileName),
      saveLabel: this.nls.t('dialog.saveLabel'),
      filters: { [this.nls.t('dialog.markdown')]: ['md', 'markdown'] },
    });
    if (!target) return;
    await vscode.workspace.fs.writeFile(target, Buffer.from(markdown || '', 'utf8'));
    const open = vscode.workspace.getConfiguration('xlsxToMd').get('openAfterExport');
    if (open === 'editor') await vscode.window.showTextDocument(target, { preview: false });
    else if (open === 'preview') await vscode.commands.executeCommand('markdown.showPreview', target);
    else vscode.window.showInformationMessage(this.nls.t('info.exported', { file: path.basename(target.path) }));
  }
}

function sanitize(name) {
  return String(name).replace(/[\\/:*?"<>|\u0000-\u001f]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80);
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

module.exports = { WorkbookPanel, packSheet, matches };
