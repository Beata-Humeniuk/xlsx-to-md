'use strict';

const vscode = require('vscode');
const { forLanguage } = require('./nls');
const { FilterStore, summary, conditionsText } = require('./filters');
const { WorkbookPanel } = require('./panel');

const VIEW_TYPE = 'xlsxToMd.workbook';
const EXTENSIONS = ['xlsx', 'xlsm', 'xltx', 'xltm'];

const nls = forLanguage(vscode.env && vscode.env.language);

// Opens .xlsx files with "Open With…" and from the extension's own command.
class WorkbookEditorProvider {
  constructor(context, store) {
    this.context = context;
    this.store = store;
    this.panels = new Set();
  }

  openCustomDocument(uri) {
    return { uri, dispose() {} };
  }

  resolveCustomEditor(document, webviewPanel) {
    const panel = new WorkbookPanel({
      context: this.context,
      store: this.store,
      nls,
      panel: webviewPanel,
      uri: document.uri,
      onDispose: (p) => this.panels.delete(p),
    });
    this.panels.add(panel);
  }
}

async function open(uri, uris) {
  let files = Array.isArray(uris) && uris.length ? uris : uri instanceof vscode.Uri ? [uri] : null;
  if (!files) {
    files = await vscode.window.showOpenDialog({
      canSelectMany: false,
      openLabel: nls.t('dialog.openLabel'),
      filters: { [nls.t('dialog.filter')]: EXTENSIONS.concat('xls') },
    }) || [];
  }
  for (const file of files) {
    if (/\.xls$/i.test(file.path)) {
      vscode.window.showErrorMessage(nls.t('error.oldXls', { file: file.path.split('/').pop() }));
      continue;
    }
    await vscode.commands.executeCommand('vscode.openWith', file, VIEW_TYPE);
  }
}

// A list of the saved filters with rename and delete buttons.
async function manageFilters(store) {
  if (!store.list().length) {
    vscode.window.showInformationMessage(nls.t('info.noFilters'));
    return;
  }
  const renameButton = { iconPath: new vscode.ThemeIcon('edit'), tooltip: nls.t('action.rename') };
  const deleteButton = { iconPath: new vscode.ThemeIcon('trash'), tooltip: nls.t('action.delete') };
  const pick = vscode.window.createQuickPick();
  pick.title = nls.t('manage.title');
  pick.placeholder = nls.t('manage.placeholder');
  pick.matchOnDescription = true;
  pick.matchOnDetail = true;
  const refresh = () => {
    const filters = store.list();
    if (!filters.length) {
      pick.hide();
      return;
    }
    pick.items = filters.map((f) => ({
      label: f.name,
      description: summary(f, nls),
      detail: conditionsText(f, nls) || undefined,
      buttons: [renameButton, deleteButton],
      filter: f,
    }));
  };
  refresh();
  const sub = store.onDidChange(refresh);
  pick.onDidTriggerItemButton(async ({ item, button }) => {
    const f = item.filter;
    if (button === deleteButton) {
      const del = nls.t('action.delete');
      const choice = await vscode.window.showWarningMessage(nls.t('prompt.deleteFilter', { name: f.name }), { modal: true }, del);
      if (choice === del) {
        await store.remove(f.id);
        vscode.window.showInformationMessage(nls.t('info.filterDeleted', { name: f.name }));
      }
    } else if (button === renameButton) {
      const name = await vscode.window.showInputBox({
        prompt: nls.t('prompt.renameFilter'),
        value: f.name,
        validateInput: (v) => {
          if (!String(v).trim()) return nls.t('prompt.filterNameEmpty');
          const clash = store.findByName(v);
          return clash && clash.id !== f.id ? nls.t('prompt.nameTaken', { name: clash.name }) : null;
        },
      });
      if (name && name.trim()) await store.rename(f.id, name);
    }
    pick.show();
  });
  pick.onDidHide(() => {
    sub.dispose();
    pick.dispose();
  });
  pick.show();
}

function activate(context) {
  const store = new FilterStore(context.globalState);
  const provider = new WorkbookEditorProvider(context, store);
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(VIEW_TYPE, provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: true,
    }),
    vscode.commands.registerCommand('xlsxToMd.open', open),
    vscode.commands.registerCommand('xlsxToMd.manageFilters', () => manageFilters(store)),
  );
}

function deactivate() {}

module.exports = { activate, deactivate };
