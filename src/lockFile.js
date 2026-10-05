'use strict';

// Excel's owner ("lock") files: while a workbook is open, Excel keeps a small
// hidden file next to it named "~$" + the workbook name — for longer names
// the "~$" replaces the first two characters ("Index ante.xlsx" →
// "~$ex ante.xlsx"). The file holds only who has the workbook open.

function isLockFile(name) {
  return /^~\$/.test(String(name));
}

// The workbook a lock file belongs to, among the names in the same folder;
// null when it cannot be told.
function workbookFor(lockName, siblings) {
  const rest = String(lockName).slice(2);
  const candidates = siblings.filter((n) => !isLockFile(n) && n !== lockName);
  const exact = candidates.find((n) => n === rest);
  if (exact) return exact;
  const replaced = candidates.filter((n) => n.length === rest.length + 2 && n.slice(2) === rest);
  return replaced.length === 1 ? replaced[0] : null;
}

module.exports = { isLockFile, workbookFor };
