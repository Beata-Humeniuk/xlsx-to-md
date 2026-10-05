'use strict';

// A value field with a drop-down of the values that exist in the column:
// search, tick one or several, or use a typed value that is not in the list.

let openPopup = null;

function closePopup() {
  if (openPopup) {
    openPopup.remove();
    openPopup = null;
  }
}

document.addEventListener('mousedown', (e) => {
  if (openPopup && !openPopup.contains(e.target) && !(openPopup.owner && openPopup.owner.contains(e.target))) closePopup();
}, true);
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && openPopup) {
    // Closes only the list, not a dialog the picker sits in.
    e.preventDefault();
    e.stopImmediatePropagation();
    const owner = openPopup.owner;
    closePopup();
    if (owner) owner.focus();
  }
}, true);
window.addEventListener('blur', closePopup);

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
}

// options: {
//   values: [string], multi: boolean, placeholder: string,
//   options: () => [{ value, count }],     (called when the list opens)
//   onChange(values), t(key, params)
// }
function valuePicker(opts) {
  let values = (opts.values || []).slice();
  const box = el('div', 'picker');
  box.tabIndex = 0;
  box.setAttribute('role', 'combobox');
  box.setAttribute('aria-haspopup', 'listbox');

  const draw = () => {
    box.replaceChildren();
    if (!values.length) {
      box.appendChild(el('span', 'picker-placeholder', opts.placeholder || ''));
    } else {
      for (const v of values) {
        const chip = el('span', 'chip');
        chip.appendChild(el('span', 'chip-text', v));
        if (opts.multi) {
          const x = el('button', 'chip-x', '×');
          x.type = 'button';
          x.tabIndex = -1;
          x.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            set(values.filter((y) => y !== v));
          });
          chip.appendChild(x);
        }
        box.appendChild(chip);
      }
    }
    box.appendChild(el('span', 'picker-caret codicon-like', '▾'));
  };

  const set = (next) => {
    values = next;
    draw();
    opts.onChange(values.slice());
    if (openPopup && openPopup.owner === box) openPopup.refresh();
  };

  const open = () => {
    if (openPopup && openPopup.owner === box) {
      closePopup();
      return;
    }
    closePopup();
    const list = opts.options ? opts.options() : [];
    const pop = el('div', 'picker-pop');
    pop.owner = box;
    const search = el('input', 'picker-search');
    search.type = 'text';
    search.placeholder = opts.t('ui.search');
    pop.appendChild(search);
    const ul = el('div', 'picker-list');
    ul.setAttribute('role', 'listbox');
    pop.appendChild(ul);
    const foot = el('div', 'picker-foot');
    const clear = el('button', 'link', opts.t('ui.clear'));
    clear.type = 'button';
    clear.addEventListener('click', () => {
      set([]);
      search.focus();
    });
    foot.appendChild(clear);
    pop.appendChild(foot);
    let focusIndex = 0;

    const refresh = () => {
      const q = search.value.trim().toLowerCase();
      ul.replaceChildren();
      const shown = list.filter((o) => !q || o.value.toLowerCase().includes(q)).slice(0, 500);
      const typed = search.value.trim();
      const items = [];
      if (typed && !list.some((o) => o.value.toLowerCase() === typed.toLowerCase())) {
        items.push({ value: typed, label: opts.t('ui.addValue', { value: typed }), typed: true });
      }
      for (const o of shown) items.push({ value: o.value, label: o.value, count: o.count });
      // Values chosen earlier that are not in this data stay visible.
      for (const v of values) {
        if (!items.some((i) => i.value === v) && (!q || v.toLowerCase().includes(q))) items.push({ value: v, label: v });
      }
      if (focusIndex >= items.length) focusIndex = Math.max(0, items.length - 1);
      items.forEach((item, i) => {
        const row = el('div', 'picker-item' + (i === focusIndex ? ' is-focus' : '') + (item.typed ? ' is-typed' : ''));
        row.setAttribute('role', 'option');
        const on = values.includes(item.value);
        row.setAttribute('aria-selected', String(on));
        if (opts.multi) {
          const cb = el('span', 'check' + (on ? ' is-on' : ''), on ? '✓' : '');
          row.appendChild(cb);
        }
        row.appendChild(el('span', 'picker-label', item.label));
        if (item.count !== undefined) row.appendChild(el('span', 'picker-count', String(item.count)));
        row.addEventListener('mousedown', (e) => {
          e.preventDefault();
          choose(item);
        });
        ul.appendChild(row);
      });
      pop.items = items;
    };

    const choose = (item) => {
      if (opts.multi) {
        const on = values.includes(item.value);
        set(on ? values.filter((v) => v !== item.value) : values.concat([item.value]));
        if (item.typed) search.value = '';
        refresh();
        search.focus();
      } else {
        set([item.value]);
        closePopup();
        box.focus();
      }
    };

    search.addEventListener('input', () => {
      focusIndex = 0;
      refresh();
    });
    search.addEventListener('keydown', (e) => {
      const items = pop.items || [];
      if (e.key === 'ArrowDown') {
        focusIndex = Math.min(items.length - 1, focusIndex + 1);
        refresh();
        e.preventDefault();
      } else if (e.key === 'ArrowUp') {
        focusIndex = Math.max(0, focusIndex - 1);
        refresh();
        e.preventDefault();
      } else if (e.key === 'Enter') {
        if (items[focusIndex]) choose(items[focusIndex]);
        e.preventDefault();
      } else if (e.key === 'Tab') {
        closePopup();
      }
    });

    pop.refresh = refresh;
    refresh();
    document.body.appendChild(pop);
    const rect = box.getBoundingClientRect();
    const popW = Math.max(rect.width, 220);
    pop.style.width = popW + 'px';
    pop.style.left = Math.min(rect.left, window.innerWidth - popW - 8) + 'px';
    const below = window.innerHeight - rect.bottom;
    if (below < 260 && rect.top > below) {
      pop.style.bottom = window.innerHeight - rect.top + 2 + 'px';
    } else {
      pop.style.top = rect.bottom + 2 + 'px';
    }
    openPopup = pop;
    search.focus();
  };

  box.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    open();
  });
  box.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ' || e.key === 'ArrowDown') {
      e.preventDefault();
      open();
    } else if ((e.key === 'Backspace' || e.key === 'Delete') && values.length) {
      set(values.slice(0, -1));
    }
  });
  draw();
  return box;
}

module.exports = { valuePicker, closePopup };
