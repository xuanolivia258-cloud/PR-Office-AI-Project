const bundledDataRevision = '1795710786.8378496';

const defaultRecords = [
  { prPreparer: 'Olivia', prNumber: '705157580003567', prLineDescription: 'Computer', prCreationDate: '2026-10-03', prStatus: '—', poNumber: '7051H00005758', poCreationDate: '2026-11-26', needByDate: '', servicePeriodStartDate: '2026-10-30', servicePeriodEndDate: '2026-11-30', poLineStatus: '—', poLineAmount: 0, receivedAmount: 1000, billedAmount: 1000, sourceFile: 'procurement-projects.xlsx', sourceSheet: 'Projects' },
  { prPreparer: 'Elaine', prNumber: '12345678987654300', prLineDescription: '—', prCreationDate: '', prStatus: '—', poNumber: '7051H00003890', poCreationDate: '2026-11-25', needByDate: '', servicePeriodStartDate: '2026-09-30', servicePeriodEndDate: '2026-10-05', poLineStatus: '—', poLineAmount: 0, receivedAmount: 0, billedAmount: 0, sourceFile: 'procurement-projects.xlsx', sourceSheet: 'Projects' }
];

function loadStoredRecords() {
  try {
    const stored = JSON.parse(localStorage.getItem('pr-po-tracker-records'));
    if (stored?.revision === bundledDataRevision && Array.isArray(stored.records)) return stored.records;
    const highlightKeys = [];
    for (let index = 0; index < localStorage.length; index += 1) {
      const key = localStorage.key(index);
      if (key?.startsWith('pr-po-creation-highlight-dismissed:')) highlightKeys.push(key);
    }
    highlightKeys.forEach(key => localStorage.removeItem(key));
    localStorage.setItem('pr-po-tracker-records', JSON.stringify({ revision: bundledDataRevision, records: defaultRecords }));
    return defaultRecords;
  } catch {
    return defaultRecords;
  }
}

const records = loadStoredRecords();

const tableBody = document.querySelector('#requestTableBody');
const searchInput = document.querySelector('#searchInput');
const resultCount = document.querySelector('#resultCount');
const modal = document.querySelector('#requestModal');
const toast = document.querySelector('#toast');
const settingsView = document.querySelector('#settingsView');
const personSearchForm = document.querySelector('#personSearchForm');
const personSearchInput = document.querySelector('#personSearchInput');
const personSearchResult = document.querySelector('#personSearchResult');
const peopleList = document.querySelector('#peopleList');
const accountMenuButton = document.querySelector('#accountMenuButton');
const accountMenu = document.querySelector('#accountMenu');
const excelFileInput = document.querySelector('#excelFileInput');
const excelDropZone = document.querySelector('#excelDropZone');
const dropPrompt = document.querySelector('#dropPrompt');
const excelFileList = document.querySelector('#excelFileList');
const importStatus = document.querySelector('#importStatus');
const confirmExcelImport = document.querySelector('#confirmExcelImport');
const poContextMenu = document.createElement('div');
poContextMenu.className = 'po-context-menu';
poContextMenu.hidden = true;
poContextMenu.innerHTML = '<button type="button">Restore Highlight</button>';
document.body.append(poContextMenu);

const assignedPeople = [
  { id: 'local-test-user', name: 'Local Test User', role: 'Administrator', initialAdmin: true }
];

let pendingPerson = null;
let pendingExcelFiles = [];
let contextMenuItem = null;
let toastTimer = null;
let pendingUndoAction = null;

function money(value) {
  if (value === '' || value === null || value === undefined) return '—';
  return new Intl.NumberFormat('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 }).format(Number(value) || 0);
}

function display(value) {
  return value === '' || value === null || value === undefined ? '—' : String(value);
}

function escapeHtml(value) {
  return display(value).replace(/[&<>'"]/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[character]));
}

function escapeAttribute(value) {
  return String(value ?? '').replace(/[&<>'"]/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[character]));
}

function remarkStorageKey(item) {
  const hasPoNumber = item.poNumber && item.poNumber !== '—';
  return `pr-po-remark:${hasPoNumber ? `po:${item.poNumber}` : `pr:${item.prNumber}`}`;
}

function savedRemark(item) {
  try {
    return localStorage.getItem(remarkStorageKey(item)) || localStorage.getItem(`pr-po-remark:${item.prNumber}`) || '';
  } catch {
    return '';
  }
}

function clearPreviousImportState() {
  try {
    const removablePrefixes = ['pr-po-creation-highlight-dismissed:'];
    const keysToRemove = [];
    for (let index = 0; index < localStorage.length; index += 1) {
      const key = localStorage.key(index);
      if (key && removablePrefixes.some(prefix => key.startsWith(prefix))) keysToRemove.push(key);
    }
    keysToRemove.forEach(key => localStorage.removeItem(key));
  } catch {
    // Continue the import if browser storage is unavailable.
  }
}

function poHighlightKey(item) {
  return `pr-po-creation-highlight-dismissed:${item.prNumber}:${item.poNumber}:${item.poCreationDate}`;
}

function isPoHighlightDismissed(item) {
  try {
    return localStorage.getItem(poHighlightKey(item)) === 'true';
  } catch {
    return false;
  }
}

function isPoWithinHighlightWindow(item) {
  if (!item.poNumber || item.poNumber === '—' || !item.poCreationDate) return false;
  const creationTime = new Date(item.poCreationDate).getTime();
  if (!Number.isFinite(creationTime)) return false;
  return Math.abs(Date.now() - creationTime) <= 48 * 60 * 60 * 1000;
}

function isRecentPo(item) {
  return isPoWithinHighlightWindow(item) && !isPoHighlightDismissed(item);
}

function statusClass(value) {
  if (value === 'Approved' || value === 'PO Released' || value === 'Delivered') return 'status success';
  if (value === 'Rejected') return 'status danger';
  if (value === 'Pending Approval' || value === 'PO Processing') return 'status warning';
  return 'status';
}

function isOverdue(item) {
  return item.needByDate && item.poLineStatus !== 'PO Released' && item.poLineStatus !== 'Delivered' && item.needByDate < new Date().toISOString().slice(0, 10);
}

function strictLocalDate(value) {
  const match = String(value || '').trim().match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day);
  return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day ? date : null;
}

function isWithinOrPastMonthBefore(value) {
  const target = strictLocalDate(value);
  if (!target) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  target.setHours(0, 0, 0, 0);
  const previousMonth = new Date(target.getFullYear(), target.getMonth() - 1, 1);
  const lastDayOfPreviousMonth = new Date(previousMonth.getFullYear(), previousMonth.getMonth() + 1, 0).getDate();
  const warningStart = new Date(previousMonth.getFullYear(), previousMonth.getMonth(), Math.min(target.getDate(), lastDayOfPreviousMonth));
  return today >= warningStart;
}

function hasUpcomingDateWarning(item) {
  return [item.needByDate, item.servicePeriodEndDate].some(isWithinOrPastMonthBefore);
}

function updateMetrics() {
  document.querySelector('#openPoCount').textContent = records.filter(item => item.poNumber && item.poNumber !== '—').length;
  document.querySelector('#poPendingIssuanceCount').textContent = records.filter(isRecentPo).length;
  document.querySelector('#dateAttentionCount').textContent = records.filter(hasUpcomingDateWarning).length;
}

function render() {
  const query = searchInput.value.trim().toLowerCase();
  const filtered = records.filter(item => {
    const matchesText = !query || [item.prNumber, item.poNumber, item.prPreparer, item.prLineDescription].some(value => String(value || '').toLowerCase().includes(query));
    return matchesText;
  });

  tableBody.replaceChildren();
  filtered.forEach(item => {
    const row = document.createElement('tr');
    if (hasUpcomingDateWarning(item)) row.classList.add('upcoming-date-row');
    const highlightPo = isRecentPo(item);
    row.innerHTML = `
      <td>${escapeHtml(item.prPreparer)}</td>
      <td><strong>${escapeHtml(item.prNumber)}</strong></td>
      <td>${escapeHtml(item.prLineDescription)}</td>
      <td>${escapeHtml(item.prCreationDate)}</td>
      <td><span class="${statusClass(item.prStatus)}">${escapeHtml(item.prStatus)}</span></td>
      <td>${highlightPo ? `<button class="po-number recent-po-highlight" type="button" title="PO created within the last 48 hours. Click to remove highlight." aria-label="PO ${escapeAttribute(item.poNumber)} was created within 48 hours. Click to remove highlight.">${escapeHtml(item.poNumber)}</button>` : `<span class="po-number" tabindex="0">${escapeHtml(item.poNumber)}</span>`}</td>
      <td>${escapeHtml(item.poCreationDate)}</td>
      <td>${escapeHtml(item.needByDate)}${isOverdue(item) ? '<small class="overdue-label">Overdue</small>' : ''}</td>
      <td>${escapeHtml(item.servicePeriodStartDate)}</td>
      <td>${escapeHtml(item.servicePeriodEndDate)}</td>
      <td><span class="${statusClass(item.poLineStatus)}">${escapeHtml(item.poLineStatus)}</span></td>
      <td>${money(item.poLineAmount)}</td>
      <td>${money(item.receivedAmount)}</td>
      <td>${money(item.billedAmount)}</td>
      <td><input class="remark-input" type="text" aria-label="Remark for PR ${escapeAttribute(item.prNumber)}" placeholder="Add a remark" value="${escapeAttribute(item.remark ?? savedRemark(item))}"></td>`;
    const remarkInput = row.querySelector('.remark-input');
    remarkInput.addEventListener('input', () => {
      item.remark = remarkInput.value;
      try {
        localStorage.setItem(remarkStorageKey(item), item.remark);
      } catch {
        // The current browser may block local storage; keep the remark for this session.
      }
    });
    const poHighlight = row.querySelector('.recent-po-highlight');
    poHighlight?.addEventListener('click', () => {
      try {
        localStorage.setItem(poHighlightKey(item), 'true');
      } catch {
        // If storage is unavailable, rerendering still removes the highlight for this view.
      }
      render();
      showUndoToast('PO highlight removed.', () => {
        try {
          localStorage.removeItem(poHighlightKey(item));
        } catch {
          // Keep the current view usable if storage is unavailable.
        }
        render();
        showToast('PO highlight restored.');
      });
    });
    row.querySelector('.po-number')?.addEventListener('contextmenu', event => {
      event.preventDefault();
      contextMenuItem = item;
      poContextMenu.hidden = false;
      const menuWidth = 170;
      const menuHeight = 44;
      poContextMenu.style.left = `${Math.min(event.clientX, window.innerWidth - menuWidth - 8)}px`;
      poContextMenu.style.top = `${Math.min(event.clientY, window.innerHeight - menuHeight - 8)}px`;
      poContextMenu.querySelector('button').focus();
    });
    tableBody.append(row);
  });

  resultCount.textContent = `${filtered.length} record${filtered.length === 1 ? '' : 's'}`;
  updateMetrics();
}

function showToast(message) {
  window.clearTimeout(toastTimer);
  pendingUndoAction = null;
  toast.textContent = message;
  toast.classList.add('show');
  toastTimer = window.setTimeout(() => toast.classList.remove('show'), 2400);
}

function showUndoToast(message, undo) {
  window.clearTimeout(toastTimer);
  pendingUndoAction = undo;
  toast.replaceChildren(document.createTextNode(message));
  const undoButton = document.createElement('button');
  undoButton.type = 'button';
  undoButton.className = 'toast-undo';
  undoButton.textContent = 'Undo';
  toast.append(undoButton);
  toast.classList.add('show');
  toastTimer = window.setTimeout(() => {
    pendingUndoAction = null;
    toast.classList.remove('show');
  }, 5000);
}

toast.addEventListener('click', event => {
  if (!event.target.classList.contains('toast-undo') || !pendingUndoAction) return;
  const undo = pendingUndoAction;
  pendingUndoAction = null;
  undo();
});

poContextMenu.querySelector('button').addEventListener('click', () => {
  const item = contextMenuItem;
  poContextMenu.hidden = true;
  if (!item) return;
  try {
    localStorage.removeItem(poHighlightKey(item));
  } catch {
    // Keep the current view usable if storage is unavailable.
  }
  render();
  showToast(isPoWithinHighlightWindow(item) ? 'PO highlight restored.' : 'This PO is outside the 48-hour highlight window.');
});

document.addEventListener('click', event => {
  if (!poContextMenu.contains(event.target)) poContextMenu.hidden = true;
});

document.addEventListener('keydown', event => {
  if (event.key === 'Escape') poContextMenu.hidden = true;
});

searchInput.addEventListener('input', render);

function showView(view, sourceButton = null) {
  const validView = ['dashboard', 'requests', 'settings'].includes(view) ? view : 'dashboard';
  document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
  if (sourceButton) sourceButton.classList.add('active');
  document.querySelector('#dashboardView').hidden = validView !== 'dashboard';
  settingsView.hidden = validView !== 'settings';
  const placeholder = validView === 'requests';
  document.querySelector('#placeholderView').hidden = !placeholder;
  if (placeholder && sourceButton) document.querySelector('#placeholderTitle').textContent = sourceButton.textContent.trim();
  try {
    localStorage.setItem('pr-po-current-view', validView);
  } catch {
    // Keep navigation working if browser storage is unavailable.
  }
}

document.querySelectorAll('.nav-item').forEach(button => button.addEventListener('click', () => {
  showView(button.dataset.view, button);
}));

function closeAccountMenu() {
  accountMenu.hidden = true;
  accountMenuButton.setAttribute('aria-expanded', 'false');
}

accountMenuButton.addEventListener('click', event => {
  event.stopPropagation();
  const shouldOpen = accountMenu.hidden;
  accountMenu.hidden = !shouldOpen;
  accountMenuButton.setAttribute('aria-expanded', String(shouldOpen));
  if (shouldOpen) accountMenu.querySelector('[role="menuitem"]').focus();
});

document.querySelector('#accountSettingsButton').addEventListener('click', () => {
  closeAccountMenu();
  showView('settings');
});

document.querySelector('#logoutButton').addEventListener('click', () => {
  closeAccountMenu();
  document.body.classList.add('signed-out');
  document.querySelector('#signInAgainButton').focus();
});

document.querySelector('#signInAgainButton').addEventListener('click', () => {
  document.body.classList.remove('signed-out');
  const dashboardButton = document.querySelector('[data-view="dashboard"]');
  showView('dashboard', dashboardButton);
  accountMenuButton.focus();
});

document.addEventListener('click', event => {
  if (!event.target.closest('.account-menu-wrap')) closeAccountMenu();
});

document.addEventListener('keydown', event => {
  if (event.key === 'Escape' && !accountMenu.hidden) {
    closeAccountMenu();
    accountMenuButton.focus();
  }
});

function initials(name) {
  return name.split(/\s+/).map(part => part[0]).slice(0, 2).join('').toUpperCase();
}

function renderPeople() {
  peopleList.replaceChildren();
  document.querySelector('#assignedCount').textContent = assignedPeople.length;

  assignedPeople.forEach(person => {
    const row = document.createElement('div');
    row.className = 'person-row';

    const identity = document.createElement('div');
    identity.className = 'person-identity';
    const avatar = document.createElement('span');
    avatar.className = 'person-avatar';
    avatar.textContent = initials(person.name);
    const details = document.createElement('span');
    const name = document.createElement('strong');
    name.textContent = person.name;
    const employeeId = document.createElement('small');
    employeeId.textContent = person.id;
    details.append(name, employeeId);
    identity.append(avatar, details);

    const controls = document.createElement('div');
    controls.className = 'person-controls';
    const role = document.createElement('span');
    role.className = 'role-badge';
    role.textContent = 'Administrator';

    if (person.initialAdmin) {
      const label = document.createElement('span');
      label.className = 'initial-admin';
      label.textContent = 'Initial admin';
      controls.append(role, label);
    } else {
      const remove = document.createElement('button');
      remove.className = 'remove-person';
      remove.type = 'button';
      remove.textContent = 'Remove';
      remove.addEventListener('click', () => {
        const index = assignedPeople.indexOf(person);
        assignedPeople.splice(index, 1);
        renderPeople();
        showToast(`${person.name} was removed.`);
      });
      controls.append(role, remove);
    }

    row.append(identity, controls);
    peopleList.append(row);
  });
}

personSearchForm.addEventListener('submit', event => {
  event.preventDefault();
  const query = personSearchInput.value.trim();
  if (!query) return;
  const normalized = query.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
  const displayName = query.replace(/\b\w/g, letter => letter.toUpperCase());
  pendingPerson = { id: normalized.startsWith('emp-') ? normalized.toUpperCase() : `EMP-${String(1048 + assignedPeople.length).padStart(4, '0')}`, name: displayName, role: 'Administrator', initialAdmin: false };
  document.querySelector('#searchAvatar').textContent = initials(displayName);
  document.querySelector('#searchName').textContent = displayName;
  document.querySelector('#searchEmployeeId').textContent = pendingPerson.id;
  personSearchResult.hidden = false;
});

document.querySelector('#addPersonButton').addEventListener('click', () => {
  if (!pendingPerson) return;
  const exists = assignedPeople.some(person => person.name.toLowerCase() === pendingPerson.name.toLowerCase());
  if (exists) {
    showToast('This person already has access.');
    return;
  }
  assignedPeople.push(pendingPerson);
  showToast(`${pendingPerson.name} was added as an Administrator.`);
  pendingPerson = null;
  personSearchResult.hidden = true;
  personSearchForm.reset();
  renderPeople();
});

document.querySelector('#newRequestButton').addEventListener('click', () => {
  modal.hidden = false;
  excelDropZone.focus();
});

function closeModal() {
  modal.hidden = true;
  pendingExcelFiles = [];
  excelFileInput.value = '';
  renderPendingExcelFiles();
  confirmExcelImport.disabled = true;
  confirmExcelImport.textContent = 'Import Record';
  importStatus.textContent = '';
  importStatus.classList.remove('error');
}

document.querySelector('#closeModal').addEventListener('click', closeModal);
modal.addEventListener('click', event => { if (event.target === modal) closeModal(); });

function normalizedHeader(value) {
  return String(value).split('/')[0].toLowerCase().replace(/[^a-z0-9]/g, '');
}

const excelFields = {
  prPreparer: ['prpreparer', 'requestor', 'requester'],
  prNumber: ['prno', 'prnumber'],
  prLineDescription: ['prlinedescription', 'description'],
  prCreationDate: ['prcreationdate'],
  prStatus: ['prstatus'],
  poNumber: ['pono', 'ponumber'],
  poCreationDate: ['pocreationdate'],
  needByDate: ['needbydate', 'expectedpodate'],
  servicePeriodStartDate: ['serviceperiodstartdate'],
  servicePeriodEndDate: ['serviceperiodenddate'],
  poLineStatus: ['polinestatus', 'postatus'],
  poLineAmount: ['polineamount', 'amount', 'amountusd'],
  receivedAmount: ['receivedamount'],
  billedAmount: ['billedamount']
};

function excelValue(row, aliases) {
  const entries = Object.entries(row);
  for (const alias of aliases) {
    const match = entries.find(([key]) => normalizedHeader(key) === alias);
    if (match) return match[1];
  }
  return '';
}

function amountValue(value) {
  if (typeof value === 'number') return value;
  const parsed = Number(String(value || '').replace(/[$,\s]/g, ''));
  return Number.isFinite(parsed) ? parsed : 0;
}

function rowsFromWorksheet(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
  const headerIndex = matrix.findIndex(row => {
    const headers = row.map(normalizedHeader);
    return headers.some(header => excelFields.prNumber.includes(header)) && headers.filter(header => Object.values(excelFields).some(aliases => aliases.includes(header))).length >= 2;
  });
  if (headerIndex < 0) return [];
  const headers = matrix[headerIndex];
  return matrix.slice(headerIndex + 1).map(values => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ''])));
}

function isoDate(value, fallback = '') {
  if (!value) return fallback;
  const text = String(value).trim();
  const dateMatch = text.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/);
  if (dateMatch) {
    const year = Number(dateMatch[1]);
    const month = Number(dateMatch[2]);
    const day = Number(dateMatch[3]);
    const checked = new Date(Date.UTC(year, month - 1, day));
    if (checked.getUTCFullYear() !== year || checked.getUTCMonth() !== month - 1 || checked.getUTCDate() !== day) return text;
    return `${dateMatch[1]}-${dateMatch[2].padStart(2, '0')}-${dateMatch[3].padStart(2, '0')}`;
  }
  const date = value instanceof Date ? value : new Date(text);
  if (Number.isNaN(date.getTime())) return String(value);
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
}

function trackingDates(needByDate, servicePeriodStartDate, servicePeriodEndDate) {
  const startDate = isoDate(servicePeriodStartDate);
  const endDate = isoDate(servicePeriodEndDate);
  const hasServicePeriod = Boolean(startDate || endDate);
  return {
    needByDate: hasServicePeriod ? '' : isoDate(needByDate),
    servicePeriodStartDate: startDate,
    servicePeriodEndDate: endDate
  };
}

async function importExcelFiles(files) {
  if (!files.length) return;
  confirmExcelImport.disabled = true;
  confirmExcelImport.textContent = 'Importing...';
  importStatus.classList.remove('error');
  importStatus.textContent = `Reading ${files.length} file${files.length === 1 ? '' : 's'}...`;

  try {
    if (!window.XLSX) throw new Error('The Excel reader could not be loaded. Check the internet connection and try again.');
    const imported = [];
    for (const file of files) {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
      workbook.SheetNames.forEach(sheetName => {
        rowsFromWorksheet(workbook.Sheets[sheetName]).forEach(row => {
          const value = field => excelValue(row, excelFields[field]);
          const prNumber = String(value('prNumber')).trim();
          if (!prNumber) return;
          const selectedDates = trackingDates(value('needByDate'), value('servicePeriodStartDate'), value('servicePeriodEndDate'));
          imported.push({
            prPreparer: String(value('prPreparer') || '—').trim(),
            prNumber,
            prLineDescription: String(value('prLineDescription') || '—').trim(),
            prCreationDate: isoDate(value('prCreationDate')),
            prStatus: String(value('prStatus') || '—').trim(),
            poNumber: String(value('poNumber') || '—').trim(),
            poCreationDate: isoDate(value('poCreationDate')),
            ...selectedDates,
            poLineStatus: String(value('poLineStatus') || '—').trim(),
            poLineAmount: amountValue(value('poLineAmount')),
            receivedAmount: amountValue(value('receivedAmount')),
            billedAmount: amountValue(value('billedAmount')),
            sourceFile: file.name,
            sourceSheet: sheetName
          });
        });
      });
    }

    if (!imported.length) throw new Error('No records were found. Make sure at least one worksheet includes the 14-column PR / PO template and a PR NO. value.');
    clearPreviousImportState();
    records.splice(0, records.length, ...imported);
    try {
      localStorage.setItem('pr-po-tracker-records', JSON.stringify({ revision: bundledDataRevision, records: imported }));
    } catch {
      // Keep the replacement data for this session if browser storage is unavailable.
    }
    render();
    showToast(`${imported.length} Excel record${imported.length === 1 ? '' : 's'} imported. Previous records were replaced.`);
    closeModal();
  } catch (error) {
    importStatus.classList.add('error');
    importStatus.textContent = error.message || 'Unable to read this Excel file.';
    confirmExcelImport.disabled = false;
    confirmExcelImport.textContent = 'Import Record';
  }
}

function renderPendingExcelFiles() {
  excelFileList.replaceChildren();
  excelFileList.hidden = pendingExcelFiles.length === 0;
  confirmExcelImport.disabled = pendingExcelFiles.length === 0;

  pendingExcelFiles.forEach((file, index) => {
    const row = document.createElement('div');
    row.className = 'excel-file-preview';

    const icon = document.createElement('span');
    icon.className = 'excel-file-icon';
    icon.setAttribute('aria-hidden', 'true');
    icon.innerHTML = '<svg viewBox="0 0 24 24"><path d="M5 3h10l4 4v14H5z"/><path d="M15 3v5h4"/><path d="m8 12 4 5M12 12l-4 5"/></svg>';

    const details = document.createElement('span');
    details.className = 'excel-file-details';
    const name = document.createElement('strong');
    name.textContent = file.name;
    const status = document.createElement('small');
    status.textContent = 'Ready to import';
    details.append(name, status);

    const remove = document.createElement('button');
    remove.className = 'remove-excel-file';
    remove.type = 'button';
    remove.setAttribute('aria-label', `Remove ${file.name}`);
    remove.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7h16M9 7V4h6v3M7 7l1 13h8l1-13M10 11v5M14 11v5"/></svg>';
    remove.addEventListener('click', () => {
      pendingExcelFiles.splice(index, 1);
      renderPendingExcelFiles();
    });

    row.append(icon, details, remove);
    excelFileList.append(row);
  });
}

function stageExcelFiles(files) {
  importStatus.classList.remove('error');
  const incoming = Array.from(files || []);
  const invalid = incoming.filter(file => !/\.(xlsx|xls|csv)$/i.test(file.name));
  const valid = incoming.filter(file => /\.(xlsx|xls|csv)$/i.test(file.name));

  valid.forEach(file => {
    const duplicate = pendingExcelFiles.some(existing => existing.name === file.name && existing.size === file.size && existing.lastModified === file.lastModified);
    if (!duplicate) pendingExcelFiles.push(file);
  });
  renderPendingExcelFiles();

  if (invalid.length) {
    importStatus.classList.add('error');
    importStatus.textContent = `${invalid.length} unsupported file${invalid.length === 1 ? '' : 's'} skipped.`;
    return;
  }
  importStatus.textContent = '';
}

['dragenter', 'dragover'].forEach(eventName => excelDropZone.addEventListener(eventName, event => {
  event.preventDefault();
  event.stopPropagation();
  excelDropZone.classList.add('drag-over');
  if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
}));

['dragleave', 'drop'].forEach(eventName => excelDropZone.addEventListener(eventName, event => {
  event.preventDefault();
  event.stopPropagation();
  excelDropZone.classList.remove('drag-over');
}));

excelDropZone.addEventListener('drop', event => {
  stageExcelFiles(event.dataTransfer?.files);
});

excelFileInput.addEventListener('change', () => {
  stageExcelFiles(excelFileInput.files);
});

confirmExcelImport.addEventListener('click', () => importExcelFiles([...pendingExcelFiles]));

renderPeople();
render();
let initialView = 'dashboard';
try {
  const storedView = localStorage.getItem('pr-po-current-view');
  if (['dashboard', 'requests', 'settings'].includes(storedView)) initialView = storedView;
} catch {
  // Use Dashboard when browser storage is unavailable.
}
showView(initialView, document.querySelector(`[data-view="${initialView}"]`));
document.documentElement.classList.remove('app-loading');
