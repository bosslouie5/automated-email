// Minimal UI interactions: attachments drag/drop, preview, fake send animation
const attachmentsArea = document.getElementById('attachmentsArea');
const fileInput = document.getElementById('fileInput');
const browseBtn = document.getElementById('browseBtn');
const fileList = document.getElementById('fileList');
const emailForm = document.getElementById('emailForm');
const excelInput = document.getElementById('excelInput');
const excelBrowseBtn = document.getElementById('excelBrowseBtn');
const excelName = document.getElementById('excelName');
const excelSummary = document.getElementById('excelSummary');
const absentList = document.getElementById('absentList');
const subjectInput = document.getElementById('subject');
const recipientsInput = document.getElementById('recipients');

// Settings Modal Elements
const settingsBtn = document.getElementById('settingsBtn');
const settingsModal = document.getElementById('settingsModal');
const closeSettingsBtn = document.getElementById('closeSettingsBtn');
const closeSettingsFooterBtn = document.getElementById('closeSettingsFooterBtn');
const saveSettingsBtn = document.getElementById('saveSettingsBtn');
const defaultEmailSelect = document.getElementById('defaultEmailSelect');
const connectOutlookBtn = document.getElementById('connectOutlookBtn');
const addManualEmailBtn = document.getElementById('addManualEmailBtn');
const manualEmailInput = document.getElementById('manualEmailInput');
const outlookAccountsList = document.getElementById('outlookAccountsList');

// Auto-Detect Elements
const autoDetectBtn = document.getElementById('autoDetectBtn');
const autoDetectLoading = document.getElementById('autoDetectLoading');
const autoDetectResults = document.getElementById('autoDetectResults');
const detectedEmailsList = document.getElementById('detectedEmailsList');
const confirmDetectedBtn = document.getElementById('confirmDetectedBtn');
const cancelDetectBtn = document.getElementById('cancelDetectBtn');

// Store detected emails temporarily
let detectedEmails = [];
let selectedDetectedEmails = [];

// Email Configuration Storage
let emailAccounts = [];
let defaultEmail = '';

// Load saved email settings from localStorage
function loadEmailSettings(){
  const saved = localStorage.getItem('emailSettings');
  if(saved){
    try{
      const data = JSON.parse(saved);
      emailAccounts = data.accounts || [];
      defaultEmail = data.defaultEmail || '';
      updateEmailSelectOptions();
      renderOutlookAccountsList();
    } catch(err){
      console.error('Error loading email settings:', err);
    }
  }
}

// Save email settings to localStorage
function saveEmailSettings(){
  localStorage.setItem('emailSettings', JSON.stringify({
    accounts: emailAccounts,
    defaultEmail: defaultEmail
  }));
}

// Update email select dropdown
function updateEmailSelectOptions(){
  defaultEmailSelect.innerHTML = '<option value="">Select an email...</option>';
  emailAccounts.forEach(account => {
    const option = document.createElement('option');
    option.value = account.email;
    option.textContent = account.email;
    if(account.email === defaultEmail) option.selected = true;
    defaultEmailSelect.appendChild(option);
  });
}

// Render Outlook accounts list
function renderOutlookAccountsList(){
  if(emailAccounts.length === 0){
    outlookAccountsList.innerHTML = `
      <div class="empty-state">
        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M3 8v10a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V8M3 8l9-5 9 5M9 12h6M9 16h4"></path>
        </svg>
        <p>No accounts connected</p>
      </div>`;
    return;
  }
  
  outlookAccountsList.innerHTML = emailAccounts.map((account, idx) => `
    <div class="account-item">
      <div class="account-item-info">
        <div class="account-item-icon">${account.email.charAt(0).toUpperCase()}</div>
        <div class="account-item-details">
          <div class="account-item-email">${escapeHtml(account.email)}</div>
          <div class="account-item-status">${account.source === 'outlook' ? '✓ Outlook Connected' : 'Manual Entry'}</div>
        </div>
      </div>
      <button class="account-item-remove" data-index="${idx}" type="button">Remove</button>
    </div>
  `).join('');
  
  outlookAccountsList.classList.add('has-accounts');
  document.querySelectorAll('.account-item-remove').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const idx = parseInt(e.target.dataset.index);
      removeEmailAccount(idx);
    });
  });
}

// Add email account
function addEmailAccount(email, source = 'manual'){
  if(!email || !email.includes('@')) return false;
  if(emailAccounts.some(acc => acc.email === email)) return false;
  
  emailAccounts.push({ email, source, addedDate: new Date().toISOString() });
  updateEmailSelectOptions();
  renderOutlookAccountsList();
  saveEmailSettings();
  return true;
}

// Remove email account
function removeEmailAccount(index){
  if(index >= 0 && index < emailAccounts.length){
    emailAccounts.splice(index, 1);
    if(defaultEmail === emailAccounts[index]?.email) defaultEmail = '';
    updateEmailSelectOptions();
    renderOutlookAccountsList();
    saveEmailSettings();
  }
}

// Simulate Outlook integration
function connectOutlookAccount(){
  // Simulate fetching Outlook accounts (in real scenario, this would use MS Graph API)
  const mockOutlookEmails = [
    'john.doe@company.com',
    'john.d@company.outlook.com'
  ];
  
  // Show a selection dialog (simulated)
  let selected = prompt(`Available Outlook accounts:\n\n${mockOutlookEmails.join('\n')}\n\nEnter the email you want to add (or click Cancel):`, mockOutlookEmails[0]);
  
  if(selected && selected.trim()){
    selected = selected.trim();
    if(addEmailAccount(selected, 'outlook')){
      showNotification('Outlook account connected successfully!', 'success');
    } else {
      showNotification('Account already exists or invalid email.', 'error');
    }
  }
}

// Auto-Detect Outlook Emails
function startAutoDetect(){
  // Show loading state
  autoDetectLoading.classList.remove('hidden');
  autoDetectResults.classList.add('hidden');
  detectedEmails = [];
  selectedDetectedEmails = [];
  
  // Simulate network delay for realistic UX
  setTimeout(() => {
    // Simulate fetching from Outlook/system (in real scenario, use MS Graph API)
    detectedEmails = [
      { email: 'john.doe@company.com', type: 'outlook', status: 'active' },
      { email: 'john.d@company.outlook.com', type: 'outlook', status: 'active' },
      { email: 'j.doe@mail.company.com', type: 'outlook', status: 'active' }
    ];
    
    // Filter out already added emails
    detectedEmails = detectedEmails.filter(d => !emailAccounts.some(acc => acc.email === d.email));
    
    if(detectedEmails.length === 0){
      autoDetectLoading.classList.add('hidden');
      showNotification('No new Outlook accounts found. All accounts are already added.', 'info');
      return;
    }
    
    // Hide loading and show results
    autoDetectLoading.classList.add('hidden');
    renderDetectedEmails();
    autoDetectResults.classList.remove('hidden');
  }, 1200);
}

function renderDetectedEmails(){
  detectedEmailsList.innerHTML = detectedEmails.map((email, idx) => `
    <div class="detected-email-item" style="animation: slideInUp 0.3s ease ${idx * 0.08}s backwards">
      <input type="checkbox" id="detect-${idx}" class="email-checkbox" data-email="${email.email}" data-index="${idx}" checked>
      <label for="detect-${idx}" class="email-checkbox-label">
        <div class="email-checkbox-content">
          <div class="email-checkbox-main">${escapeHtml(email.email)}</div>
          <div class="email-checkbox-status">
            <span class="status-badge outlook">Outlook</span>
          </div>
        </div>
      </label>
    </div>
  `).join('');
  
  // Add event listeners for checkboxes
  document.querySelectorAll('.email-checkbox').forEach(checkbox => {
    checkbox.addEventListener('change', (e) => {
      const email = e.target.dataset.email;
      if(e.target.checked){
        if(!selectedDetectedEmails.includes(email)){
          selectedDetectedEmails.push(email);
        }
      } else {
        selectedDetectedEmails = selectedDetectedEmails.filter(e => e !== email);
      }
    });
    // Pre-select all by default
    selectedDetectedEmails.push(checkbox.dataset.email);
  });
}

function confirmDetectedEmails(){
  if(selectedDetectedEmails.length === 0){
    showNotification('Please select at least one email.', 'error');
    return;
  }
  
  let addedCount = 0;
  selectedDetectedEmails.forEach(email => {
    if(addEmailAccount(email, 'outlook')){
      addedCount++;
    }
  });
  
  if(addedCount > 0){
    showNotification(`Successfully added ${addedCount} Outlook account${addedCount > 1 ? 's' : ''}!`, 'success');
  }
  
  // Reset UI
  autoDetectResults.classList.add('hidden');
  detectedEmails = [];
  selectedDetectedEmails = [];
}

function cancelAutoDetect(){
  autoDetectLoading.classList.add('hidden');
  autoDetectResults.classList.add('hidden');
  detectedEmails = [];
  selectedDetectedEmails = [];
}

// Show notification
function showNotification(message, type = 'info'){
  const notification = document.createElement('div');
  notification.style.cssText = `
    position: fixed;
    top: 20px;
    right: 20px;
    padding: 14px 18px;
    background: ${type === 'success' ? 'rgba(45, 212, 191, 0.15)' : type === 'error' ? 'rgba(255, 120, 120, 0.15)' : 'rgba(107, 91, 255, 0.15)'};
    border: 1px solid ${type === 'success' ? 'rgba(45, 212, 191, 0.4)' : type === 'error' ? 'rgba(255, 120, 120, 0.4)' : 'rgba(107, 91, 255, 0.4)'};
    color: ${type === 'success' ? '#2dd4bf' : type === 'error' ? '#ff7878' : 'rgba(255, 255, 255, 0.8)'};
    border-radius: 10px;
    font-size: 13px;
    z-index: 10000;
    animation: slideIn 0.3s ease;
  `;
  notification.textContent = message;
  document.body.appendChild(notification);
  
  setTimeout(() => {
    notification.style.animation = 'slideOut 0.3s ease';
    setTimeout(() => notification.remove(), 300);
  }, 3000);
}

// Settings Modal Event Listeners
settingsBtn.addEventListener('click', () => {
  settingsModal.classList.remove('hidden');
});

closeSettingsBtn.addEventListener('click', () => {
  settingsModal.classList.add('hidden');
});

closeSettingsFooterBtn.addEventListener('click', () => {
  settingsModal.classList.add('hidden');
});

const modalOverlay = document.querySelector('.modal-overlay');
if(modalOverlay){
  modalOverlay.addEventListener('click', () => {
    settingsModal.classList.add('hidden');
  });
}

saveSettingsBtn.addEventListener('click', () => {
  defaultEmail = defaultEmailSelect.value;
  saveEmailSettings();
  showNotification('Settings saved successfully!', 'success');
  setTimeout(() => {
    settingsModal.classList.add('hidden');
  }, 500);
});

connectOutlookBtn.addEventListener('click', connectOutlookAccount);

autoDetectBtn.addEventListener('click', startAutoDetect);
confirmDetectedBtn.addEventListener('click', confirmDetectedEmails);
cancelDetectBtn.addEventListener('click', cancelAutoDetect);

addManualEmailBtn.addEventListener('click', () => {
  const email = manualEmailInput.value.trim();
  if(email){
    if(addEmailAccount(email, 'manual')){
      manualEmailInput.value = '';
      showNotification('Email added successfully!', 'success');
    } else {
      showNotification('Invalid email or account already exists.', 'error');
    }
  } else {
    showNotification('Please enter a valid email address.', 'error');
  }
});

// Load settings on page load
loadEmailSettings();

function humanSize(bytes){
  if(bytes<1024) return bytes+' B';
  if(bytes<1024*1024) return (bytes/1024).toFixed(1)+' KB';
  return (bytes/(1024*1024)).toFixed(2)+' MB';
}

browseBtn.addEventListener('click', ()=>fileInput.click());
attachmentsArea.addEventListener('click', ()=>fileInput.click());
excelBrowseBtn.addEventListener('click', ()=>excelInput.click());

attachmentsArea.addEventListener('dragover', e=>{e.preventDefault();attachmentsArea.classList.add('dragover')});
attachmentsArea.addEventListener('dragleave', e=>{attachmentsArea.classList.remove('dragover')});
attachmentsArea.addEventListener('drop', e=>{
  e.preventDefault();attachmentsArea.classList.remove('dragover');
  const files = Array.from(e.dataTransfer.files||[]);
  handleFiles(files);
});

fileInput.addEventListener('change', e=>handleFiles(Array.from(e.target.files)));
excelInput.addEventListener('change', e=>{
  const file = e.target.files && e.target.files[0];
  if(file) parseExcelFile(file);
});

function handleFiles(files){
  files.forEach(f=>{
    const li = document.createElement('li');
    li.innerHTML = `<div><strong>${escapeHtml(f.name)}</strong> <span style="color:rgba(255,255,255,0.6);font-size:12px;margin-left:8px">${humanSize(f.size)}</span></div><button class="btn ghost remove">Remove</button>`;
    fileList.appendChild(li);
    li.querySelector('.remove').addEventListener('click', ()=>li.remove());
  });
}

function parseExcelFile(file){
  excelName.textContent = file.name;
  const reader = new FileReader();
  // Choose read method based on file extension; CSV is handled as text for reliability
  const ext = (file.name || '').split('.').pop().toLowerCase();
  reader.onload = e => {
    try {
      let workbook;
      if(ext === 'csv'){
        // CSV as string
        const text = e.target.result;
        workbook = XLSX.read(text, { type: 'string' });
      } else {
        // Excel binary formats
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: 'array' });
      }

      const sheetName = workbook.SheetNames && workbook.SheetNames[0];
      if(!sheetName){
        throw new Error('No sheets found in workbook');
      }
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      const summary = createExcelSummary(rows);
      showExcelSummary(summary, file.name);
      showAbsentList(summary.absentRows);
    } catch (err) {
      // Try a safe fallback: if initial parsing failed and we didn't try CSV path, attempt text parse
      if(ext !== 'csv'){
        try {
          const textFallbackReader = new FileReader();
          textFallbackReader.onload = te => {
            try {
              const workbook = XLSX.read(te.target.result, { type: 'string' });
              const sheetName = workbook.SheetNames && workbook.SheetNames[0];
              const sheet = workbook.Sheets[sheetName];
              const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
              const summary = createExcelSummary(rows);
              showExcelSummary(summary, file.name + ' (parsed as text fallback)');
              showAbsentList(summary.absentRows);
              return;
            } catch (err2) {
              console.error('Fallback parse failed', err2);
            }
          };
          textFallbackReader.readAsText(file);
          return;
        } catch (fbErr){
          console.error('Fallback attempt failed', fbErr);
        }
      }

      showExcelError('Unable to parse Excel file. Please use a valid .xlsx, .xls, or .csv format.');
      console.error(err);
    }
  };

  if(ext === 'csv'){
    reader.readAsText(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

function createExcelSummary(rows){
  const normalizedRows = rows.map(row => {
    const normalized = {};
    Object.keys(row).forEach(key => {
      normalized[key.trim().toLowerCase()] = row[key];
    });
    return normalized;
  });

  const firstRow = normalizedRows[0] || {};
  console.log('First Row Keys:', Object.keys(firstRow));
  
  const codeField = findField(firstRow, ['code', 'employee code', 'emp code', 'id', 'employee id']);
  const nameField = findField(firstRow, ['name', 'employee name', 'full name', 'employee']);
  const designationField = findField(firstRow, ['designation', 'job title', 'position', 'role']);
  const departmentField = findField(firstRow, ['department', 'dept', 'team']);
  const dateField = findField(firstRow, ['date', 'day', 'attendance date', 'clock date']);
  const emailField = findField(firstRow, ['email', 'e-mail', 'work email', 'employee email', 'recipient']);
  const notesField = findField(firstRow, ['notes for emails', 'notes for email', 'notes', 'remarks', 'comments']);
  let statusField = findField(firstRow, ['status', 'attendance', 'presence', 'absent', 'day status']);
  
  console.log('Detected Fields - Code:', codeField, 'Status:', statusField, 'Notes:', notesField, 'Total Rows:', rows.length);

  const code = codeField ? firstRow[codeField] : '';
  const name = nameField ? firstRow[nameField] : '';

  /* Strict absent filter: only accept rows where the status cell is exactly '-' or 'A' (case-insensitive).
     Build a sanitized list limited to the required columns: Code, Name, Designation, Department, Date, Status. */
  let absenteeRows = normalizedRows.reduce((acc, row) => {
    if(!statusField) return acc; // no explicit status column, skip strict filtering
    const raw = row[statusField];
    if(raw === undefined || raw === null) return acc;
    const v = String(raw).trim();
    const noteText = notesField ? String(row[notesField] || '').trim().toLowerCase() : '';
    if(noteText.includes("don't send email") || noteText.includes('dont send email') || noteText.includes('do not send email')) return acc;
    if(v === '-' || v.toUpperCase() === 'A'){
      acc.push({
        code: String(row[codeField] || row['code'] || row['employee code'] || row.id || row['employee id'] || '').trim(),
        name: String(row[nameField] || row['name'] || row['employee name'] || row['full name'] || row.employee || '').trim(),
        designation: String(row[designationField] || row['designation'] || row['job title'] || row.position || '').trim(),
        department: String(row[departmentField] || row['department'] || row.dept || row.team || '').trim(),
        date: String(row[dateField] || row['date'] || row['attendance date'] || '').trim(),
        status: 'Absent'
      });
    }
    return acc;
  }, []);
  
  console.log('Absent Rows Found:', absenteeRows.length, absenteeRows);

  const uniqueRecipients = emailField ? Array.from(new Set(normalizedRows.map(row => String(row[emailField]).trim()).filter(Boolean))) : [];

  // If no explicit status column found, try to auto-detect the most likely column
  // that contains '-' or 'A' entries (handles varied header names).
  if(!statusField){
    const dashChars = ['-','\u2013','\u2014'];
    const keyScores = {};
    Object.keys(firstRow).forEach(k=>keyScores[k]=0);
    normalizedRows.forEach(r=>{
      Object.keys(r).forEach(k=>{
        const raw = r[k];
        if(raw===undefined || raw===null) return;
        const s = String(raw).trim();
        const up = s.toUpperCase();
        if(s === '-' || up === 'A' || dashChars.includes(s) || s === '\u2012'){
          keyScores[k] = (keyScores[k]||0) + 1;
        }
      });
    });
    const sorted = Object.keys(keyScores).sort((a,b)=> (keyScores[b]||0) - (keyScores[a]||0));
    if(sorted.length && (keyScores[sorted[0]]||0) > 0){
      statusField = sorted[0];
    }
  }

  // capture a small sample for debugging / preview in the UI
  const sampleRows = normalizedRows.slice(0,6).map(r => {
    const obj = {};
    Object.keys(r).forEach(k=> obj[k] = r[k]);
    return obj;
  });

  // Fallback: if no absentee rows found yet, scan all fields for '-' or 'A' and include those rows.
  if((!absenteeRows || absenteeRows.length === 0) && normalizedRows.length){
    const fallback = [];
    normalizedRows.forEach(row => {
      const noteText = notesField ? String(row[notesField] || '').trim().toLowerCase() : '';
      if(noteText.includes("don't send email") || noteText.includes('dont send email') || noteText.includes('do not send email')) return;
      const anyMatch = Object.values(row).some(v => {
        if(v === undefined || v === null) return false;
        const s = String(v).trim();
        return s === '-' || s.toUpperCase() === 'A';
      });
      if(anyMatch){
        fallback.push({
          code: String(row[codeField] || row['code'] || row['employee code'] || row.id || row['employee id'] || '').trim(),
          name: String(row[nameField] || row['name'] || row['employee name'] || row['full name'] || row.employee || '').trim(),
          designation: String(row[designationField] || row['designation'] || row['job title'] || row.position || '').trim(),
          department: String(row[departmentField] || row['department'] || row.dept || row.team || '').trim(),
          date: String(row[dateField] || row['date'] || row['attendance date'] || '').trim(),
          status: 'Absent'
        });
      }
    });
    if(fallback.length) absenteeRows = fallback;
  }

  if(code && name){
    subjectInput.value = `Absence Report For ${code} - ${name}`;
  }

  if(uniqueRecipients.length){
    recipientsInput.value = uniqueRecipients.join(', ');
  }

  return {
    fileName: excelName.textContent,
    rows: rows.length,
    absentCount: absenteeRows.length,
    absentRows: absenteeRows,
    code: code || 'Not found',
    name: name || 'Not found',
    recipients: uniqueRecipients,
    detectedFields: {
      codeField, nameField, designationField, departmentField, dateField, emailField, statusField
    },
    sampleRows,
    missingFields: {
      code: !code,
      name: !name,
      email: !emailField,
      status: !statusField
    }
  };
}

function findField(row, candidates){
  return Object.keys(row).find(key => candidates.includes(key.trim().toLowerCase()));
}

function isAbsentRow(row, statusField){
  if(statusField && typeof row[statusField] === 'string'){
    const value = row[statusField].toLowerCase();
    return /absent|a|no|missing|not present/.test(value) || /late|leave|sick/.test(value);
  }
  const values = Object.values(row).map(v => String(v).toLowerCase());
  return values.some(v => /absent|missing|no|not present/.test(v));
}

function showExcelSummary(summary, fileName){
  // Keep the summary data available internally, but hide the visible summary panel.
  excelSummary.classList.add('hidden');
}

function renderSampleTable(sampleRows, statusField){
  if(!Array.isArray(sampleRows) || sampleRows.length===0) return '<div style="opacity:.7">No sample rows available.</div>';
  const headers = Object.keys(sampleRows[0]).slice(0,6);
  const head = `<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:6px">${headers.map(h=>`<div style="min-width:110px;font-weight:700">${escapeHtml(h)}</div>`).join('')}</div>`;
  const rows = sampleRows.map(r=>{
    return `<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:6px">${headers.map(h=>`<div style="min-width:110px;opacity:.9">${escapeHtml(String(r[h]||''))}</div>`).join('')}</div>`;
  }).join('');
  const statusPreview = statusField ? `<div style="margin-top:6px;font-size:13px">Status field used: <strong>${escapeHtml(statusField)}</strong></div>` : '<div style="margin-top:6px;font-size:13px;opacity:.8">No status field detected yet.</div>';
  return head + rows + statusPreview;
}

function showExcelError(message){
  excelSummary.innerHTML = `<div class="summary-card"><strong>Error</strong><p>${escapeHtml(message)}</p></div>`;
  excelSummary.classList.remove('hidden');
  absentList.classList.add('hidden');
}

function showAbsentList(absentRows){
  console.log('showAbsentList called with rows:', absentRows);
  // Always show the list element
  absentList.classList.remove('hidden');
  
  if(!Array.isArray(absentRows) || absentRows.length === 0){
    absentList.innerHTML = `
      <h4>Absence List</h4>
      <div style="padding:20px;text-align:center;color:var(--muted)">
        <p>✓ No absences found in the data</p>
        <p style="font-size:13px;margin-top:10px">All employees appear to be present on the selected dates.</p>
      </div>`;
    return;
  }

  // Render a simple table with only the required columns
  const headers = ['Code','Name','Designation','Department','Date','Status'];
  const rowsHtml = absentRows.map(r => `
    <tr>
      <td>${escapeHtml(String(r.code || ''))}</td>
      <td>${escapeHtml(String(r.name || ''))}</td>
      <td>${escapeHtml(String(r.designation || ''))}</td>
      <td>${escapeHtml(String(r.department || ''))}</td>
      <td>${escapeHtml(String(r.date || ''))}</td>
      <td>${escapeHtml(String(r.status || 'Absent'))}</td>
    </tr>`).join('');

  absentList.innerHTML = `
    <h4>Absence List</h4>
    <div class="table-wrap">
      <table class="absent-table">
        <thead><tr>${headers.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>`;
}

function escapeHtml(s){return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}

emailForm.addEventListener('submit', e=>{
  e.preventDefault();
  
  // Check if default email is set
  if(!defaultEmail){
    showNotification('Please configure an email account in Settings first.', 'error');
    return;
  }
  
  const sendBtn = document.getElementById('sendBtn');
  sendBtn.disabled = true; sendBtn.textContent = 'Sending...';
  sendBtn.animate([{transform:'scale(1)'},{transform:'scale(.98)'}],{duration:120,iterations:1});
  setTimeout(()=>{
    sendBtn.textContent = 'Sent ✓';
    sendBtn.classList.add('sent');
    showNotification(`Email sent from ${defaultEmail}`, 'success');
    setTimeout(()=>{sendBtn.disabled=false;sendBtn.textContent='Send';sendBtn.classList.remove('sent')},1600);
  },1400);
});

// Add slideOut animation
const style = document.createElement('style');
style.textContent = `
  @keyframes slideOut{
    from{transform:translateX(0);opacity:1}
    to{transform:translateX(400px);opacity:0}
  }
`;
document.head.appendChild(style);
