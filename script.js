// Booking Number Tagger App
let bookings = [];
let selectedRows = new Set();
let taggers = ['Mark', 'Mays', 'Jerk'];
let currentTagger = 'Mark';
let filter = 'all';
let agentStats = {};

const taggerList = document.getElementById('tagger-list');
const addTaggerBtn = document.getElementById('add-tagger-btn');
const filterBtns = document.querySelectorAll('.filter');
const excelUpload = document.getElementById('excel-upload');
const tableHeader = document.getElementById('table-header');
const tableBody = document.getElementById('table-body');
const tagBtn = document.getElementById('tag-btn');
const assignName = document.getElementById('assign-name');
const agentStatsDiv = document.getElementById('agent-stats');
const agentSearch = document.getElementById('agent-search');
const modalOverlay = document.getElementById('modal-overlay');
const modalTaggerName = document.getElementById('modal-tagger-name');
const modalAddBtn = document.getElementById('modal-add-btn');
const modalCancelBtn = document.getElementById('modal-cancel-btn');
const downloadBtn = document.getElementById('download-btn');
const downloadUntaggedBtn = document.getElementById('download-untagged-btn');
const downloadAllBtn = document.getElementById('download-all-btn');

function renderTaggers() {
  taggerList.innerHTML = '';
  taggers.forEach(name => {
    const li = document.createElement('li');
    li.className = 'tagger' + (name === currentTagger ? ' active' : '');
    li.textContent = name;
    li.dataset.name = name;
    li.onclick = () => {
      currentTagger = name;
      renderTaggers();
      renderTable();
      renderStats();
    };
    taggerList.appendChild(li);
  });
}

addTaggerBtn.onclick = () => {
  if (modalOverlay) {
    modalOverlay.style.display = 'flex';
    modalTaggerName.value = '';
    setTimeout(() => modalTaggerName.focus(), 100);
  }
};
if (modalAddBtn && modalTaggerName && modalOverlay) {
  modalAddBtn.onclick = () => {
    const name = modalTaggerName.value.trim();
    if (name && !taggers.includes(name)) {
      taggers.push(name);
      renderTaggers();
      renderStats();
      modalOverlay.style.display = 'none';
    } else {
      modalTaggerName.style.border = '1.5px solid #e74c3c';
      setTimeout(() => { modalTaggerName.style.border = '1px solid #eebbc3'; }, 1200);
    }
  };
}
if (modalCancelBtn && modalOverlay) {
  modalCancelBtn.onclick = () => {
    modalOverlay.style.display = 'none';
  };
}
if (modalTaggerName) {
  modalTaggerName.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') modalAddBtn.click();
  });
}

filterBtns.forEach(btn => {
  btn.onclick = () => {
    filterBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    filter = btn.dataset.filter;
    renderTable();
  };
});

excelUpload.onchange = (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (json.length < 2) return;
    const headers = json[0];
    bookings = json.slice(1).map((row, i) => {
      const obj = { id: i, tagged: false, agent: '', row, undo: false, date: '', time: '' };
      headers.forEach((h, idx) => obj[h] = row[idx]);
      return obj;
    });
    selectedRows.clear();
    renderTable(headers);
    renderStats();
  };
  reader.readAsArrayBuffer(file);
};

function renderTable(headers) {
  if (!headers && bookings.length) headers = Object.keys(bookings[0]).filter(k => !['id','tagged','undo','row','taggedAt','agent','updatedAt','date','time'].includes(k));
  // Only add agent, date, and time columns if not already present
  if (bookings.length) {
    if (!headers.includes('agent')) headers.push('agent');
    if (!headers.includes('date')) headers.push('date');
    if (!headers.includes('time')) headers.push('time');
  }
  tableHeader.innerHTML = '';
  if (!headers) return;
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    tableHeader.appendChild(th);
  });
  const thAction = document.createElement('th');
  thAction.textContent = 'Action';
  tableHeader.appendChild(thAction);

  tableBody.innerHTML = '';
  let filtered = bookings;
  if (filter === 'untagged') filtered = bookings.filter(b => !b.tagged);
  if (filter === 'tagged') filtered = bookings.filter(b => b.tagged);

  filtered.forEach(b => {
    const tr = document.createElement('tr');
    if (b.tagged) tr.classList.add('tagged');
    if (selectedRows.has(b.id)) tr.classList.add('selected');
    headers.forEach(h => {
      const td = document.createElement('td');
      td.textContent = b[h] || '';
      td.style.fontSize = '0.92rem';
      td.onclick = () => {
        if (!b.tagged) {
          if (selectedRows.has(b.id)) selectedRows.delete(b.id);
          else selectedRows.add(b.id);
          renderTable(headers);
        }
      };
      tr.appendChild(td);
    });
    // Action cell
    const tdAction = document.createElement('td');
    tdAction.className = 'action-cell';
    if (b.tagged) {
      const undoBtn = document.createElement('button');
      undoBtn.textContent = 'Undo';
      undoBtn.disabled = false; // Always enabled for tagged rows
      undoBtn.onclick = () => {
        b.tagged = false;
        b.agent = '';
        b.undo = false;
        // b.updatedAt = new Date().toLocaleString(); // removed updatedAt
        selectedRows.delete(b.id);
        renderTable(headers);
        renderStats();
      };
      tdAction.appendChild(undoBtn);
    } else {
      tdAction.textContent = '-';
    }
    tr.appendChild(tdAction);
    tableBody.appendChild(tr);
  });
  tagBtn.disabled = selectedRows.size === 0 || !assignName.value.trim();
}

assignName.oninput = () => {
  tagBtn.disabled = selectedRows.size === 0 || !assignName.value.trim();
};

tagBtn.onclick = () => {
  const agent = assignName.value.trim();
  if (!agent) return;
  const now = new Date();
  const dateStr = now.toLocaleDateString();
  const timeStr = now.toLocaleTimeString();
  // Tag all selected rows, send each to Google Sheets individually
  bookings.forEach(b => {
    if (selectedRows.has(b.id) && !b.tagged) {
      b.tagged = true;
      b.agent = agent;
      b.undo = false;
      b.date = dateStr;
      b.time = timeStr;
      // Send each tagged booking individually
      fetch('https://script.google.com/macros/s/AKfycbx6gRWt6s5STKYZNBjLpslXeQRE-ipvLr_h4k5KEVRXZR__-o0KMb7ZPXGMXhhnAShk/exec', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          bookingNumber: b['Booking Number'] || b['bookingNumber'] || b['BookingNumber'] || b['booking_number'] || '',
          agent: b.agent,
          date: b.date,
          time: b.time
        })
      });
    }
  });
  selectedRows.clear();
  renderTable();
  renderStats();
};

function renderStats() {
  agentStats = {};
  bookings.forEach(b => {
    if (b.agent) {
      if (!agentStats[b.agent]) agentStats[b.agent] = { pending: 0, completed: 0 };
      if (b.tagged) agentStats[b.agent].completed++;
      else agentStats[b.agent].pending++;
    }
  });
  let html = '';
  const searchValue = agentSearch ? agentSearch.value.trim().toLowerCase() : '';
  if (searchValue) {
    const found = Object.keys(agentStats).filter(agent => agent.toLowerCase().includes(searchValue));
    if (found.length) {
      found.forEach(agent => {
        html += `<div><b>${agent}</b>: <span style="color:#7be495">${agentStats[agent].completed}</span> completed, <span style="color:#eebbc3">${agentStats[agent].pending}</span> pending</div>`;
      });
    } else {
      html = '<div>No agent found.</div>';
    }
  } else {
    Object.keys(agentStats).forEach(agent => {
      html += `<div><b>${agent}</b>: <span style="color:#7be495">${agentStats[agent].completed}</span> completed, <span style="color:#eebbc3">${agentStats[agent].pending}</span> pending</div>`;
    });
    if (!html) html = '<div>No tagged bookings yet.</div>';
  }
  agentStatsDiv.innerHTML = html;
}

if (agentSearch) {
  agentSearch.addEventListener('input', renderStats);
}

// Download button logic
if (downloadBtn) {
  downloadBtn.onclick = () => {
    const table = document.getElementById('booking-table');
    const headerCells = table.querySelectorAll('thead th');
    const bodyRows = table.querySelectorAll('tbody tr');
    if (!headerCells.length || !bodyRows.length) {
      alert('No bookings to download.');
      return;
    }
    // Get headers as shown in the table
    const headers = Array.from(headerCells).map(th => th.textContent);
    // Get only tagged rows as shown in the table
    const taggedRows = Array.from(bodyRows).filter(tr => tr.classList.contains('tagged'));
    const data = taggedRows.map(tr =>
      Array.from(tr.querySelectorAll('td')).slice(0, headers.length).map(td => td.textContent)
    );
    // Remove the last column if it's 'Action' (not needed in export)
    let exportHeaders = headers;
    let exportData = data;
    if (headers[headers.length-1].toLowerCase() === 'action') {
      exportHeaders = headers.slice(0, -1);
      exportData = data.map(row => row.slice(0, -1));
    }
    if (exportData.length === 0) {
      alert('No tagged bookings to download.');
      return;
    }
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([exportHeaders, ...exportData]);
    XLSX.utils.book_append_sheet(wb, ws, 'Tagged Bookings');
    XLSX.writeFile(wb, 'tagged_bookings.xlsx');
  };
}
// Download Untagged button logic
if (downloadUntaggedBtn) {
  downloadUntaggedBtn.onclick = () => {
    const table = document.getElementById('booking-table');
    const headerCells = table.querySelectorAll('thead th');
    const bodyRows = table.querySelectorAll('tbody tr');
    if (!headerCells.length || !bodyRows.length) {
      alert('No bookings to download.');
      return;
    }
    // Get headers as shown in the table
    const headers = Array.from(headerCells).map(th => th.textContent);
    // Get only untagged rows as shown in the table
    const untaggedRows = Array.from(bodyRows).filter(tr => !tr.classList.contains('tagged'));
    const data = untaggedRows.map(tr =>
      Array.from(tr.querySelectorAll('td')).slice(0, headers.length).map(td => td.textContent)
    );
    // Remove the last column if it's 'Action' (not needed in export)
    let exportHeaders = headers;
    let exportData = data;
    if (headers[headers.length-1].toLowerCase() === 'action') {
      exportHeaders = headers.slice(0, -1);
      exportData = data.map(row => row.slice(0, -1));
    }
    if (exportData.length === 0) {
      alert('No untagged bookings to download.');
      return;
    }
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([exportHeaders, ...exportData]);
    XLSX.utils.book_append_sheet(wb, ws, 'Untagged Bookings');
    XLSX.writeFile(wb, 'untagged_bookings.xlsx');
  };
}
// Download All Bookings button logic
if (downloadAllBtn) {
  downloadAllBtn.onclick = () => {
    const table = document.getElementById('booking-table');
    const headerCells = table.querySelectorAll('thead th');
    const bodyRows = table.querySelectorAll('tbody tr');
    if (!headerCells.length || !bodyRows.length) {
      alert('No bookings to download.');
      return;
    }
    // Get headers as shown in the table
    const headers = Array.from(headerCells).map(th => th.textContent);
    // Get all rows as shown in the table
    const allData = Array.from(bodyRows).map(tr =>
      Array.from(tr.querySelectorAll('td')).slice(0, headers.length).map(td => td.textContent)
    );
    // Get tagged and untagged rows
    let exportHeaders = headers;
    let taggedData = [];
    let untaggedData = [];
    if (headers[headers.length-1].toLowerCase() === 'action') {
      exportHeaders = headers.slice(0, -1);
      taggedData = Array.from(bodyRows).filter(tr => tr.classList.contains('tagged')).map(tr =>
        Array.from(tr.querySelectorAll('td')).slice(0, headers.length-1).map(td => td.textContent)
      );
      untaggedData = Array.from(bodyRows).filter(tr => !tr.classList.contains('tagged')).map(tr =>
        Array.from(tr.querySelectorAll('td')).slice(0, headers.length-1).map(td => td.textContent)
      );
    } else {
      taggedData = Array.from(bodyRows).filter(tr => tr.classList.contains('tagged')).map(tr =>
        Array.from(tr.querySelectorAll('td')).map(td => td.textContent)
      );
      untaggedData = Array.from(bodyRows).filter(tr => !tr.classList.contains('tagged')).map(tr =>
        Array.from(tr.querySelectorAll('td')).map(td => td.textContent)
      );
    }
    const wb = XLSX.utils.book_new();
    const wsTagged = XLSX.utils.aoa_to_sheet([exportHeaders, ...taggedData]);
    XLSX.utils.book_append_sheet(wb, wsTagged, 'Tagged');
    const wsUntagged = XLSX.utils.aoa_to_sheet([exportHeaders, ...untaggedData]);
    XLSX.utils.book_append_sheet(wb, wsUntagged, 'Untagged');
    XLSX.writeFile(wb, 'all_bookings.xlsx');
  };
}

// Show current date and time above taggers
function updateCurrentDatetime() {
  const dtDiv = document.getElementById('current-datetime');
  if (dtDiv) {
    const now = new Date();
    dtDiv.textContent = now.toLocaleString();
  }
}
setInterval(updateCurrentDatetime, 1000);
updateCurrentDatetime();

// Initial render
renderTaggers();
renderStats();
