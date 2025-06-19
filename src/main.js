import './style.css';

// Replace with your Power Automate flow endpoint
const FLOW_URL = 'https://prod-24.southafricanorth.logic.azure.com:443/workflows/8d3e1ec3e5b649d8ba1637861dd6a2c4/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=iINQOHWvmhaSZJRMsGKhqasw4WR6bzfGGEQrZneYmpE';

function getStartOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day;
  const weekStart = new Date(d.setDate(diff));
  weekStart.setHours(0,0,0,0);
  return weekStart;
}

function getWeekDays(startDate) {
  const days = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(startDate);
    d.setDate(d.getDate() + i);
    days.push(d);
  }
  return days;
}

function isDateInWeek(date, weekStart) {
  const d = new Date(date);
  d.setHours(0,0,0,0);
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 7);
  return d >= weekStart && d < weekEnd;
}

function createCalendarGrid(weekStart) {
  const calendar = document.createElement('div');
  calendar.id = 'calendar';
  calendar.innerHTML = '<h2>Week View</h2>';
  calendar.style.overflowX = 'auto';

  const table = document.createElement('table');
  table.id = 'calendar-table';
  table.style.minWidth = '1100px';
  table.style.width = 'auto';

  // Header row
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  headerRow.innerHTML = '<th>Time</th>';
  const weekDays = getWeekDays(weekStart);
  weekDays.forEach(day => {
    headerRow.innerHTML += `<th>${day.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric' })}</th>`;
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // 30-min slots (07:00–20:00)
  const tbody = document.createElement('tbody');
  for (let hour = 7; hour <= 20; hour++) {
    for (let half = 0; half < 2; half++) {
      const row = document.createElement('tr');
      const timeLabel = `${hour.toString().padStart(2, '0')}:${half === 0 ? '00' : '30'}`;
      row.innerHTML = `<td>${timeLabel}</td>`;
      for (let d = 0; d < 7; d++) {
        row.innerHTML += `<td data-day="${d}" data-slot="${(hour - 7) * 2 + half}"></td>`;
      }
      tbody.appendChild(row);
    }
  }
  table.appendChild(tbody);
  calendar.appendChild(table);
  document.body.appendChild(calendar);
  return { weekStart, weekDays };
}

async function fetchMeetings() {
  try {
    const response = await fetch(FLOW_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({}) // send empty object or add payload if needed
    });
    if (!response.ok) throw new Error('Failed to fetch meetings');
    const data = await response.json();
    // Return an array of { calendarIndex, meeting } for color coding
    if (Array.isArray(data)) {
      return data.flatMap((cal, idx) =>
        Array.isArray(cal.value) ? cal.value.map(meeting => ({ ...meeting, calendarIndex: idx })) : []
      );
    }
    return [];
  } catch (e) {
    console.error(e);
    return [];
  }
}

function renderMeetings(meetings, weekStart) {
  // Clear all cells
  document.querySelectorAll('#calendar-table td[data-day]').forEach(cell => {
    cell.innerHTML = '';
    cell.style.display = '';
    cell.removeAttribute('rowSpan');
    cell.style.minWidth = '';
    cell.style.position = '';
    cell.style.verticalAlign = '';
  });
  // Up to 20 visually distinct colors for calendars
  const colors = [
    '#e3f0ff', '#ffe3e3', '#e3ffe7', '#fffbe3', '#f3e3ff', '#e3fff9', '#ffe3f7', '#e3eaff', '#fbe3ff', '#e3fff0',
    '#ffe9e3', '#e3fff6', '#e3f7ff', '#f7ffe3', '#e3e3ff', '#fff3e3', '#e3fff3', '#ffe3f0', '#e3f9ff', '#f0ffe3'
  ];
  const borders = [
    '#1976d2', '#d21919', '#19d26b', '#bfa100', '#7c19d2', '#19bfae', '#d2198a', '#195ad2', '#b819d2', '#19d25a',
    '#d26a19', '#19d2b7', '#199ad2', '#a1d219', '#6a19d2', '#d2b719', '#19d2a1', '#d2196a', '#19d2f9', '#6ad219'
  ];
  // Group meetings by day
  const meetingsByDay = Array.from({ length: 7 }, () => []);
  meetings.forEach(meeting => {
    // Use startWithTimeZone and endWithTimeZone if present, else fallback
    const startStr = meeting.startWithTimeZone || meeting.start?.dateTime || meeting.start;
    const endStr = meeting.endWithTimeZone || meeting.end?.dateTime || meeting.end;
    const start = new Date(startStr);
    const end = new Date(endStr);
    const dayIdx = (start.getDay() + 7 - weekStart.getDay()) % 7;
    meetingsByDay[dayIdx].push({ meeting, start, end });
  });
  // Calculate max overlaps per day
  const maxOverlaps = meetingsByDay.map(dayMeetings => {
    // Build an array of all time slots (26 slots: 7:00-20:00, 30min each)
    const slots = Array(26).fill(0);
    dayMeetings.forEach(({ start, end }) => {
      let startSlot = (start.getHours() - 7) * 2 + (start.getMinutes() >= 30 ? 1 : 0);
      let endSlot = (end.getHours() - 7) * 2 + (end.getMinutes() > 0 ? (end.getMinutes() > 30 ? 2 : 1) : 0);
      if (end.getMinutes() === 0 && end.getSeconds() === 0) endSlot = (end.getHours() - 7) * 2;
      if (end.getMinutes() === 30 && end.getSeconds() === 0) endSlot = (end.getHours() - 7) * 2 + 1;
      if (endSlot <= startSlot) endSlot = startSlot + 1;
      for (let i = startSlot; i < endSlot; i++) slots[i]++;
    });
    return Math.max(1, ...slots);
  });
  // Set min-width for each day column
  for (let d = 0; d < 7; d++) {
    const th = document.querySelector(`#calendar-table th:nth-child(${d + 2})`);
    if (th) th.style.minWidth = `${maxOverlaps[d] * 120}px`;
    for (let slot = 0; slot < 26; slot++) {
      const td = document.querySelector(`#calendar-table td[data-day="${d}"][data-slot="${slot}"]`);
      if (td) {
        td.style.minWidth = `${maxOverlaps[d] * 120}px`;
        td.style.position = 'relative';
        td.style.verticalAlign = 'top';
        td.style.padding = '0';
        td.innerHTML = '<div class="meeting-row" style="display:flex;position:relative;height:100%;width:100%;"></div>';
      }
    }
  }
  // Place meetings in correct slots, side by side
  meetingsByDay.forEach((dayMeetings, dayIdx) => {
    // Sort by start time
    dayMeetings.sort((a, b) => a.start - b.start);
    // For each meeting, find its slot and add to the flex row
    dayMeetings.forEach(({ meeting, start, end }) => {
      let startSlot = (start.getHours() - 7) * 2 + (start.getMinutes() >= 30 ? 1 : 0);
      let endSlot = (end.getHours() - 7) * 2 + (end.getMinutes() > 0 ? (end.getMinutes() > 30 ? 2 : 1) : 0);
      if (end.getMinutes() === 0 && end.getSeconds() === 0) endSlot = (end.getHours() - 7) * 2;
      if (end.getMinutes() === 30 && end.getSeconds() === 0) endSlot = (end.getHours() - 7) * 2 + 1;
      if (endSlot <= startSlot) endSlot = startSlot + 1;
      // Hide covered cells (for rowspan effect)
      for (let slot = startSlot + 1; slot < endSlot; slot++) {
        const coveredCell = document.querySelector(`#calendar-table td[data-day="${dayIdx}"][data-slot="${slot}"]`);
        if (coveredCell) coveredCell.style.display = 'none';
      }
      // Place block in the flex row of the top cell
      const cell = document.querySelector(`#calendar-table td[data-day="${dayIdx}"][data-slot="${startSlot}"]`);
      if (cell) {
        cell.rowSpan = endSlot - startSlot;
        const color = colors[meeting.calendarIndex % colors.length];
        const border = borders[meeting.calendarIndex % borders.length];
        const rowDiv = cell.querySelector('.meeting-row');
        if (rowDiv) {
          rowDiv.innerHTML += `<div class="meeting" style="background:${color};border-left:4px solid ${border};height:100%;min-width:100px;flex:1 1 0;max-width:${100 / maxOverlaps[dayIdx]}%;margin:1px;box-sizing:border-box;overflow:auto;padding:2px 4px;display:flex;flex-direction:column;justify-content:center;align-items:flex-start;z-index:1;"><strong>${meeting.subject}</strong><br>${start.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })} - ${end.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</div>`;
        }
      }
    });
  });
}

function createWeekSwitcher(currentWeekStart, currentMonth, onChange) {
  // Find all week starts in the current month
  const year = currentWeekStart.getFullYear();
  const firstOfMonth = new Date(year, currentMonth, 1);
  const lastOfMonth = new Date(year, currentMonth + 1, 0);
  let weekStarts = [];
  let d = getStartOfWeek(firstOfMonth);
  // Only include weeks that have at least one day in the current month
  while (d <= lastOfMonth) {
    const weekEnd = new Date(d);
    weekEnd.setDate(d.getDate() + 6);
    if (weekEnd >= firstOfMonth && d <= lastOfMonth) {
      weekStarts.push(new Date(d));
    }
    d.setDate(d.getDate() + 7);
  }
  // Build control
  const switcher = document.createElement('div');
  switcher.id = 'week-switcher';
  switcher.style.display = 'flex';
  switcher.style.justifyContent = 'center';
  switcher.style.alignItems = 'center';
  switcher.style.gap = '8px';
  switcher.style.margin = '16px 0 8px 0';
  weekStarts.forEach((ws, idx) => {
    const btn = document.createElement('button');
    btn.textContent = `${ws.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })}`;
    btn.style.padding = '4px 10px';
    btn.style.borderRadius = '6px';
    btn.style.border = ws.getTime() === currentWeekStart.getTime() ? '2px solid #1976d2' : '1px solid #aaa';
    btn.style.background = ws.getTime() === currentWeekStart.getTime() ? '#e3f0ff' : '#fff';
    btn.style.cursor = 'pointer';
    btn.onclick = () => onChange(ws);
    switcher.appendChild(btn);
  });
  return switcher;
}

function showLoader() {
  let loader = document.getElementById('calendar-loader');
  if (!loader) {
    loader = document.createElement('div');
    loader.id = 'calendar-loader';
    loader.style.position = 'fixed';
    loader.style.top = '0';
    loader.style.left = '0';
    loader.style.width = '100vw';
    loader.style.height = '100vh';
    loader.style.background = 'rgba(255,255,255,0.7)';
    loader.style.display = 'flex';
    loader.style.justifyContent = 'center';
    loader.style.alignItems = 'center';
    loader.style.zIndex = '1000';
    loader.innerHTML = '<div style="padding:32px 48px;background:#fff;border-radius:12px;box-shadow:0 2px 16px #0002;font-size:1.3em;font-weight:bold;display:flex;align-items:center;gap:12px;"><span class="loader-spinner" style="width:24px;height:24px;border:4px solid #1976d2;border-top:4px solid #e3f0ff;border-radius:50%;display:inline-block;animation:spin 1s linear infinite;"></span>Loading…</div>';
    document.body.appendChild(loader);
    // Add spinner animation
    const style = document.createElement('style');
    style.innerHTML = '@keyframes spin{0%{transform:rotate(0deg);}100%{transform:rotate(360deg);}}';
    document.head.appendChild(style);
  } else {
    loader.style.display = 'flex';
  }
}

function hideLoader() {
  const loader = document.getElementById('calendar-loader');
  if (loader) loader.style.display = 'none';
}

window.onload = async () => {
  document.body.innerHTML = '';
  const today = new Date();
  const currentMonth = today.getMonth();
  let weekStart = getStartOfWeek(today);
  // If this week is not in the current month, move to the first week of the month
  if (weekStart.getMonth() !== currentMonth) {
    const firstOfMonth = new Date(today.getFullYear(), currentMonth, 1);
    weekStart = getStartOfWeek(firstOfMonth);
    if (weekStart.getMonth() !== currentMonth) {
      weekStart.setDate(weekStart.getDate() + 7);
    }
  }
  let selectedWeekStart = new Date(weekStart);

  showLoader();
  // Fetch meetings only once
  const meetings = await fetchMeetings();
  hideLoader();

  async function renderWeek(weekStart) {
    document.body.innerHTML = '';
    // Add week switcher
    const switcher = createWeekSwitcher(weekStart, currentMonth, async (ws) => {
      selectedWeekStart = new Date(ws);
      await renderWeek(selectedWeekStart);
    });
    document.body.appendChild(switcher);
    const { weekDays } = createCalendarGrid(weekStart);
    // Only show meetings in this week
    const filteredMeetings = meetings.filter(m => {
      const startStr = m.startWithTimeZone || m.start?.dateTime || m.start;
      return isDateInWeek(startStr, weekStart);
    });
    renderMeetings(filteredMeetings, weekStart);
  }

  await renderWeek(selectedWeekStart);
};
