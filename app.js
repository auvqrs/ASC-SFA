// Enhanced per-year timetable app (stability fixes & scheduler timeout to prevent crashes)
// - Adds safety checks to auto-scheduler to avoid long-running recursion/timeouts.
// - Early capacity check to detect impossible placement counts.
// - Time-bounded MRV/backtracking attempt, falling back to greedy placement if timeout.
// - Better error handling to ensure UI does not crash.

(function(){
  // State
  const state = {
    days: ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'],
    periods: 5, // P1..P5 (slot 0 is Form)
    years: ['Y7','Y8','Y9','Y10','Y11','LSU','SF'],
    subjects: [], // {id,name,color,defaultCount, mergeYears: []}
    teachers: [], // {id, name, code, workingAvailability: { Mon: [true,...slots], ... }}
    rooms: [], // {id, name}
    lessons: [], // {id, years:[], subjectId, teacherId, roomId, count}
    // grid: for each year -> days -> slots (slots = 1 + periods, slot0 = Form)
    grid: [], // built by createEmptyGrid()
    pickedLessonId: null,
    filter: { type: 'all', value: '' }
  };

  // Utils
  const $ = (sel, root=document) => root.querySelector(sel);
  const $$ = (sel, root=document) => Array.from(root.querySelectorAll(sel));
  const uid = (p='id') => p + '_' + Math.random().toString(36).slice(2,9);

  // Elements
  const el = {
    selectDays: $('#selectDays'),
    inputPeriods: $('#inputPeriods'),
    btnApplySettings: $('#btnApplySettings'),
    yearsList: $('#yearsList'),
    subjectsList: $('#subjectsList'),
    roomsList: $('#roomsList'),
    teachersList: $('#teachersList'),
    teacherFullName: $('#teacherFullName'),
    btnAddTeacher: $('#btnAddTeacher'),
    lessonYear: $('#lessonYear'),
    lessonSubject: $('#lessonSubject'),
    lessonTeacher: $('#lessonTeacher'),
    lessonRoom: $('#lessonRoom'),
    lessonCount: $('#lessonCount'),
    btnAddLesson: $('#btnAddLesson'),
    btnGenerateLessonsForYears: $('#btnGenerateLessonsForYears'),
    tokens: $('#tokens'),
    timetable: $('#timetable'),
    btnAutoSchedule: $('#btnAutoSchedule'),
    btnClear: $('#btnClear'),
    btnResetData: $('#btnResetData'),
    btnExportJSON: $('#btnExportJSON'),
    btnImportJSON: $('#btnImportJSON'),
    fileImport: $('#fileImport'),
    btnExportCSV: $('#btnExportCSV'),
    btnExportExcel: $('#btnExportExcel'),
    btnPrint: $('#btnPrint'),
    filterType: $('#filterType'),
    filterValue: $('#filterValue'),
    btnApplyFilter: $('#btnApplyFilter'),
    btnClearFilter: $('#btnClearFilter'),
    btnAddSubject: $('#btnAddSubject'),
    btnEditSubject: $('#btnEditSubject')
  };

  // Default subject list (includes Form Time)
  const defaultSubjects = [
    ['Form Time',0],
    ['English Language',3],
    ['Mathematics',3],
    ['Science',3],
    ['History',2],
    ['Geography',2],
    ['Religious Education',1],
    ['Business Studies',1],
    ['German',2],
    ['French',2],
    ['Spanish',1],
    ['Art & Design',1],
    ['Drama',1],
    ['Music',1],
    ['Fashion Design',1],
    ['Food Technology',2],
    ['Design Technology',1],
    ['Computer Science',2],
    ['Sociology',1],
    ['Child Development',1],
    ['Psychology',1],
    ['Economics',1],
    ['Physical Education',1],
    ['Sports Science',1]
  ];

  // Simple DOM modal helpers (non-blocking). They return Promises so callers can await them.
  function alertDialog(message, title = 'Notice') {
    return new Promise(resolve => {
      const modal = createModalContainer();
      const box = createModalBox(title);

      const msg = document.createElement('div');
      msg.style.margin = '8px 0 12px';
      msg.textContent = message;
      box.appendChild(msg);

      const btn = document.createElement('button');
      btn.textContent = 'OK';
      stylePrimaryButton(btn);
      btn.addEventListener('click', () => { document.body.removeChild(modal); resolve(); });
      box.appendChild(btn);

      modal.appendChild(box);
      document.body.appendChild(modal);
      btn.focus();
    });
  }

  function confirmDialog(message, title = 'Confirm') {
    return new Promise(resolve => {
      const modal = createModalContainer();
      const box = createModalBox(title);

      const msg = document.createElement('div');
      msg.style.margin = '8px 0 12px';
      msg.textContent = message;
      box.appendChild(msg);

      const row = document.createElement('div');
      row.style.display = 'flex';
      row.style.justifyContent = 'flex-end';
      row.style.gap = '8px';

      const cancel = document.createElement('button');
      cancel.textContent = 'Cancel';
      cancel.style.padding = '8px 12px';
      cancel.style.borderRadius = '6px';
      cancel.style.border = '1px solid #e0e6ef';
      cancel.addEventListener('click', () => { document.body.removeChild(modal); resolve(false); });

      const ok = document.createElement('button');
      ok.textContent = 'OK';
      stylePrimaryButton(ok);
      ok.addEventListener('click', () => { document.body.removeChild(modal); resolve(true); });

      row.appendChild(cancel);
      row.appendChild(ok);
      box.appendChild(row);

      modal.appendChild(box);
      document.body.appendChild(modal);
      ok.focus();
    });
  }

  function createModalContainer(){
    const modal = document.createElement('div');
    modal.style.position = 'fixed';
    modal.style.left = '0';
    modal.style.top = '0';
    modal.style.right = '0';
    modal.style.bottom = '0';
    modal.style.background = 'rgba(0,0,0,0.25)';
    modal.style.zIndex = 9999;
    modal.addEventListener('click', (e)=>{ if(e.target === modal) { /*click outside does nothing*/ } });
    return modal;
  }
  function createModalBox(title){
    const box = document.createElement('div');
    box.style.width = '520px';
    box.style.maxHeight = '80vh';
    box.style.overflow = 'auto';
    box.style.margin = '40px auto';
    box.style.background = '#fff';
    box.style.padding = '14px';
    box.style.borderRadius = '10px';
    box.style.boxShadow = '0 10px 40px rgba(10,20,40,0.2)';
    const h = document.createElement('h3');
    h.textContent = title;
    h.style.margin = '0 0 8px';
    box.appendChild(h);
    return box;
  }
  function stylePrimaryButton(btn){
    btn.style.padding = '8px 12px';
    btn.style.borderRadius = '6px';
    btn.style.border = 'none';
    btn.style.background = '#1976d2';
    btn.style.color = '#fff';
    btn.style.cursor = 'pointer';
  }

  // Storage
  function save() {
    const data = {
      days: state.days, periods: state.periods, years: state.years,
      subjects: state.subjects, teachers: state.teachers, rooms: state.rooms,
      lessons: state.lessons, grid: state.grid
    };
    try { localStorage.setItem('school_timetable_per_year_v2', JSON.stringify(data)); } catch(e) { console.warn('Save failed', e); }
  }
  function load(){
    try {
      const raw = localStorage.getItem('school_timetable_per_year_v2');
      if(raw){
        const d = JSON.parse(raw);
        state.days = d.days || state.days;
        state.periods = d.periods || state.periods;
        state.years = d.years || state.years;
        state.subjects = d.subjects || [];
        state.teachers = d.teachers || [];
        state.rooms = d.rooms || [];
        state.lessons = d.lessons || [];
        state.grid = d.grid || createEmptyGrid();
        normalizeGrid();
        normalizeTeachersAvailability();
        return;
      }
    } catch(e){ console.warn('Load failed', e); }
    // If no saved data, initialize defaults
    populateDefaults();
    state.grid = createEmptyGrid();
    normalizeGrid();
    save();
  }

  // Defaults
  function populateDefaults(){
    state.subjects = defaultSubjects.map(s=>({ id: uid('sub'), name: s[0], color: randomColor(), defaultCount: s[1], mergeYears: [] }));
    // rooms
    const rooms = [];
    for(let i=101;i<=116;i++) rooms.push('G-'+i);
    for(let i=117;i<=126;i++) rooms.push('F-'+i);
    for(let i=127;i<=138;i++) rooms.push('S-'+i);
    state.rooms = rooms.map(r=>({id:uid('room'), name:r}));
    state.teachers = [];
    state.lessons = [];
  }

  function randomColor(){
    const hue = Math.floor(Math.random()*360);
    return `hsl(${hue} 70% 60%)`;
  }

  // Grid helpers
  function slotsCount(){ return 1 + state.periods; } // slot0 = Form
  function createEmptyGrid(){
    const slots = slotsCount();
    const grid = [];
    for(let y=0; y<state.years.length; y++){
      const yearRow = [];
      for(let d=0; d<state.days.length; d++){
        const daySlots = new Array(slots).fill(null);
        yearRow.push(daySlots);
      }
      grid.push(yearRow);
    }
    return grid;
  }

  function normalizeGrid(){
    if(!Array.isArray(state.grid)) state.grid = createEmptyGrid();
    const slots = slotsCount();
    while(state.grid.length < state.years.length) state.grid.push([]);
    if(state.grid.length > state.years.length) state.grid.splice(state.years.length);
    for(let y=0; y<state.years.length; y++){
      if(!Array.isArray(state.grid[y])) state.grid[y] = [];
      while(state.grid[y].length < state.days.length) state.grid[y].push(new Array(slots).fill(null));
      if(state.grid[y].length > state.days.length) state.grid[y].splice(state.days.length);
      for(let d=0; d<state.days.length; d++){
        if(!Array.isArray(state.grid[y][d])) state.grid[y][d] = new Array(slots).fill(null);
        if(state.grid[y][d].length < slots){
          while(state.grid[y][d].length < slots) state.grid[y][d].push(null);
        } else if(state.grid[y][d].length > slots){
          state.grid[y][d].splice(slots);
        }
      }
    }
  }

  // Ensure teachers' availability structures match current slots & days
  function normalizeTeachersAvailability(){
    const slots = slotsCount();
    state.teachers.forEach(t=>{
      if(!t.workingAvailability) t.workingAvailability = {};
      state.days.forEach(d=>{
        if(!Array.isArray(t.workingAvailability[d])) t.workingAvailability[d] = new Array(slots).fill(true);
        else {
          // normalize length
          if(t.workingAvailability[d].length < slots){
            while(t.workingAvailability[d].length < slots) t.workingAvailability[d].push(true);
          } else if(t.workingAvailability[d].length > slots){
            t.workingAvailability[d].splice(slots);
          }
        }
      });
    });
  }

  // Rendering (try/catch keeps UI alive)
  function renderAll(){
    try {
      normalizeGrid();
      normalizeTeachersAvailability();
      renderSettings();
      renderYears();
      renderSubjects();
      renderRooms();
      renderTeachers();
      renderLessonSelectors();
      renderTokens();
      renderGrid();
      renderFilterOptions();
      save();
    } catch(err){
      console.error('renderAll error', err);
      el.timetable.innerHTML = '<div style="color:#b00;padding:12px">An error occurred — check the console for details.</div>';
    }
  }

  function renderSettings(){
    for(const opt of el.selectDays.options) opt.selected = state.days.includes(opt.value);
    el.inputPeriods.value = state.periods;
  }

  function renderYears(){
    const container = el.yearsList; container.innerHTML = '';
    state.years.forEach(y=>{
      const d = document.createElement('div'); d.className='list-item';
      const left = document.createElement('div'); left.textContent = y;
      const right = document.createElement('div'); right.innerHTML = `<span class="badge">click to toggle</span>`;
      d.appendChild(left); d.appendChild(right);
      d.addEventListener('click', ()=> {
        d.classList.toggle('active');
        d.style.background = d.classList.contains('active') ? '#eef8ff' : '';
      });
      container.appendChild(d);
    });
  }

  function renderSubjects(){
    const container = el.subjectsList; container.innerHTML = '';
    state.subjects.forEach(s=>{
      const div = document.createElement('div'); div.className='subject-row';
      const left = document.createElement('div'); left.className='sub-left';
      const color = document.createElement('span'); color.className='sub-color'; color.style.background=s.color;
      const name = document.createElement('div');
      name.innerHTML = `<strong>${s.name}</strong><div class="small">per week: ${s.defaultCount} • merges: ${s.mergeYears && s.mergeYears.length ? s.mergeYears.join(', ') : 'none'}</div>`;
      left.appendChild(color); left.appendChild(name);
      const right = document.createElement('div');
      right.innerHTML = `<button data-id="${s.id}" class="edit-subject">Edit</button> <button data-id="${s.id}" class="remove-subject">Remove</button>`;
      div.appendChild(left); div.appendChild(right);
      container.appendChild(div);
    });
    $$('.edit-subject', container).forEach(b=> b.addEventListener('click', ()=> editSubjectModal(b.dataset.id)));
    $$('.remove-subject', container).forEach(b=> b.addEventListener('click', async ()=> {
      const id = b.dataset.id;
      const ok = await confirmDialog('Remove subject and its lessons?');
      if(!ok) return;
      const removedLessonIds = state.lessons.filter(l => l.subjectId === id).map(l => l.id);
      state.lessons = state.lessons.filter(l=>l.subjectId !== id);
      for(let y=0;y<state.grid.length;y++){
        for(let d=0;d<state.grid[y].length;d++){
          for(let sidx=0;sidx<state.grid[y][d].length;sidx++){
            const cell = state.grid[y][d][sidx];
            if(cell && (removedLessonIds.includes(cell.lessonId) || cell.subjectId === id)){
              state.grid[y][d][sidx] = null;
            }
          }
        }
      }
      state.subjects = state.subjects.filter(s=>s.id !== id);
      renderAll();
    }));
  }

  function renderRooms(){
    const c = el.roomsList; c.innerHTML='';
    state.rooms.forEach(r=>{
      const d = document.createElement('div'); d.className='list-item'; d.textContent = r.name; c.appendChild(d);
    });
  }

  function renderTeachers(){
    const c = el.teachersList; c.innerHTML='';
    state.teachers.forEach(t=>{
      const d = document.createElement('div'); d.className='list-item';
      const left = document.createElement('div');
      left.innerHTML = `<strong>${t.name}</strong> <span class="teacher-badge">${t.code}</span><div class="small">Availability: ${formatTeacherSummary(t)}</div>`;
      const right = document.createElement('div');
      right.innerHTML = `<button data-id="${t.id}" class="edit-teacher">Edit</button> <button data-id="${t.id}" class="remove-teacher">Remove</button>`;
      d.appendChild(left); d.appendChild(right);
      c.appendChild(d);
    });
    $$('.edit-teacher', c).forEach(btn=> btn.addEventListener('click', ()=> editTeacherModal(btn.dataset.id)));
    $$('.remove-teacher', c).forEach(btn=> btn.addEventListener('click', async ()=>{
      const id = btn.dataset.id;
      const ok = await confirmDialog('Remove teacher and their lessons?');
      if(!ok) return;
      state.lessons = state.lessons.filter(l=>l.teacherId !== id);
      for(let y=0;y<state.grid.length;y++){
        for(let d=0;d<state.grid[y].length;d++){
          for(let s=0;s<state.grid[y][d].length;s++){
            if(state.grid[y][d][s] && state.grid[y][d][s].teacherId === id) state.grid[y][d][s] = null;
          }
        }
      }
      state.teachers = state.teachers.filter(t=>t.id!==id);
      renderAll();
    }));
  }

  function formatTeacherSummary(t){
    // Give short summary like Mon:P1-5, Tue:P1-3...
    const parts = [];
    const slots = slotsCount();
    state.days.forEach(d=>{
      if(!t.workingAvailability || !t.workingAvailability[d]) return;
      const arr = t.workingAvailability[d];
      if(arr.every(v=>v)) parts.push(d);
      else if(arr.every(v=>!v)) {}
      else {
        const on = [];
        for(let i=0;i<arr.length;i++) if(arr[i]) on.push(i===0?'Form':('P'+i));
        parts.push(`${d}(${on.join(',')})`);
      }
    });
    return parts.slice(0,4).join(', ') + (parts.length>4 ? '...' : '');
  }

  function renderLessonSelectors(){
    if(el.lessonYear) { el.lessonYear.innerHTML=''; state.years.forEach(y=>{ const o=document.createElement('option'); o.value=y; o.textContent=y; el.lessonYear.appendChild(o); }); }
    if(el.lessonSubject) { el.lessonSubject.innerHTML=''; state.subjects.forEach(s=>{ const o=document.createElement('option'); o.value=s.id; o.textContent=`${s.name} (${s.defaultCount}/wk)`; el.lessonSubject.appendChild(o); }); }
    if(el.lessonTeacher) { el.lessonTeacher.innerHTML=''; const blank=document.createElement('option'); blank.value=''; blank.textContent='(none)'; el.lessonTeacher.appendChild(blank);
      state.teachers.forEach(t=>{ const o=document.createElement('option'); o.value=t.id; o.textContent=`${t.name} • ${t.code}`; el.lessonTeacher.appendChild(o); }); }
    if(el.lessonRoom) { el.lessonRoom.innerHTML=''; const none=document.createElement('option'); none.value=''; none.textContent='(none)'; el.lessonRoom.appendChild(none);
      state.rooms.forEach(r=>{ const o=document.createElement('option'); o.value=r.id; o.textContent=r.name; el.lessonRoom.appendChild(o); }); }
  }

  // Tokens - unassigned lesson instances (respect years)
  function renderTokens(){
    const c = el.tokens; c.innerHTML='';
    const counts = {};
    state.lessons.forEach(l => counts[l.id] = (counts[l.id]||0) + l.count );
    // subtract placed instances (count each placedId once for merged placements)
    const placedTracker = new Set();
    for(let y=0;y<state.grid.length;y++){
      for(let d=0; d<state.grid[y].length; d++){
        for(let s=0;s<state.grid[y][d].length;s++){
          const cell = state.grid[y][d][s];
          if(cell && cell.lessonId){
            if(cell.placedId){
              if(!placedTracker.has(cell.placedId)){
                placedTracker.add(cell.placedId);
                counts[cell.lessonId] = Math.max(0,(counts[cell.lessonId]||0)-1);
              }
            } else {
              // fallback for old cells without placedId: decrement per cell but guard against negative
              counts[cell.lessonId] = Math.max(0,(counts[cell.lessonId]||0)-1);
            }
          }
        }
      }
    }
    state.lessons.forEach(l=>{
      const remaining = counts[l.id] || 0;
      if(remaining <= 0) return;
      const sub = state.subjects.find(s=>s.id===l.subjectId);
      const teacher = state.teachers.find(t=>t.id===l.teacherId);
      const room = state.rooms.find(r=>r.id===l.roomId);
      for(let i=0;i<remaining;i++){
        const div = document.createElement('div'); div.className='token'; div.draggable=true;
        div.dataset.lessonId = l.id;
        div.style.background = sub ? sub.color : '#777';
        const yearsLabel = (l.years && l.years.length>1) ? l.years.join('+') : (l.years && l.years[0]) || '';
        div.innerHTML = `<div style="display:flex;flex-direction:column"><div style="font-weight:700">${sub?sub.name:'?'}</div><div style="font-size:12px">${yearsLabel} ${teacher?('• '+teacher.code):''}${room?(' • '+room.name):''}</div></div>`;
        div.addEventListener('dragstart', tokenDragStart);
        div.addEventListener('click', ()=> { state.pickedLessonId = l.id; highlightPickedToken(); });
        c.appendChild(div);
      }
    });
    highlightPickedToken();
  }

  function highlightPickedToken(){
    $$('.token').forEach(t=> t.style.outline = (t.dataset.lessonId === state.pickedLessonId) ? '3px solid rgba(0,0,0,0.12)' : 'none');
  }

  // Render grid table: rows = years, columns = days × slots
  function renderGrid(){
    el.timetable.innerHTML = '';
    const slots = slotsCount();
    const days = state.days;
    const years = state.years;

    const table = document.createElement('table'); table.className = 'timetable-table';

    // THEAD top: day headers (colspan = slots)
    const thead = document.createElement('thead');
    const trTop = document.createElement('tr');
    const thCorner = document.createElement('th'); thCorner.textContent = 'Year / Day'; trTop.appendChild(thCorner);
    days.forEach(day=>{
      const th = document.createElement('th'); th.colSpan = slots; th.textContent = day; trTop.appendChild(th);
    });
    thead.appendChild(trTop);

    // THEAD slot labels
    const trSlots = document.createElement('tr');
    const thEmpty = document.createElement('th'); thEmpty.textContent = ''; trSlots.appendChild(thEmpty);
    days.forEach(()=> {
      const thForm = document.createElement('th'); thForm.className='slot-header'; thForm.textContent = 'Form'; trSlots.appendChild(thForm);
      for(let p=1;p<slots;p++){
        const thP = document.createElement('th'); thP.className='slot-header'; thP.textContent = 'P'+p; trSlots.appendChild(thP);
      }
    });
    thead.appendChild(trSlots);
    table.appendChild(thead);

    // TBODY: each year row
    const tbody = document.createElement('tbody');
    for(let y=0;y<years.length;y++){
      const tr = document.createElement('tr');
      const thYear = document.createElement('th'); thYear.className='year-label'; thYear.textContent = years[y]; tr.appendChild(thYear);

      for(let d=0; d<days.length; d++){
        for(let s=0; s<slots; s++){
          const td = document.createElement('td'); td.className='cell';
          td.dataset.yearIndex = y; td.dataset.dayIndex = d; td.dataset.slotIndex = s;
          td.addEventListener('dragover', cellDragOver);
          td.addEventListener('dragleave', cellDragLeave);
          td.addEventListener('drop', cellDrop);
          td.addEventListener('click', cellClick);

          const assigned = state.grid[y] && state.grid[y][d] && state.grid[y][d][s];
          if(assigned){
            const sub = state.subjects.find(x=>x.id===assigned.subjectId);
            const t = state.teachers.find(x=>x.id===assigned.teacherId);
            const r = state.rooms.find(x=>x.id===assigned.roomId);
            const wrapper = document.createElement('div');
            wrapper.innerHTML = `<div class="assign" style="background:${sub?sub.color:'#999'};color:white;padding:4px;border-radius:4px">${sub?sub.name:'(Form)'} ${assigned.mergedYears?('• '+assigned.mergedYears.join('+')):''}</div>
              <div class="meta">${t?('<span class="teacher-badge">'+t.code+'</span>') : ''} ${r?(' • '+r.name):''}</div>`;
            td.appendChild(wrapper);

            const btn = document.createElement('button');
            btn.textContent='✕'; btn.title='Remove assignment';
            btn.style.position='absolute'; btn.style.right='6px'; btn.style.top='6px';
            btn.style.border='none'; btn.style.background='rgba(0,0,0,0.06)'; btn.style.borderRadius='4px'; btn.style.cursor='pointer';
            btn.addEventListener('click',(ev)=>{ ev.stopPropagation(); removePlacedById(assigned.placedId); renderAll(); });
            td.appendChild(btn);
          } else {
            td.innerHTML = '<div class="small">empty</div>';
          }

          // filter dimming
          if(!cellMatchesFilter(assigned, y)){
            td.classList.add('dimmed');
          } else td.classList.remove('dimmed');

          // teacher working day/slot conflict highlight
          if(assigned && assigned.teacherId){
            const teacher = state.teachers.find(tt=>tt.id===assigned.teacherId);
            const dayName = state.days[d];
            if(teacher && (!teacher.workingAvailability || !teacher.workingAvailability[dayName] || !teacher.workingAvailability[dayName][s])){
              td.style.boxShadow = 'inset 0 0 0 1000px rgba(255,80,80,0.06)';
            } else td.style.boxShadow = '';
          }

          tr.appendChild(td);
        }
      }

      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    el.timetable.appendChild(table);
  }

  // Remove placed object by its placedId across all years
  function removePlacedById(placedId){
    for(let y=0;y<state.grid.length;y++){
      for(let d=0; d<state.grid[y].length; d++){
        for(let s=0; s<state.grid[y][d].length; s++){
          if(state.grid[y][d][s] && state.grid[y][d][s].placedId === placedId) state.grid[y][d][s] = null;
        }
      }
    }
  }

  // Filter helpers
  function renderFilterOptions(){
    el.filterType.value = state.filter.type;
    el.filterValue.innerHTML = '';
    if(state.filter.type === 'all' || state.filter.type === 'staff'){ el.filterValue.style.display='none'; return; }
    el.filterValue.style.display='inline-block';
    if(state.filter.type === 'year'){
      state.years.forEach(y=>{ const o=document.createElement('option'); o.value=y; o.textContent=y; el.filterValue.appendChild(o); });
    } else if(state.filter.type === 'teacher'){
      const blank=document.createElement('option'); blank.value=''; blank.textContent='(choose)'; el.filterValue.appendChild(blank);
      state.teachers.forEach(t=>{ const o=document.createElement('option'); o.value=t.id; o.textContent=`${t.name} (${t.code})`; el.filterValue.appendChild(o); });
    } else if(state.filter.type === 'room'){
      const blank=document.createElement('option'); blank.value=''; blank.textContent='(choose)'; el.filterValue.appendChild(blank);
      state.rooms.forEach(r=>{ const o=document.createElement('option'); o.value=r.id; o.textContent=r.name; el.filterValue.appendChild(o); });
    } else if(state.filter.type === 'subject'){
      const blank=document.createElement('option'); blank.value=''; blank.textContent='(choose)'; el.filterValue.appendChild(blank);
      state.subjects.forEach(s=>{ const o=document.createElement('option'); o.value=s.id; o.textContent=s.name; el.filterValue.appendChild(o); });
    }
    if(state.filter.value) el.filterValue.value = state.filter.value;
  }

  function cellMatchesFilter(assigned, yearIndex){
    const f = state.filter;
    if(!f || f.type === 'all') return true;
    if(f.type === 'staff') return assigned && !!assigned.teacherId;
    if(f.type === 'year'){
      const val = f.value;
      if(!val) return true;
      return state.years[yearIndex] === val;
    }
    if(!assigned) return false;
    if(f.type === 'teacher') return assigned.teacherId === f.value;
    if(f.type === 'room') return assigned.roomId === f.value;
    if(f.type === 'subject') return assigned.subjectId === f.value;
    return true;
  }

  // Drag & drop
  let draggingLessonId = null;
  function tokenDragStart(e){
    draggingLessonId = e.currentTarget.dataset.lessonId;
    try { e.dataTransfer.setData('text/plain', draggingLessonId); } catch(err) { /* some browsers restrict setData on certain events */ }
  }
  function cellDragOver(e){ e.preventDefault(); e.currentTarget.classList.add('drag-over'); }
  function cellDragLeave(e){ e.currentTarget.classList.remove('drag-over'); }

  async function cellDrop(e){
    e.preventDefault();
    e.currentTarget.classList.remove('drag-over');
    const lessonId = (e.dataTransfer && e.dataTransfer.getData && e.dataTransfer.getData('text/plain')) || draggingLessonId || state.pickedLessonId;
    if(!lessonId) return;
    const y = +e.currentTarget.dataset.yearIndex;
    const d = +e.currentTarget.dataset.dayIndex;
    const s = +e.currentTarget.dataset.slotIndex;
    const lesson = state.lessons.find(l=>l.id===lessonId);
    if(!lesson) return;
    // check if single-year placement matches row if not merged
    if(!(lesson.years && lesson.years.includes(state.years[y]))){
      await alertDialog(`This lesson targets ${lesson.years.join(', ')}. You can only place it on one of those year rows (or pick a slot while the token is selected).`);
      return;
    }
    // Prevent placing non-Form subjects into slot 0
    const subject = state.subjects.find(su => su.id === lesson.subjectId);
    const isForm = subject && typeof subject.name === 'string' && subject.name.toLowerCase().includes('form');
    if(s === 0 && !isForm){
      await alertDialog('Only Form lessons may be placed in the Form slot. Pick a period (P1..) for this lesson.');
      return;
    }
    await assignLessonToCell(lessonId, y, d, s);
    draggingLessonId = null;
    state.pickedLessonId = null;
    renderAll();
  }

  async function assignLessonToCell(lessonId, yearIndex, dayIndex, slotIndex){
    const lesson = state.lessons.find(l=>l.id===lessonId);
    if(!lesson) return;
    const subject = state.subjects.find(su => su.id === lesson.subjectId);
    const isForm = subject && typeof subject.name === 'string' && subject.name.toLowerCase().includes('form');
    if(slotIndex === 0 && !isForm){
      await alertDialog('Only Form lessons can be placed in the Form slot.');
      return;
    }

    // build target years array
    const targetYears = lesson.years && lesson.years.length ? lesson.years : [state.years[yearIndex]];
    // map to indices
    const yearIndices = targetYears.map(y => state.years.indexOf(y)).filter(i=>i>=0);
    if(yearIndices.length !== targetYears.length){
      await alertDialog('One or more target years missing from the configuration. Update years list or lesson definition.');
      return;
    }

    // check teacher availability and room conflicts (hard constraints)
    if(lesson.teacherId){
      const teacher = state.teachers.find(t=>t.id===lesson.teacherId);
      const dayName = state.days[dayIndex];
      let unavailableForSome = false;
      if(teacher){
        for(const yi of yearIndices){
          if(!(teacher.workingAvailability && teacher.workingAvailability[dayName] && teacher.workingAvailability[dayName][slotIndex])){
            unavailableForSome = true; break;
          }
        }
      }
      if(unavailableForSome){
        const ok = await confirmDialog(`Teacher ${teacher.name} does not work that slot for one or more selected year(s). Place anyway?`);
        if(!ok) return;
      }
    }

    // check rooms overlapping and slot occupancy across target years (hard)
    if(lesson.roomId){
      for(const yi of yearIndices){
        for(let y=0;y<state.grid.length;y++){
          const cell = state.grid[y][dayIndex] && state.grid[y][dayIndex][slotIndex];
          if(cell && cell.roomId === lesson.roomId){
            if(!(cell.lessonId === lesson.id && yearIndices.includes(y))) {
              await alertDialog(`Room ${state.rooms.find(r=>r.id===lesson.roomId).name} is already used in ${state.years[y]} on ${state.days[dayIndex]} ${slotLabel(slotIndex)}. Choose a different slot/room.`);
              return;
            }
          }
        }
      }
    }

    // assign: create placed object and set into grid in all yearIndices for that day/slot
    const placed = {
      lessonId: lesson.id,
      subjectId: lesson.subjectId,
      teacherId: lesson.teacherId,
      roomId: lesson.roomId,
      mergedYears: targetYears.length>1 ? targetYears.slice() : null,
      placedId: uid('placed')
    };

    // ensure target slots are empty (unless the slot already has same placedId)
    for(const yi of yearIndices){
      const existing = state.grid[yi][dayIndex][slotIndex];
      if(existing && existing.placedId !== placed.placedId){
        const ok = await confirmDialog(`Slot for ${state.years[yi]} ${state.days[dayIndex]} ${slotLabel(slotIndex)} is already occupied. Overwrite?`);
        if(!ok) return;
      }
    }

    for(const yi of yearIndices){
      state.grid[yi][dayIndex][slotIndex] = Object.assign({}, placed);
    }
  }

  function slotLabel(s){
    return s===0 ? 'Form' : 'P'+s;
  }

  // Cell click quick assign
  async function cellClick(e){
    const y = +e.currentTarget.dataset.yearIndex;
    const d = +e.currentTarget.dataset.dayIndex;
    const s = +e.currentTarget.dataset.slotIndex;
    if(state.pickedLessonId){
      const lesson = state.lessons.find(l=>l.id===state.pickedLessonId);
      if(lesson && !(lesson.years && lesson.years.includes(state.years[y]))){
        await alertDialog(`Picked lesson targets ${lesson.years.join(', ')}. Click a matching year row or create the lesson for this year.`);
        return;
      }
      // Prevent placing non-Form into Form slot
      const subject = state.subjects.find(su => su.id === lesson.subjectId);
      const isForm = subject && subject.name && subject.name.toLowerCase().includes('form');
      if(s === 0 && !isForm){
        await alertDialog('Only Form lessons can be placed in the Form slot.');
        return;
      }
      await assignLessonToCell(state.pickedLessonId, y, d, s);
      state.pickedLessonId = null; renderAll(); return;
    }
    quickAssignPopup(e.currentTarget, y, d, s);
  }

  function quickAssignPopup(cellEl, yearIndex, dayIndex, slotIndex){
    // compute remaining counts considering merged lessons count consumed once per placement
    const remainingCounts = {};
    state.lessons.forEach(l => remainingCounts[l.id] = (remainingCounts[l.id]||0) + l.count );
    const placedTracker = new Set();
    for(let y=0;y<state.grid.length;y++){
      for(let d=0; d<state.grid[y].length; d++){
        for(let si=0; si<state.grid[y][d].length; si++){
          const cell = state.grid[y][d][si];
          if(cell && cell.lessonId && !placedTracker.has(cell.placedId)){
            remainingCounts[cell.lessonId] = Math.max(0,(remainingCounts[cell.lessonId]||0)-1);
            placedTracker.add(cell.placedId);
          }
        }
      }
    }

    const available = state.lessons.filter(l => l.years && l.years.includes(state.years[yearIndex]) && (remainingCounts[l.id]||0) > 0);

    const menu = document.createElement('div');
    menu.style.position='absolute'; menu.style.left='10px'; menu.style.top='10px';
    menu.style.zIndex=1000; menu.style.background='white'; menu.style.border='1px solid #e2e8f0';
    menu.style.borderRadius='6px'; menu.style.padding='8px'; menu.style.boxShadow='0 8px 24px rgba(0,0,0,0.08)';
    const title = document.createElement('div'); title.textContent=`Assign for ${state.years[yearIndex]} (${state.days[dayIndex]} ${slotLabel(slotIndex)})`; title.style.fontWeight='700'; title.style.marginBottom='6px'; menu.appendChild(title);

    if(available.length === 0){
      const none = document.createElement('div'); none.textContent = 'No assignable lessons remaining for this year.'; menu.appendChild(none);
    } else {
      available.slice(0,300).forEach(l=>{
        const s = state.subjects.find(x=>x.id===l.subjectId);
        const t = state.teachers.find(x=>x.id===l.teacherId);
        const yearsStr = l.years.join('+');
        const row = document.createElement('div'); row.style.marginBottom='6px';
        const btn = document.createElement('button'); btn.textContent = `${yearsStr} • ${s? s.name : '?'} ${t?(' • '+t.code):''}`; btn.style.width='100%';
        btn.addEventListener('click', async ()=>{ 
          // enforce Form-only on Form slot
          const isForm = s && s.name && s.name.toLowerCase().includes('form');
          if(slotIndex === 0 && !isForm){
            await alertDialog('Only Form lessons may be placed in the Form slot.');
            return;
          }
          await assignLessonToCell(l.id, yearIndex, dayIndex, slotIndex); document.body.removeChild(menu); renderAll(); 
        });
        row.appendChild(btn); menu.appendChild(row);
      });
    }
    const cancel = document.createElement('button'); cancel.textContent='Close'; cancel.style.marginTop='6px'; cancel.addEventListener('click', ()=>document.body.removeChild(menu)); menu.appendChild(cancel);
    document.body.appendChild(menu);
    const rect = cellEl.getBoundingClientRect();
    menu.style.left = Math.min(window.innerWidth-520, rect.left + 12) + 'px';
    menu.style.top = Math.min(window.innerHeight-220, rect.top + 12) + 'px';
  }

  // ---------- Improved Auto-scheduler with crash prevention ----------
  // Helpers for scheduler constraints and scoring

  function isFormSubjectById(subjectId){
    const subj = state.subjects.find(s => s.id === subjectId);
    return subj && typeof subj.name === 'string' && subj.name.toLowerCase().includes('form');
  }

  function teacherTeachesYearOnDay(grid, teacherId, yearIndex, dayIndex){
    let cnt = 0;
    const slots = slotsCount();
    for(let s=0; s<slots; s++){
      const cell = grid[yearIndex] && grid[yearIndex][dayIndex] && grid[yearIndex][dayIndex][s];
      if(cell && cell.teacherId === teacherId) cnt++;
    }
    return cnt;
  }

  // compute options with penalties (hard constraints enforced)
  function getOptionsWithPenalties(lesson, grid, opts = { allowSubjectSameDay:false, allowTeacherMultiplePerDay:false }){
    const slots = slotsCount();
    const options = [];
    const lessonYears = lesson.years && lesson.years.length ? lesson.years.slice() : [];
    if(lessonYears.length === 0) return options;
    const yearIndices = lessonYears.map(y => state.years.indexOf(y)).filter(i=>i>=0);
    if(yearIndices.length === 0) return options;

    const allowForm = isFormSubjectById(lesson.subjectId);

    for(let d=0; d<state.days.length; d++){
      for(let s=0; s<slots; s++){
        if(!allowForm && s === 0) continue; // hard rule

        let hardFail = false;
        // slots for all target years must be empty
        for(const yi of yearIndices){
          if(grid[yi][d][s]) { hardFail = true; break; }
        }
        if(hardFail) continue;

        // teacher availability and teacher-same-slot conflict are hard
        if(lesson.teacherId){
          const teacher = state.teachers.find(t=>t.id===lesson.teacherId);
          if(!teacher){ continue; }
          const dayName = state.days[d];
          for(const yi of yearIndices){
            if(!(teacher.workingAvailability && teacher.workingAvailability[dayName] && teacher.workingAvailability[dayName][s])){
              hardFail = true; break;
            }
          }
          if(hardFail) continue;
          // teacher should not be teaching same slot anywhere (cross-year)
          for(let y=0;y<grid.length;y++){
            const cell = grid[y][d][s];
            if(cell && cell.teacherId === lesson.teacherId){ hardFail = true; break; }
          }
          if(hardFail) continue;
        }

        // room conflict across all years - hard
        if(lesson.roomId){
          for(let y=0;y<grid.length;y++){
            const cell = grid[y][d][s];
            if(cell && cell.roomId === lesson.roomId){ hardFail = true; break; }
          }
          if(hardFail) continue;
        }

        // compute penalties (soft)
        let penalty = 0;
        // avoid same subject twice same day in any target year
        let subjectSameDay = false;
        for(const yi of yearIndices){
          for(let si=0; si<slots; si++){
            const cell = grid[yi][d][si];
            if(cell && cell.subjectId === lesson.subjectId){ subjectSameDay = true; break; }
          }
          if(subjectSameDay) break;
        }
        if(subjectSameDay && !opts.allowSubjectSameDay) penalty += 20;

        // penalise teacher teaching same year more than once a day
        if(lesson.teacherId && !opts.allowTeacherMultiplePerDay){
          for(const yi of yearIndices){
            const cnt = teacherTeachesYearOnDay(grid, lesson.teacherId, yi, d);
            if(cnt > 0) penalty += 8 * cnt;
          }
        }

        // prefer earlier in week/day slightly
        penalty += d * 0.1 + s * 0.01;

        options.push({ day:d, slot:s, penalty, subjectSameDay });
      }
    }
    return options;
  }

  // Place a lesson on a grid (mutates grid). returns placedId.
  function placeLessonOnGrid(grid, lesson, day, slot){
    const placedId = uid('placed');
    const yrs = (lesson.years||[]).slice();
    for(const y of yrs){
      const yi = state.years.indexOf(y);
      if(yi>=0){
        grid[yi][day][slot] = {
          lessonId: lesson.id,
          subjectId: lesson.subjectId,
          teacherId: lesson.teacherId,
          roomId: lesson.roomId,
          mergedYears: yrs.length>1 ? yrs.slice() : null,
          placedId
        };
      }
    }
    return placedId;
  }

  function removePlacedFromGrid(grid, placedId){
    for(let y=0;y<grid.length;y++){
      for(let d=0; d<grid[y].length; d++){
        for(let s=0;s<grid[y][d].length;s++){
          if(grid[y][d][s] && grid[y][d][s].placedId === placedId) grid[y][d][s] = null;
        }
      }
    }
  }

  // Time-bounded MRV/backtracking solver
  function tryPlaceAllTokens(tokens, initialGrid, options = {}, timeoutMs = 4500){
    // Defensive copy of grid
    let grid;
    try {
      grid = JSON.parse(JSON.stringify(initialGrid || createEmptyGrid()));
    } catch(e){
      // JSON stringify may fail for huge structures; fallback to fresh empty grid
      grid = createEmptyGrid();
    }

    const start = Date.now();
    let timedOut = false;

    // We'll operate on a local copy of tokens for swapping
    const toks = tokens.map(t => ({ lesson: t.lesson, tokenUid: t.tokenUid }));

    const placements = [];

    function timeExceeded(){
      if(Date.now() - start > timeoutMs){
        timedOut = true;
        return true;
      }
      return false;
    }

    function recurse(idx){
      if(timeExceeded()) return false;

      if(idx >= toks.length){
        // success: copy local grid into state.grid (caller will use)
        return true;
      }

      // MRV selection among remaining tokens
      let bestIndex = -1;
      let bestOpts = null;
      for(let i=idx;i<toks.length;i++){
        const list = getOptionsWithPenalties(toks[i].lesson, grid, options);
        if(!list || list.length === 0){
          // can't place this token -> fail early
          return false;
        }
        if(bestOpts === null || list.length < bestOpts.length){
          bestOpts = list;
          bestIndex = i;
        }
      }

      if(bestIndex !== idx) {
        [toks[idx], toks[bestIndex]] = [toks[bestIndex], toks[idx]];
      }

      const token = toks[idx];
      let optList = getOptionsWithPenalties(token.lesson, grid, options);
      // sort by low penalty first
      optList.sort((a,b)=> a.penalty - b.penalty);

      for(const opt of optList){
        if(timeExceeded()) return false;
        // place
        const pid = placeLessonOnGrid(grid, token.lesson, opt.day, opt.slot);
        placements.push(pid);
        // recurse
        const ok = recurse(idx+1);
        if(ok) return true;
        // backtrack
        const last = placements.pop();
        removePlacedFromGrid(grid, last);
      }
      return false;
    }

    const success = recurse(0);
    return { success, timedOut, grid: success ? grid : grid };
  }

  // Helper to compute maximum available placements (quick impossibility check)
  function maxAvailablePlacements(){
    const slots = slotsCount();
    // per year/day count excluding Form slot for non-Form lessons, but conservative: allow all slots
    return state.years.length * state.days.length * slots;
  }

  // Auto-scheduler main function with progressive relaxation and safeguards
  async function autoSchedule(){
    try {
      if(!state.lessons.length){ await alertDialog('No lessons to schedule. Create lessons first.'); return; }
      const ok = await confirmDialog('Auto-schedule will clear existing assignments. Continue?');
      if(!ok) return;
      state.grid = createEmptyGrid();
      normalizeGrid();

      // Build token list: one token per placement (counts)
      const tokens = [];
      let totalNeeded = 0;
      state.lessons.forEach(l=>{
        for(let i=0;i<l.count;i++){
          tokens.push({ tokenUid: uid('tk'), lesson: l });
          totalNeeded++;
        }
      });

      // Quick capacity check
      const capacity = maxAvailablePlacements();
      if(totalNeeded > capacity){
        await alertDialog(`There are ${totalNeeded} lesson instances but only ${capacity} available slots across all years/days/periods. Increase days/periods or reduce counts before auto-scheduling.`);
        // continue anyway (we'll try to place as many as possible), but be explicit to user
      }

      // Shuffle tokens to reduce deterministic bias
      for(let i=tokens.length-1;i>0;i--){
        const j = Math.floor(Math.random()*(i+1)); [tokens[i],tokens[j]] = [tokens[j],tokens[i]];
      }

      // Try MRV/backtracking with time budget. We'll progressively relax soft constraints.
      const baseGrid = createEmptyGrid();
      let result;

      // Attempt 1: strict preferences
      result = tryPlaceAllTokens(tokens, baseGrid, { allowSubjectSameDay:false, allowTeacherMultiplePerDay:false }, 3500);
      if(result.timedOut || !result.success){
        // Attempt 2: relax subjectSameDay
        result = tryPlaceAllTokens(tokens, baseGrid, { allowSubjectSameDay:true, allowTeacherMultiplePerDay:false }, 2500);
      }
      if(result.timedOut || !result.success){
        // Attempt 3: relax teacher multiple per day
        result = tryPlaceAllTokens(tokens, baseGrid, { allowSubjectSameDay:true, allowTeacherMultiplePerDay:true }, 2000);
      }

      if(result.success){
        // commit grid
        state.grid = result.grid;
        normalizeGrid();
        renderAll();
      } else {
        // fallback greedy safe placement (single pass) with hard constraint checks, to avoid crashes/long runs
        const grid = createEmptyGrid();
        for(const tok of tokens){
          // generate options that satisfy hard constraints only
          const opts = getOptionsWithPenalties(tok.lesson, grid, { allowSubjectSameDay:true, allowTeacherMultiplePerDay:true });
          if(opts && opts.length){
            opts.sort((a,b)=> a.penalty - b.penalty);
            placeLessonOnGrid(grid, tok.lesson, opts[0].day, opts[0].slot);
          } else {
            // cannot place this token; skip - will remain unassigned
          }
        }
        state.grid = grid;
        normalizeGrid();
        renderAll();
      }

      // Final count of remaining unassigned instances
      const remainingCounts = {};
      state.lessons.forEach(l => remainingCounts[l.id] = (remainingCounts[l.id]||0) + l.count );
      const placedTracker = new Set();
      for(let y=0;y<state.grid.length;y++){
        for(let d=0; d<state.grid[y].length; d++){
          for(let s=0;s<state.grid[y][d].length; s++){
            const cell = state.grid[y][d][s];
            if(cell && cell.lessonId && !placedTracker.has(cell.placedId)){
              remainingCounts[cell.lessonId] = Math.max(0,(remainingCounts[cell.lessonId]||0)-1);
              placedTracker.add(cell.placedId);
            }
          }
        }
      }
      let left = 0;
      for(const k in remainingCounts) left += remainingCounts[k] || 0;

      if(left > 0){
        await alertDialog(`Auto-scheduler finished but ${left} lesson instances remain unassigned. The solver used a time-bounded search and then a fallback greedy pass. To improve coverage consider adding more periods/days, freeing rooms, or adjusting availability.`);
      } else {
        await alertDialog('Auto-scheduler finished. All lessons placed (where possible).');
      }
    } catch(err){
      console.error('Auto-schedule error', err);
      // Prevent crash — show an alert and keep current state
      await alertDialog('An unexpected error occurred during auto-schedule. See console for details.');
      renderAll();
    }
  }

  // ---------- End improved scheduler ----------

  // Export/Import/CSV/Excel (CSV)
  function exportJSON(){
    const data = {
      days: state.days, periods: state.periods, years: state.years,
      subjects: state.subjects, teachers: state.teachers, rooms: state.rooms, lessons: state.lessons, grid: state.grid
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], {type:'application/json'});
    const url = URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download='timetable_per_year_v2.json'; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  }
  function importJSONFile(file){
    const reader = new FileReader();
    reader.onload = async (ev)=>{
      try{
        const d = JSON.parse(ev.target.result);
        state.days = d.days || state.days; state.periods = d.periods || state.periods;
        state.years = d.years || state.years; state.subjects = d.subjects || []; state.teachers = d.teachers || []; state.rooms = d.rooms || []; state.lessons = d.lessons || []; state.grid = d.grid || createEmptyGrid();
        normalizeGrid();
        normalizeTeachersAvailability();
        renderAll();
      }catch(e){ await alertDialog('Invalid JSON'); console.error(e); }
    };
    reader.readAsText(file);
  }
  function exportGridCSV(){
    const slots = slotsCount();
    const header = ['Year'];
    state.days.forEach(day=>{
      for(let s=0;s<slots;s++){
        header.push(`${day} ${s===0?'Form':('P'+s)}`);
      }
    });
    let csv = header.join(',') + '\n';
    for(let y=0;y<state.years.length;y++){
      const row = [state.years[y]];
      for(let d=0; d<state.days.length; d++){
        for(let s=0; s<slots; s++){
          const cell = state.grid[y] && state.grid[y][d] && state.grid[y][d][s];
          if(cell){
            const subj = state.subjects.find(x=>x.id===cell.subjectId);
            const t = state.teachers.find(x=>x.id===cell.teacherId);
            const r = state.rooms.find(x=>x.id===cell.roomId);
            row.push(`"${(subj?subj.name:'(Form)')}${t?(' • '+t.code):''}${r?(' • '+r.name):''}"`);
          } else row.push('""');
        }
      }
      csv += row.join(',') + '\n';
    }
    const blob = new Blob([csv], {type:'text/csv'}); const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download='timetable_grid.csv'; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  }

  // Export Excel-friendly CSV (one row per placed lesson instance or remaining unassigned)
  function exportExcelCSV(){
    // Columns: lessonId, teacherId, subject, room, years, day, period (slot)
    const lines = [];
    lines.push(['lessonId','teacherId','subject','room','years','day','period'].join(','));
    const slots = slotsCount();
    // placed cells: iterate unique placedIds to produce one row per actual placement
    const placedSeen = new Set();
    for(let y=0;y<state.grid.length;y++){
      for(let d=0; d<state.grid[y].length; d++){
        for(let s=0; s<state.grid[y][d].length; s++){
          const cell = state.grid[y][d][s];
          if(cell && cell.placedId && !placedSeen.has(cell.placedId)){
            placedSeen.add(cell.placedId);
            const lesson = state.lessons.find(l=>l.id===cell.lessonId);
            const subject = state.subjects.find(su=>su.id===cell.subjectId);
            const teacher = state.teachers.find(t=>t.id===cell.teacherId);
            const room = state.rooms.find(r=>r.id===cell.roomId);
            const yearsStr = cell.mergedYears ? cell.mergedYears.join('+') : (lesson && lesson.years ? lesson.years.join('+') : '');
            lines.push([
              `"${cell.lessonId}"`,
              `"${cell.teacherId||''}"`,
              `"${subject?subject.name:''}"`,
              `"${room?room.name:''}"`,
              `"${yearsStr}"`,
              `"${state.days[d]}"`,
              `"${slotLabel(s)}"`
            ].join(','));
          }
        }
      }
    }
    // include unassigned tokens
    const remainingCounts = {};
    state.lessons.forEach(l => remainingCounts[l.id] = (remainingCounts[l.id]||0) + l.count );
    // subtract placed
    const placedTracker = new Set();
    for(let y=0;y<state.grid.length;y++){
      for(let d=0; d<state.grid[y].length; d++){
        for(let si=0; si<state.grid[y][d].length; si++){
          const cell = state.grid[y][d][si];
          if(cell && cell.lessonId && !placedTracker.has(cell.placedId)){
            remainingCounts[cell.lessonId] = Math.max(0,(remainingCounts[cell.lessonId]||0)-1);
            placedTracker.add(cell.placedId);
          }
        }
      }
    }
    state.lessons.forEach(l=>{
      const rem = remainingCounts[l.id] || 0;
      for(let i=0;i<rem;i++){
        const subject = state.subjects.find(su=>su.id===l.subjectId);
        const teacher = state.teachers.find(t=>t.id===l.teacherId);
        const room = state.rooms.find(r=>r.id===l.roomId);
        lines.push([
          `"${l.id}"`,
          `"${l.teacherId||''}"`,
          `"${subject?subject.name:''}"`,
          `"${room?room.name:''}"`,
          `"${l.years.join('+')}"`,
          '""',
          '""'
        ].join(','));
      }
    });

    const csv = lines.join('\n');
    const blob = new Blob([csv], {type:'text/csv'}); const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download='timetable_export.csv'; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  }

  // UI wiring
  function wireEvents(){
    el.btnApplySettings && el.btnApplySettings.addEventListener('click', async ()=>{
      const days = Array.from(el.selectDays.selectedOptions).map(o=>o.value);
      if(days.length<1){ await alertDialog('Pick at least one day'); return; }
      const p = Math.max(1, Math.min(12, parseInt(el.inputPeriods.value) || 5));
      state.days = days;
      state.periods = p;
      state.grid = createEmptyGrid();
      normalizeGrid();
      normalizeTeachersAvailability();
      renderAll();
    });

    el.btnAddTeacher && el.btnAddTeacher.addEventListener('click', async ()=>{
      const nm = el.teacherFullName.value.trim(); if(!nm) return;
      const existingCodes = state.teachers.map(t => t.code);
      const code = generateTeacherCode(nm, existingCodes);
      // default working availability: all true for configured days and slots
      const wd = {};
      const slots = slotsCount();
      state.days.forEach(d=> wd[d] = new Array(slots).fill(true));
      const t = { id: uid('t'), name: nm, code, workingAvailability: wd };
      state.teachers.push(t);
      el.teacherFullName.value='';
      renderAll();
    });

    el.btnAddSubject && el.btnAddSubject.addEventListener('click', ()=> {
      addSubjectModal();
    });

    el.btnEditSubject && el.btnEditSubject.addEventListener('click', ()=> {
      const sid = el.lessonSubject.value;
      if(sid) editSubjectModal(sid);
    });

    el.btnGenerateLessonsForYears && el.btnGenerateLessonsForYears.addEventListener && el.btnGenerateLessonsForYears.addEventListener('click', async ()=>{
      let selectedYears = Array.from(el.lessonYear.selectedOptions).map(o=>o.value);
      if(selectedYears.length === 0){
        const nodes = Array.from(el.yearsList.children).filter(n => n.classList.contains('active'));
        selectedYears = nodes.map(n => n.firstChild.textContent);
      }
      if(selectedYears.length===0){ await alertDialog('Select years (in Create Lesson or click years) first.'); return; }
      selectedYears.forEach(y=>{
        state.subjects.forEach(s=>{
          const exist = state.lessons.find(l => l.years.length === 1 && l.years[0] === y && l.subjectId === s.id);
          if(!exist){
            state.lessons.push({ id: uid('les'), years: [y], subjectId: s.id, teacherId: null, roomId: null, count: s.defaultCount });
          }
        });
      });
      renderAll();
    });

    // Add Lesson handler (robust)
    if(el.btnAddLesson){
      el.btnAddLesson.addEventListener('click', async ()=>{
        const selectedYears = el.lessonYear ? Array.from(el.lessonYear.selectedOptions).map(o=>o.value) : [];
        if(selectedYears.length === 0){
          await alertDialog('Select one or more years in the lesson Year selector to add lesson(s).'); return;
        }
        const subjectId = el.lessonSubject ? el.lessonSubject.value : null;
        const teacherId = el.lessonTeacher ? (el.lessonTeacher.value || null) : null;
        const roomId = el.lessonRoom ? (el.lessonRoom.value || null) : null;
        const count = el.lessonCount ? Math.max(1, Math.min(50, parseInt(el.lessonCount.value) || 1)) : 1;
        const mergedEl = $('#lessonMerged');
        const merged = mergedEl ? mergedEl.checked : false;
        if(!subjectId){
          await alertDialog('Please pick a subject before adding lessons.');
          return;
        }
        const subj = state.subjects.find(s=>s.id===subjectId);
        if(merged && selectedYears.length>1){
          if(!subj || !subj.mergeYears || !selectedYears.every(y => subj.mergeYears.includes(y))){
            await alertDialog('Merged lessons for these years are not allowed for this subject. Edit the subject merge settings first, or uncheck Merged.');
            return;
          }
        }
        if(merged && selectedYears.length>1){
          const exists = state.lessons.find(l => arraysEqual(l.years, selectedYears) && l.subjectId === subjectId && l.teacherId === teacherId && l.roomId === roomId);
          if(!exists){
            state.lessons.push({ id: uid('les'), years: selectedYears.slice(), subjectId, teacherId, roomId, count });
          } else {
            exists.count = (exists.count || 0) + count;
          }
        } else {
          selectedYears.forEach(year => {
            const exists = state.lessons.find(l => l.years.length===1 && l.years[0] === year && l.subjectId === subjectId && l.teacherId === teacherId && l.roomId === roomId);
            if(!exists){
              state.lessons.push({ id: uid('les'), years: [year], subjectId, teacherId, roomId, count });
            } else {
              exists.count = (exists.count || 0) + count;
            }
          });
        }
        renderAll();
      });
    }

    el.btnAutoSchedule && el.btnAutoSchedule.addEventListener('click', autoSchedule);
    el.btnClear && el.btnClear.addEventListener('click', async ()=> { const ok = await confirmDialog('Clear all assignments?'); if(ok){ state.grid = createEmptyGrid(); normalizeGrid(); renderAll(); }});
    el.btnResetData && el.btnResetData.addEventListener('click', async ()=> {
      const ok = await confirmDialog('Reset will REMOVE all teachers and lessons and clear assignments. Subjects and rooms will be kept. Continue?');
      if(!ok) return;
      state.teachers = [];
      state.lessons = [];
      state.grid = createEmptyGrid();
      normalizeGrid();
      save();
      renderAll();
    });
    el.btnExportJSON && el.btnExportJSON.addEventListener('click', exportJSON);
    el.btnImportJSON && el.btnImportJSON.addEventListener('click', ()=> el.fileImport.click());
    el.fileImport && el.fileImport.addEventListener('change', (ev)=>{ const f = ev.target.files[0]; if(f) importJSONFile(f); ev.target.value=''; });
    el.btnExportCSV && el.btnExportCSV.addEventListener('click', exportGridCSV);
    el.btnExportExcel && el.btnExportExcel.addEventListener('click', exportExcelCSV);
    el.btnPrint && el.btnPrint.addEventListener('click', ()=> window.print());

    el.filterType && el.filterType.addEventListener('change', (ev)=>{ state.filter.type = ev.target.value; state.filter.value = ''; renderFilterOptions(); });
    el.btnApplyFilter && el.btnApplyFilter.addEventListener('click', ()=>{ state.filter.type = el.filterType.value; state.filter.value = el.filterValue.value || ''; renderGrid(); });
    el.btnClearFilter && el.btnClearFilter.addEventListener('click', ()=>{ state.filter = { type: 'all', value: '' }; renderFilterOptions(); renderGrid(); });
  }

  // Teacher code generator
  function generateTeacherCode(fullName, existingCodes){
    const parts = fullName.trim().split(/\s+/);
    if(parts.length===0) return uid('T');
    const first = parts[0];
    const surname = parts.length>1 ? parts[parts.length-1] : first;
    const firstInitial = first[0].toUpperCase();
    const lastLetter = surname.slice(-1).toUpperCase();
    for(let n=0; n<Math.max(3, surname.length); n++){
      const nth = (n < surname.length) ? surname[n].toUpperCase() : surname[0].toUpperCase();
      const code = `${firstInitial}${nth}${lastLetter}`;
      if(!existingCodes.includes(code)) return code;
    }
    let i=1;
    while(true){
      const code = `${firstInitial}${surname[0].toUpperCase()}${lastLetter}${i}`;
      if(!existingCodes.includes(code)) return code;
      i++;
    }
  }

  // Edit teacher modal (with per-day per-slot availability and editable name/code, painting, quick actions)
  function editTeacherModal(id){
    const teacher = state.teachers.find(t=>t.id===id); if(!teacher) return;
    const modal = createModalContainer();
    const box = createModalBox(`Edit ${teacher.name} • ${teacher.code}`);

    const form = document.createElement('div');
    form.style.display = 'grid';
    form.style.gap = '8px';

    // Name / code
    const nameRow = document.createElement('div'); nameRow.style.display='flex'; nameRow.style.gap='8px';
    const nameInput = document.createElement('input'); nameInput.value = teacher.name; nameInput.style.flex='1';
    const codeInput = document.createElement('input'); codeInput.value = teacher.code; codeInput.style.width='110px';
    nameRow.appendChild(nameInput); nameRow.appendChild(codeInput);
    form.appendChild(nameRow);

    // Quick action buttons
    const quickRow = document.createElement('div'); quickRow.style.display='flex'; quickRow.style.gap='8px';
    const btnAllOn = document.createElement('button'); btnAllOn.textContent = 'All ON'; stylePrimaryButton(btnAllOn);
    const btnAllOff = document.createElement('button'); btnAllOff.textContent = 'All OFF'; btnAllOff.style.background='#b0b0b0'; btnAllOff.style.color='#fff';
    const btnInvert = document.createElement('button'); btnInvert.textContent = 'Invert'; btnInvert.style.background='#f39c12'; btnInvert.style.color='#fff';
    quickRow.appendChild(btnAllOn); quickRow.appendChild(btnAllOff); quickRow.appendChild(btnInvert);

    // Copy from other teacher dropdown
    const copyRow = document.createElement('div'); copyRow.style.display='flex'; copyRow.style.gap='8px'; copyRow.style.alignItems='center';
    const copyLabel = document.createElement('div'); copyLabel.textContent = 'Copy from:'; copyLabel.style.fontSize='13px';
    const copySelect = document.createElement('select'); copySelect.style.flex='1';
    const noneOpt = document.createElement('option'); noneOpt.value=''; noneOpt.textContent='(choose teacher)'; copySelect.appendChild(noneOpt);
    state.teachers.forEach(t=>{
      if(t.id !== teacher.id){
        const o = document.createElement('option'); o.value = t.id; o.textContent = t.name + ' • ' + t.code; copySelect.appendChild(o);
      }
    });
    const copyBtn = document.createElement('button'); copyBtn.textContent = 'Copy'; copyBtn.style.background='#4caf50'; copyBtn.style.color='#fff';
    copyRow.appendChild(copyLabel); copyRow.appendChild(copySelect); copyRow.appendChild(copyBtn);

    form.appendChild(quickRow);
    form.appendChild(copyRow);

    // Availability grid
    const slots = slotsCount();
    const availTitle = document.createElement('div'); availTitle.textContent = 'Working availability (click or click-drag to paint periods the teacher CAN attend):'; availTitle.style.fontWeight='700';
    form.appendChild(availTitle);
    const table = document.createElement('table'); table.style.borderCollapse='collapse'; table.style.width='100%';
    const thead = document.createElement('thead'); const trh = document.createElement('tr');
    const cornerTh = document.createElement('th'); cornerTh.textContent = ''; cornerTh.style.padding='6px';
    trh.appendChild(cornerTh);
    state.days.forEach(d => {
      const th = document.createElement('th'); th.style.padding='6px'; th.style.textAlign='center'; th.style.cursor='pointer';
      th.textContent = d;
      th.title = 'Click to toggle full day';
      th.addEventListener('click', ()=> {
        // toggle this day: set to full if currently any false, else clear
        const current = teacher.workingAvailability && teacher.workingAvailability[d];
        const anyFalse = current && current.some(v=>!v);
        teacher.workingAvailability[d] = new Array(slots).fill(!!anyFalse);
        // refresh checkboxes
        [...table.querySelectorAll(`input[data-day="${d}"]`)].forEach(cb => cb.checked = !!anyFalse);
      });
      trh.appendChild(th);
    });
    thead.appendChild(trh); table.appendChild(thead);
    const tbody = document.createElement('tbody');

    // painting state
    let painting = false;
    let paintValue = true; // true=turn on, false=turn off

    for(let s=0;s<slots;s++){
      const tr = document.createElement('tr');
      const th = document.createElement('th'); th.style.padding='6px'; th.textContent = (s===0?'Form':('P'+s)); tr.appendChild(th);
      state.days.forEach(d=>{
        const td = document.createElement('td'); td.style.padding='6px'; td.style.textAlign='center';
        const cb = document.createElement('input'); cb.type='checkbox';
        cb.checked = !!(teacher.workingAvailability && teacher.workingAvailability[d] && teacher.workingAvailability[d][s]);
        cb.dataset.day = d; cb.dataset.slot = s;
        cb.addEventListener('mousedown', (ev)=>{
          ev.preventDefault();
          painting = true;
          paintValue = !cb.checked; // toggle mode relative to current value
          cb.checked = paintValue;
          teacher.workingAvailability[d][s] = paintValue;
        });
        cb.addEventListener('mouseenter', ()=>{
          if(painting){
            cb.checked = paintValue;
            teacher.workingAvailability[d][s] = paintValue;
          }
        });
        // support click toggle (without drag)
        cb.addEventListener('click', ()=>{
          teacher.workingAvailability[d][s] = cb.checked;
        });
        td.appendChild(cb); tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    form.appendChild(table);

    // End painting on mouseup anywhere
    const onMouseUp = ()=> { painting = false; };
    document.addEventListener('mouseup', onMouseUp);

    // Quick action handlers
    btnAllOn.addEventListener('click', ()=>{
      state.days.forEach(d => teacher.workingAvailability[d] = new Array(slots).fill(true));
      // refresh table
      table.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = true);
    });
    btnAllOff.addEventListener('click', ()=>{
      state.days.forEach(d => teacher.workingAvailability[d] = new Array(slots).fill(false));
      table.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = false);
    });
    btnInvert.addEventListener('click', ()=>{
      state.days.forEach(d => {
        teacher.workingAvailability[d] = teacher.workingAvailability[d].map(v=>!v);
      });
      table.querySelectorAll('input[type=checkbox]').forEach(cb => {
        const d = cb.dataset.day; const s = parseInt(cb.dataset.slot);
        cb.checked = !!teacher.workingAvailability[d][s];
      });
    });

    // Copy from teacher
    copyBtn.addEventListener('click', ()=>{
      const fromId = copySelect.value;
      if(!fromId) return;
      const from = state.teachers.find(t=>t.id===fromId);
      if(!from || !from.workingAvailability) return;
      // deep copy availability
      const avail = {};
      state.days.forEach(d => avail[d] = (from.workingAvailability[d]||[]).slice(0, slotsCount()).map(v=>!!v));
      teacher.workingAvailability = avail;
      // refresh checkboxes
      table.querySelectorAll('input[type=checkbox]').forEach(cb=>{
        const d = cb.dataset.day; const s = parseInt(cb.dataset.slot);
        cb.checked = !!teacher.workingAvailability[d][s];
      });
    });

    // Buttons
    const btnRow = document.createElement('div'); btnRow.style.display='flex'; btnRow.style.justifyContent='flex-end'; btnRow.style.gap='8px'; btnRow.style.marginTop='8px';
    const saveBtn = document.createElement('button'); saveBtn.textContent='Save'; stylePrimaryButton(saveBtn);
    const cancelBtn = document.createElement('button'); cancelBtn.textContent='Cancel';
    saveBtn.addEventListener('click', ()=>{
      teacher.name = nameInput.value.trim() || teacher.name;
      teacher.code = codeInput.value.trim() || teacher.code;
      // ensure availability structure complete
      normalizeTeachersAvailability();
      document.body.removeChild(modal);
      document.removeEventListener('mouseup', onMouseUp);
      renderAll();
    });
    cancelBtn.addEventListener('click', ()=> { document.body.removeChild(modal); document.removeEventListener('mouseup', onMouseUp); });
    btnRow.appendChild(cancelBtn); btnRow.appendChild(saveBtn);

    box.appendChild(form); box.appendChild(btnRow); modal.appendChild(box);
    document.body.appendChild(modal);
  }

  // Subject add/edit modal
  function addSubjectModal(){
    const modal = createModalContainer();
    const box = createModalBox('Add Subject');
    const form = document.createElement('div'); form.style.display='grid'; form.style.gap='8px';
    const name = document.createElement('input'); name.placeholder='Subject name';
    const count = document.createElement('input'); count.type='number'; count.min=0; count.value=1;
    const color = document.createElement('input'); color.type='color'; color.value='#4caf50';
    const mergeLabel = document.createElement('div'); mergeLabel.textContent = 'Select years that may be merged for this subject (optional):';
    const mergeContainer = document.createElement('div'); mergeContainer.style.display='flex'; mergeContainer.style.flexWrap='wrap'; mergeContainer.style.gap='6px';
    state.years.forEach(y=>{
      const cb = document.createElement('label'); cb.style.display='flex'; cb.style.alignItems='center'; cb.style.gap='6px';
      const input = document.createElement('input'); input.type='checkbox'; input.value = y;
      cb.appendChild(input); cb.appendChild(document.createTextNode(y));
      mergeContainer.appendChild(cb);
    });
    form.appendChild(name); form.appendChild(count); form.appendChild(color); form.appendChild(mergeLabel); form.appendChild(mergeContainer);
    const btnRow = document.createElement('div'); btnRow.style.display='flex'; btnRow.style.justifyContent='flex-end'; btnRow.style.gap='8px';
    const addBtn = document.createElement('button'); addBtn.textContent='Add'; stylePrimaryButton(addBtn);
    const cancelBtn = document.createElement('button'); cancelBtn.textContent='Cancel';
    addBtn.addEventListener('click', ()=>{
      const nm = name.value.trim(); if(!nm) return;
      const c = Math.max(0, parseInt(count.value) || 0);
      const col = color.value;
      const mergeYears = Array.from(mergeContainer.querySelectorAll('input[type=checkbox]:checked')).map(i=>i.value);
      state.subjects.push({ id: uid('sub'), name: nm, color: col, defaultCount: c, mergeYears });
      document.body.removeChild(modal);
      renderAll();
    });
    cancelBtn.addEventListener('click', ()=> document.body.removeChild(modal));
    btnRow.appendChild(cancelBtn); btnRow.appendChild(addBtn);
    box.appendChild(form); box.appendChild(btnRow); modal.appendChild(box);
    document.body.appendChild(modal);
  }

  function editSubjectModal(id){
    const s = state.subjects.find(x=>x.id===id); if(!s) return;
    const modal = createModalContainer();
    const box = createModalBox(`Edit Subject: ${s.name}`);
    const form = document.createElement('div'); form.style.display='grid'; form.style.gap='8px';
    const name = document.createElement('input'); name.value = s.name;
    const count = document.createElement('input'); count.type='number'; count.min=0; count.value=s.defaultCount;
    const color = document.createElement('input'); color.type='color'; color.value=s.color || '#4caf50';
    const mergeLabel = document.createElement('div'); mergeLabel.textContent = 'Years allowed to be merged for this subject (leave blank to disallow merges):';
    const mergeContainer = document.createElement('div'); mergeContainer.style.display='flex'; mergeContainer.style.flexWrap='wrap'; mergeContainer.style.gap='6px';
    state.years.forEach(y=>{
      const cb = document.createElement('label'); cb.style.display='flex'; cb.style.alignItems='center'; cb.style.gap='6px';
      const input = document.createElement('input'); input.type='checkbox'; input.value = y;
      if(s.mergeYears && s.mergeYears.includes(y)) input.checked = true;
      cb.appendChild(input); cb.appendChild(document.createTextNode(y));
      mergeContainer.appendChild(cb);
    });
    form.appendChild(name); form.appendChild(count); form.appendChild(color); form.appendChild(mergeLabel); form.appendChild(mergeContainer);
    const btnRow = document.createElement('div'); btnRow.style.display='flex'; btnRow.style.justifyContent='flex-end'; btnRow.style.gap='8px';
    const saveBtn = document.createElement('button'); saveBtn.textContent='Save'; stylePrimaryButton(saveBtn);
    const cancelBtn = document.createElement('button'); cancelBtn.textContent='Cancel';
    saveBtn.addEventListener('click', ()=>{
      s.name = name.value.trim() || s.name;
      s.defaultCount = Math.max(0, parseInt(count.value) || 0);
      s.color = color.value;
      s.mergeYears = Array.from(mergeContainer.querySelectorAll('input[type=checkbox]:checked')).map(i=>i.value);
      document.body.removeChild(modal);
      renderAll();
    });
    cancelBtn.addEventListener('click', ()=> document.body.removeChild(modal));
    btnRow.appendChild(cancelBtn); btnRow.appendChild(saveBtn);
    box.appendChild(form); box.appendChild(btnRow); modal.appendChild(box);
    document.body.appendChild(modal);
  }

  // small util
  function arraysEqual(a,b){
    if(!a || !b) return false;
    if(a.length !== b.length) return false;
    const aa = a.slice().sort(); const bb = b.slice().sort();
    for(let i=0;i<aa.length;i++) if(aa[i] !== bb[i]) return false;
    return true;
  }

  // Expose remove placed function to buttons
  window._timetableState = state;
  window.renderTimetable = renderAll;

  // Init
  load();
  wireEvents();
  renderAll();

})();
