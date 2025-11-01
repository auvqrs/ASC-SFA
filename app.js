// Per-year timetable app (patched): normalize grid to avoid missing/invalid arrays which cause render errors.
// Replace your existing app.js with this file and reload the page.

(function(){
  // State
  const state = {
    days: ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'],
    periods: 5, // P1..P5 (slot 0 is Form)
    years: ['Y7','Y8','Y9','Y10','Y11','LSU','SF'],
    subjects: [], // {id,name,color,defaultCount}
    teachers: [], // {id, name, code, workingDays: {Mon:true,...}}
    rooms: [], // {id,name}
    lessons: [], // {id, year, subjectId, teacherId, roomId, count}
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
    btnExportJSON: $('#btnExportJSON'),
    btnImportJSON: $('#btnImportJSON'),
    fileImport: $('#fileImport'),
    btnExportCSV: $('#btnExportCSV'),
    btnPrint: $('#btnPrint'),
    filterType: $('#filterType'),
    filterValue: $('#filterValue'),
    btnApplyFilter: $('#btnApplyFilter'),
    btnClearFilter: $('#btnClearFilter')
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

  // Storage
  function save() {
    const data = {
      days: state.days, periods: state.periods, years: state.years,
      subjects: state.subjects, teachers: state.teachers, rooms: state.rooms,
      lessons: state.lessons, grid: state.grid
    };
    localStorage.setItem('school_timetable_per_year_v1', JSON.stringify(data));
  }
  function load(){
    try {
      const raw = localStorage.getItem('school_timetable_per_year_v1');
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
    state.subjects = defaultSubjects.map(s=>({ id: uid('sub'), name: s[0], color: randomColor(), defaultCount: s[1] }));
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

  // Grid: years × days × slots (slots = 1 + periods)
  function createEmptyGrid(){
    const slots = 1 + state.periods; // slot0 = Form
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

  // New: ensure state.grid always matches years × days × slots shape
  function normalizeGrid(){
    // ensure grid is array
    if(!Array.isArray(state.grid)) state.grid = createEmptyGrid();
    const slots = 1 + state.periods;
    // ensure number of year rows
    while(state.grid.length < state.years.length) state.grid.push([]);
    if(state.grid.length > state.years.length) state.grid.splice(state.years.length);
    // for each year ensure days and slot arrays are correct
    for(let y=0; y<state.years.length; y++){
      if(!Array.isArray(state.grid[y])) state.grid[y] = [];
      // ensure days
      while(state.grid[y].length < state.days.length) state.grid[y].push(new Array(slots).fill(null));
      if(state.grid[y].length > state.days.length) state.grid[y].splice(state.days.length);
      // ensure each day has right slot length
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

  // Rendering (wrapped with try/catch to log errors and avoid silent failure)
  function renderAll(){
    try {
      normalizeGrid();
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
      // keep UI alive and show message in console
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
      const div = document.createElement('div'); div.className='list-item';
      div.innerHTML = `<div><strong style="color:${s.color}">${s.name}</strong><div class="small">per week: ${s.defaultCount}</div></div><div>${s.defaultCount}</div>`;
      container.appendChild(div);
    });
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
      left.innerHTML = `<strong>${t.name}</strong> <span class="teacher-badge">${t.code}</span><div class="small">Working days: ${formatWorkingDays(t)}</div>`;
      const right = document.createElement('div');
      right.innerHTML = `<button data-id="${t.id}" class="edit-teacher">Edit</button> <button data-id="${t.id}" class="remove-teacher">Remove</button>`;
      d.appendChild(left); d.appendChild(right);
      c.appendChild(d);
    });
    $$('.edit-teacher', c).forEach(btn=> btn.addEventListener('click', ()=> editTeacherModal(btn.dataset.id)));
    $$('.remove-teacher', c).forEach(btn=> btn.addEventListener('click', ()=>{
      const id = btn.dataset.id;
      if(confirm('Remove teacher and their lessons?')){
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
      }
    }));
  }

  function formatWorkingDays(t){
    return state.days.filter(d=>t.workingDays && t.workingDays[d]).join(', ') || '—';
  }

  function renderLessonSelectors(){
    // year select multi
    el.lessonYear.innerHTML=''; state.years.forEach(y=>{ const o=document.createElement('option'); o.value=y; o.textContent=y; el.lessonYear.appendChild(o); });
    el.lessonSubject.innerHTML=''; state.subjects.forEach(s=>{ const o=document.createElement('option'); o.value=s.id; o.textContent=`${s.name} (${s.defaultCount}/wk)`; el.lessonSubject.appendChild(o); });
    el.lessonTeacher.innerHTML=''; const blank=document.createElement('option'); blank.value=''; blank.textContent='(none)'; el.lessonTeacher.appendChild(blank);
    state.teachers.forEach(t=>{ const o=document.createElement('option'); o.value=t.id; o.textContent=`${t.name} • ${t.code}`; el.lessonTeacher.appendChild(o); });
    el.lessonRoom.innerHTML=''; const none=document.createElement('option'); none.value=''; none.textContent='(none)'; el.lessonRoom.appendChild(none);
    state.rooms.forEach(r=>{ const o=document.createElement('option'); o.value=r.id; o.textContent=r.name; el.lessonRoom.appendChild(o); });
  }

  // Tokens - unassigned lesson instances (respect year)
  function renderTokens(){
    const c = el.tokens; c.innerHTML='';
    const counts = {};
    state.lessons.forEach(l => counts[l.id] = (counts[l.id]||0) + l.count );
    // subtract placements across grid (placed lessonId)
    for(let y=0;y<state.grid.length;y++){
      for(let d=0; d<state.grid[y].length; d++){
        for(let s=0;s<state.grid[y][d].length;s++){
          const cell = state.grid[y][d][s];
          if(cell && cell.lessonId) counts[cell.lessonId] = Math.max(0,(counts[cell.lessonId]||0)-1);
        }
      }
    }
    // Render tokens by lesson remaining (year shows inside token)
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
        div.innerHTML = `<div style="display:flex;flex-direction:column"><div style="font-weight:700">${sub?sub.name:'?'}</div><div style="font-size:12px">${l.year} ${teacher?('• '+teacher.code):''}${room?(' • '+room.name):''}</div></div>`;
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
    const slots = 1 + state.periods; // slot0 = Form
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
            // FIX: use 'sub' consistently (previously declared subj but referenced as sub)
            const sub = state.subjects.find(x=>x.id===assigned.subjectId);
            const t = state.teachers.find(x=>x.id===assigned.teacherId);
            const r = state.rooms.find(x=>x.id===assigned.roomId);
            const wrapper = document.createElement('div');
            wrapper.innerHTML = `<div class="assign" style="background:${sub?sub.color:'#999'};color:white;padding:4px;border-radius:4px">${sub?sub.name:'(Form)'}</div>
              <div class="meta">${t?('<span class="teacher-badge">'+t.code+'</span>') : ''} ${r?(' • '+r.name):''}</div>`;
            td.appendChild(wrapper);

            const btn = document.createElement('button');
            btn.textContent='✕'; btn.title='Remove assignment';
            btn.style.position='absolute'; btn.style.right='6px'; btn.style.top='6px';
            btn.style.border='none'; btn.style.background='rgba(0,0,0,0.06)'; btn.style.borderRadius='4px'; btn.style.cursor='pointer';
            btn.addEventListener('click',(ev)=>{ ev.stopPropagation(); state.grid[y][d][s] = null; renderAll(); });
            td.appendChild(btn);
          } else {
            td.innerHTML = '<div class="small">empty</div>';
          }

          // filter dimming
          if(!cellMatchesFilter(assigned, y)){
            td.classList.add('dimmed');
          } else td.classList.remove('dimmed');

          // teacher working day conflict highlight
          if(assigned && assigned.teacherId){
            const teacher = state.teachers.find(tt=>tt.id===assigned.teacherId);
            const dayName = state.days[d];
            if(teacher && (!teacher.workingDays || !teacher.workingDays[dayName])){
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
    e.dataTransfer.setData('text/plain', draggingLessonId);
  }
  function cellDragOver(e){ e.preventDefault(); e.currentTarget.classList.add('drag-over'); }
  function cellDragLeave(e){ e.currentTarget.classList.remove('drag-over'); }
  function cellDrop(e){
    e.preventDefault();
    e.currentTarget.classList.remove('drag-over');
    const lessonId = e.dataTransfer.getData('text/plain') || draggingLessonId || state.pickedLessonId;
    if(!lessonId) return;
    const y = +e.currentTarget.dataset.yearIndex;
    const d = +e.currentTarget.dataset.dayIndex;
    const s = +e.currentTarget.dataset.slotIndex;
    const lesson = state.lessons.find(l=>l.id===lessonId);
    if(!lesson) return;
    if(lesson.year !== state.years[y]){
      alert(`This lesson is for ${lesson.year}. You can only place it in the ${lesson.year} row.`);
      return;
    }
    assignLessonToCell(lessonId, y, d, s);
    draggingLessonId = null;
    state.pickedLessonId = null;
    renderAll();
  }

  function assignLessonToCell(lessonId, yearIndex, dayIndex, slotIndex){
    const lesson = state.lessons.find(l=>l.id===lessonId);
    if(!lesson) return;
    if(lesson.teacherId){
      const teacher = state.teachers.find(t=>t.id===lesson.teacherId);
      const dayName = state.days[dayIndex];
      if(teacher && !(teacher.workingDays && teacher.workingDays[dayName])){
        if(!confirm(`Teacher ${teacher.name} does not work ${dayName}. Place anyway?`)) return;
      }
    }
    // normalize grid to be safe before assignment
    normalizeGrid();
    state.grid[yearIndex][dayIndex][slotIndex] = {
      lessonId: lesson.id,
      subjectId: lesson.subjectId,
      teacherId: lesson.teacherId,
      roomId: lesson.roomId,
      year: lesson.year,
      placedId: uid('placed')
    };
  }

  // Cell click quick assign
  function cellClick(e){
    const y = +e.currentTarget.dataset.yearIndex;
    const d = +e.currentTarget.dataset.dayIndex;
    const s = +e.currentTarget.dataset.slotIndex;
    if(state.pickedLessonId){
      const lesson = state.lessons.find(l=>l.id===state.pickedLessonId);
      if(lesson && lesson.year !== state.years[y]){ alert(`Picked lesson is for ${lesson.year}. Click a ${lesson.year} row.`); return; }
      assignLessonToCell(state.pickedLessonId, y, d, s);
      state.pickedLessonId = null; renderAll(); return;
    }
    quickAssignPopup(e.currentTarget, y, d, s);
  }

  function quickAssignPopup(cellEl, yearIndex, dayIndex, slotIndex){
    const remainingCounts = {};
    state.lessons.forEach(l => remainingCounts[l.id] = (remainingCounts[l.id]||0) + l.count );
    for(let y=0;y<state.grid.length;y++){
      for(let d=0; d<state.grid[y].length; d++){
        for(let si=0; si<state.grid[y][d].length; si++){
          const cell = state.grid[y][d][si];
          if(cell && cell.lessonId) remainingCounts[cell.lessonId] = Math.max(0,(remainingCounts[cell.lessonId]||0)-1);
        }
      }
    }

    const available = state.lessons.filter(l => l.year === state.years[yearIndex] && (remainingCounts[l.id]||0) > 0);

    const menu = document.createElement('div');
    menu.style.position='absolute'; menu.style.left='10px'; menu.style.top='10px';
    menu.style.zIndex=1000; menu.style.background='white'; menu.style.border='1px solid #e2e8f0';
    menu.style.borderRadius='6px'; menu.style.padding='8px'; menu.style.boxShadow='0 8px 24px rgba(0,0,0,0.08)';
    const title = document.createElement('div'); title.textContent=`Assign for ${state.years[yearIndex]} (${state.days[dayIndex]} ${slotIndex===0?'Form':('P'+slotIndex)})`; title.style.fontWeight='700'; title.style.marginBottom='6px'; menu.appendChild(title);

    if(available.length === 0){
      const none = document.createElement('div'); none.textContent = 'No assignable lessons remaining for this year.'; menu.appendChild(none);
    } else {
      available.slice(0,300).forEach(l=>{
        const s = state.subjects.find(x=>x.id===l.subjectId);
        const t = state.teachers.find(x=>x.id===l.teacherId);
        const row = document.createElement('div'); row.style.marginBottom='6px';
        const btn = document.createElement('button'); btn.textContent = `${l.year} • ${s? s.name : '?'} ${t?(' • '+t.code):''}`; btn.style.width='100%';
        btn.addEventListener('click', ()=>{ assignLessonToCell(l.id, yearIndex, dayIndex, slotIndex); document.body.removeChild(menu); renderAll(); });
        row.appendChild(btn); menu.appendChild(row);
      });
    }
    const cancel = document.createElement('button'); cancel.textContent='Close'; cancel.style.marginTop='6px'; cancel.addEventListener('click', ()=>document.body.removeChild(menu)); menu.appendChild(cancel);
    document.body.appendChild(menu);

    const rect = cellEl.getBoundingClientRect();
    menu.style.left = Math.min(window.innerWidth-320, rect.left + 12) + 'px';
    menu.style.top = Math.min(window.innerHeight-220, rect.top + 12) + 'px';
  }

  // Auto-scheduler
  function autoSchedule(){
    if(!state.lessons.length){ alert('No lessons to schedule. Create lessons first.'); return; }
    if(!confirm('Auto-schedule will clear existing assignments. Continue?')) return;
    state.grid = createEmptyGrid();
    normalizeGrid();
    const tokens = [];
    state.lessons.forEach(l=>{
      for(let i=0;i<l.count;i++) tokens.push(Object.assign({ tokenUid: uid('tk') }, l));
    });
    for(let i=tokens.length-1;i>0;i--){
      const j = Math.floor(Math.random()*(i+1)); [tokens[i],tokens[j]] = [tokens[j],tokens[i]];
    }
    const slots = 1 + state.periods;
    const yearIndexOf = y => state.years.indexOf(y);

    for(let s=0;s<slots;s++){
      for(let d=0; d<state.days.length; d++){
        for(let yi=0; yi<state.years.length; yi++){
          for(let ti=0; ti<tokens.length; ti++){
            const tok = tokens[ti];
            const targetYearIdx = yearIndexOf(tok.year);
            if(targetYearIdx !== yi) continue;
            const teacher = state.teachers.find(t=>t.id===tok.teacherId);
            const dayName = state.days[d];
            if(teacher && (!teacher.workingDays || !teacher.workingDays[dayName])) continue;
            let teacherConflict = false;
            for(let yj=0; yj<state.grid.length; yj++){
              if(state.grid[yj][d] && state.grid[yj][d][s] && state.grid[yj][d][s].teacherId && tok.teacherId && state.grid[yj][d][s].teacherId === tok.teacherId){
                teacherConflict = true; break;
              }
            }
            if(teacherConflict) continue;
            state.grid[yi][d][s] = {
              lessonId: tok.id,
              subjectId: tok.subjectId,
              teacherId: tok.teacherId,
              roomId: tok.roomId,
              year: tok.year,
              placedId: tok.tokenUid
            };
            tokens.splice(ti,1);
            break;
          }
        }
      }
    }

    renderAll();
  }

  // Export/Import/CSV
  function exportJSON(){
    const data = {
      days: state.days, periods: state.periods, years: state.years,
      subjects: state.subjects, teachers: state.teachers, rooms: state.rooms, lessons: state.lessons, grid: state.grid
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], {type:'application/json'});
    const url = URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download='timetable_per_year.json'; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  }
  function importJSONFile(file){
    const reader = new FileReader();
    reader.onload = (ev)=>{
      try{
        const d = JSON.parse(ev.target.result);
        state.days = d.days || state.days; state.periods = d.periods || state.periods;
        state.years = d.years || state.years; state.subjects = d.subjects || []; state.teachers = d.teachers || []; state.rooms = d.rooms || []; state.lessons = d.lessons || []; state.grid = d.grid || createEmptyGrid();
        normalizeGrid();
        renderAll();
      }catch(e){ alert('Invalid JSON'); console.error(e); }
    };
    reader.readAsText(file);
  }
  function exportCSV(){
    const slots = 1 + state.periods;
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
    const blob = new Blob([csv], {type:'text/csv'}); const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download='timetable_per_year.csv'; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  }

  // UI wiring
  function wireEvents(){
    el.btnApplySettings.addEventListener('click', ()=>{
      const days = Array.from(el.selectDays.selectedOptions).map(o=>o.value);
      if(days.length<1){ alert('Pick at least one day'); return; }
      const p = Math.max(1, Math.min(12, parseInt(el.inputPeriods.value) || 5));
      state.days = days;
      state.periods = p;
      state.grid = createEmptyGrid();
      normalizeGrid();
      renderAll();
    });

    el.btnAddTeacher.addEventListener('click', ()=>{
      const nm = el.teacherFullName.value.trim(); if(!nm) return;
      const existingCodes = state.teachers.map(t => t.code);
      const code = generateTeacherCode(nm, existingCodes);
      // default working days Mon-Sun
      const wd = {}; state.days.forEach(d=> wd[d] = true);
      const t = { id: uid('t'), name: nm, code, workingDays: wd };
      state.teachers.push(t);
      el.teacherFullName.value='';
      renderAll();
    });

    el.btnGenerateLessonsForYears.addEventListener('click', ()=>{
      let selectedYears = Array.from(el.lessonYear.selectedOptions).map(o=>o.value);
      if(selectedYears.length === 0){
        const nodes = Array.from(el.yearsList.children).filter(n => n.classList.contains('active'));
        selectedYears = nodes.map(n => n.firstChild.textContent);
      }
      if(selectedYears.length===0){ alert('Select years (in Create Lesson or click years) first.'); return; }
      selectedYears.forEach(y=>{
        state.subjects.forEach(s=>{
          const exist = state.lessons.find(l => l.year === y && l.subjectId === s.id);
          if(!exist){
            state.lessons.push({ id: uid('les'), year: y, subjectId: s.id, teacherId: null, roomId: null, count: s.defaultCount });
          }
        });
      });
      renderAll();
    });

    el.btnAddLesson.addEventListener('click', ()=>{
      const selectedYears = Array.from(el.lessonYear.selectedOptions).map(o=>o.value);
      if(selectedYears.length === 0){
        alert('Select one or more years in the lesson Year selector to add lesson(s).'); return;
      }
      const subjectId = el.lessonSubject.value;
      const teacherId = el.lessonTeacher.value || null;
      const roomId = el.lessonRoom.value || null;
      const count = Math.max(1, Math.min(50, parseInt(el.lessonCount.value) || 1));
      selectedYears.forEach(year => {
        const exists = state.lessons.find(l => l.year === year && l.subjectId === subjectId && l.teacherId === teacherId && l.roomId === roomId);
        if(!exists){
          state.lessons.push({ id: uid('les'), year, subjectId, teacherId, roomId, count });
        }
      });
      renderAll();
    });

    el.btnAutoSchedule.addEventListener('click', autoSchedule);
    el.btnClear.addEventListener('click', ()=> { if(confirm('Clear all assignments?')){ state.grid = createEmptyGrid(); normalizeGrid(); renderAll(); }});
    el.btnExportJSON.addEventListener('click', exportJSON);
    el.btnImportJSON.addEventListener('click', ()=> el.fileImport.click());
    el.fileImport.addEventListener('change', (ev)=>{ const f = ev.target.files[0]; if(f) importJSONFile(f); ev.target.value=''; });
    el.btnExportCSV.addEventListener('click', exportCSV);
    el.btnPrint.addEventListener('click', ()=> window.print());

    el.filterType.addEventListener('change', (ev)=>{ state.filter.type = ev.target.value; state.filter.value = ''; renderFilterOptions(); });
    el.btnApplyFilter.addEventListener('click', ()=>{ state.filter.type = el.filterType.value; state.filter.value = el.filterValue.value || ''; renderGrid(); });
    el.btnClearFilter.addEventListener('click', ()=>{ state.filter = { type: 'all', value: '' }; renderFilterOptions(); renderGrid(); });
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

  // Edit teacher modal
  function editTeacherModal(id){
    const teacher = state.teachers.find(t=>t.id===id); if(!teacher) return;
    const modal = document.createElement('div'); modal.style.position='fixed'; modal.style.left='0'; modal.style.top='0'; modal.style.right='0'; modal.style.bottom='0'; modal.style.background='rgba(0,0,0,0.25)';
    const box = document.createElement('div'); box.style.width='420px'; box.style.margin='80px auto'; box.style.background='white'; box.style.padding='16px'; box.style.borderRadius='8px';
    box.innerHTML = `<h3>Edit ${teacher.name} • ${teacher.code}</h3>`;
    const form = document.createElement('div');
    state.days.forEach(d=>{
      const label = document.createElement('label'); label.style.display='block'; label.style.marginBottom='6px';
      const cb = document.createElement('input'); cb.type='checkbox'; cb.checked = !!teacher.workingDays[d]; cb.dataset.day = d;
      label.appendChild(cb); label.appendChild(document.createTextNode(' ' + d));
      form.appendChild(label);
    });
    box.appendChild(form);
    const saveBtn = document.createElement('button'); saveBtn.textContent='Save'; saveBtn.style.marginRight='8px';
    const cancelBtn = document.createElement('button'); cancelBtn.textContent='Cancel';
    saveBtn.addEventListener('click', ()=>{
      const boxes = form.querySelectorAll('input[type=checkbox]');
      boxes.forEach(cb => teacher.workingDays[cb.dataset.day] = cb.checked);
      document.body.removeChild(modal);
      renderAll();
    });
    cancelBtn.addEventListener('click', ()=> document.body.removeChild(modal));
    box.appendChild(saveBtn); box.appendChild(cancelBtn);
    modal.appendChild(box); document.body.appendChild(modal);
  }

  // Expose state for debugging
  window._timetableState = state;

  // Init
  load();
  wireEvents();
  renderAll();

  // expose render
  window.renderTimetable = renderAll;

})();