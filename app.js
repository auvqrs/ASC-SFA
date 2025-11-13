// Advanced Timetable App — see README for features.

(function(){

  // You may wish to put the default branch here to support hot reload.
  // State
  const state = {
    days: ['Mon','Tue','Wed','Thu','Fri'],
    periods: 5, // P1..P5 (slot 0 is Form)
    years: ['Y7','Y8','Y9','Y10','Y11','LSU','SF'],
    subjects: [], // {id,name,color,defaultCount}
    teachers: [], // {id, name, code, working: {Mon:[0,1,2,...]}, subjects: []}
    rooms: [], // {id,name}
    lessons: [], // {id, year, subjectId, teacherId, roomId, count, mergedYears: []}
    grid: [], // [year][day][period]
    pickedLessonId: null,
    filter: { type: 'all', value: '' }
  };

  // DOM shortcuts
  const $ = (s,r=document)=>r.querySelector(s);
  const $$ = (s,r=document)=>Array.from(r.querySelectorAll(s));

  // Elements (add new ones as needed for merged lessons, export, etc.)
  const el = {
    selectDays: $('#selectDays'),
    inputPeriods: $('#inputPeriods'),
    btnApplySettings: $('#btnApplySettings'),
    btnResetAll: $('#btnResetAll'),
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
    lessonMergeYears: $('#lessonMergeYears'),
    tokens: $('#tokens'),
    timetable: $('#timetable'),
    btnAutoSchedule: $('#btnAutoSchedule'),
    btnClear: $('#btnClear'),
    btnExportJSON: $('#btnExportJSON'),
    btnImportJSON: $('#btnImportJSON'),
    fileImport: $('#fileImport'),
    btnExportExcel: $('#btnExportExcel'),
    btnPrint: $('#btnPrint'),
    filterType: $('#filterType'),
    filterValue: $('#filterValue'),
    btnApplyFilter: $('#btnApplyFilter'),
    btnClearFilter: $('#btnClearFilter')
  };

  const uid = (p='id') => p + '_' + Math.random().toString(36).slice(2,9);

  // Subjects — include Form (period 0)
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
    try {
      localStorage.setItem('school_timetable_v2', JSON.stringify(state));
    } catch(e) { /* ignore */ }
  }
  function load() {
    try {
      const raw = localStorage.getItem('school_timetable_v2');
      if(raw){
        const d = JSON.parse(raw);
        Object.assign(state, {days: d.days, periods: d.periods, years: d.years, subjects: d.subjects, teachers: d.teachers, rooms: d.rooms, lessons: d.lessons, grid: d.grid, filter: d.filter});
        normalizeGrid();
        return;
      }
    }catch(e){}
    // Defaults
    state.subjects = defaultSubjects.map(s=>({ id: uid('sub'), name: s[0], color: randomColor(), defaultCount: s[1] }));
    const rooms = [];
    for(let i=101;i<=116;i++) rooms.push('G-'+i);
    for(let i=117;i<=126;i++) rooms.push('F-'+i);
    for(let i=127;i<=138;i++) rooms.push('S-'+i);
    state.rooms = rooms.map(r=>({id:uid('room'), name:r}));
    state.teachers = [];
    state.lessons = [];
    state.grid = createEmptyGrid();
    state.filter = {type:'all',value:''};
    save();
  }

  function randomColor(){
    const hue = Math.floor(Math.random()*360);
    return `hsl(${hue} 70% 60%)`;
  }

  function createEmptyGrid(){
    const slots = 1 + state.periods;
    const grid = [];
    for(let y=0;y<state.years.length;y++){
      const yearRow = [];
      for(let d=0;d<state.days.length;d++){
        yearRow.push(new Array(slots).fill(null));
      }
      grid.push(yearRow);
    }
    return grid;
  }
  function normalizeGrid(){
    const slots = 1 + state.periods;
    if(!Array.isArray(state.grid)) state.grid = createEmptyGrid();
    while(state.grid.length < state.years.length) state.grid.push([]);
    if(state.grid.length > state.years.length) state.grid.splice(state.years.length);
    for(let y=0;y<state.years.length;y++){
      if(!Array.isArray(state.grid[y])) state.grid[y] = [];
      while(state.grid[y].length < state.days.length) state.grid[y].push(new Array(slots).fill(null));
      if(state.grid[y].length > state.days.length) state.grid[y].splice(state.days.length);
      for(let d=0;d<state.days.length;d++){
        if(!Array.isArray(state.grid[y][d])) state.grid[y][d] = new Array(slots).fill(null);
        if(state.grid[y][d].length < slots){
          while(state.grid[y][d].length < slots) state.grid[y][d].push(null);
        }else if(state.grid[y][d].length > slots){
          state.grid[y][d].splice(slots);
        }
      }
    }
  }

  // Rendering
  function renderAll(){
    try{
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
    }catch(err){
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
      const div = document.createElement('div'); div.className='list-item';
      // Per-week editable
      div.innerHTML = `<div><strong style="color:${s.color}">${s.name}</strong>
      <input type="number" min="0" style="width:44px" value="${s.defaultCount}" data-id="${s.id}">
      <span class="badge">/wk</span></div>`;
      container.appendChild(div);
    });
    $$('input[type=number]',container).forEach(inp=>{
      inp.addEventListener('change', e=>{
        const sub = state.subjects.find(s=>s.id===inp.dataset.id);
        if(sub) sub.defaultCount = Math.max(0,parseInt(inp.value)||0);
        renderAll();
      });
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
      left.innerHTML = `<strong>${t.name}</strong> <span class="teacher-badge">${t.code}</span>
      <div class="small">Working: ${formatWorkingPeriods(t)}</div>
      <div class="small">Subjects: ${(t.subjects||[]).map(id=>state.subjects.find(s=>s.id===id)?.name).filter(Boolean).join(", ")||'–'}</div>`;
      const right = document.createElement('div');
      right.innerHTML = `<button data-id="${t.id}" class="edit-teacher">Edit</button> <button data-id="${t.id}" class="remove-teacher">Remove</button>`;
      d.appendChild(left); d.appendChild(right);
      c.appendChild(d);
    });
    $$('.edit-teacher', c).forEach(btn=> btn.addEventListener('click', ()=> editTeacherModal(btn.dataset.id)));
    $$('.remove-teacher', c).forEach(btn=> btn.addEventListener('click', async ()=>{
      const id = btn.dataset.id;
      if(await confirmDialog('Remove teacher and their lessons?')){
        state.lessons = state.lessons.filter(l=>l.teacherId!==id);
        for(let y=0;y<state.grid.length;y++)
          for(let d=0;d<state.grid[y].length;d++)
            for(let s=0;s<state.grid[y][d].length;s++)
              if(state.grid[y][d][s] && state.grid[y][d][s].teacherId === id) state.grid[y][d][s] = null;
        state.teachers = state.teachers.filter(t=>t.id!==id);
        renderAll();
      }
    }));
  }

  function formatWorkingPeriods(t){
    return state.days.map((d,di)=>{
      if(!t.working || !t.working[d]) return '';
      return `${d}:` + t.working[d].join(',');
    }).filter(Boolean).join('; ');
  }

  function renderLessonSelectors(){
    el.lessonYear.innerHTML='';
    el.lessonMergeYears.innerHTML='';
    state.years.forEach(y=>{
      const o=document.createElement('option'); o.value=y; o.textContent=y;
      el.lessonYear.appendChild(o);
      const om=document.createElement('option'); om.value=y; om.textContent=y;
      el.lessonMergeYears.appendChild(om);
    });
    el.lessonSubject.innerHTML='';
    state.subjects.forEach(s=>{
      const o=document.createElement('option'); o.value=s.id; o.textContent=`${s.name} (${s.defaultCount}/wk)`; el.lessonSubject.appendChild(o);
    });
    el.lessonTeacher.innerHTML=''; const blank=document.createElement('option'); blank.value=''; blank.textContent='(none)'; el.lessonTeacher.appendChild(blank);
    state.teachers.forEach(t=>{
      const o=document.createElement('option'); o.value=t.id; o.textContent=`${t.name} • ${t.code}`; el.lessonTeacher.appendChild(o);
    });
    el.lessonRoom.innerHTML=''; const none=document.createElement('option'); none.value=''; none.textContent='(none)'; el.lessonRoom.appendChild(none);
    state.rooms.forEach(r=>{
      const o=document.createElement('option'); o.value=r.id; o.textContent=r.name; el.lessonRoom.appendChild(o);
    });
  }

  // Tokens — unassigned lessons
  function renderTokens(){
    const c = el.tokens; c.innerHTML='';
    const counts = {};
    state.lessons.forEach(l=>counts[l.id]=(counts[l.id]||0)+l.count);
    for(let y=0;y<state.grid.length;y++)
      for(let d=0;d<state.grid[y].length;d++)
        for(let s=0;s<state.grid[y][d].length;s++)
          if(state.grid[y][d][s]?.lessonId) counts[state.grid[y][d][s].lessonId]=Math.max(0,(counts[state.grid[y][d][s].lessonId]||0)-1);

    state.lessons.forEach(l=>{
      if(l.mergedYears && l.mergedYears.length>1) return; // merged lessons are shown only once
      const remaining = counts[l.id]||0;
      if(remaining <= 0) return;
      const sub = state.subjects.find(s=>s.id===l.subjectId);
      const t = state.teachers.find(t=>t.id===l.teacherId);
      const r = state.rooms.find(r=>r.id===l.roomId);
      for(let i=0;i<remaining;i++){
        const div = document.createElement('div'); div.className='token'; div.draggable=true;
        div.dataset.lessonId = l.id;
        div.style.background = sub ? sub.color : '#777';
        div.innerHTML = `<div style="display:flex;flex-direction:column"><div style="font-weight:700">${sub?sub.name:'?'}</div>
          <div style="font-size:12px">${l.year} ${t?('• '+t.code):''}${r?(' • '+r.name):''}</div></div>`;
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

  // Timetable grid (per year, per day, per period)
  function renderGrid(){
    el.timetable.innerHTML='';
    const slots=1+state.periods,days=state.days,years=state.years;
    const table = document.createElement('table'); table.className='timetable-table';
    // Headings
    const thead=document.createElement('thead'),trTop=document.createElement('tr');
    trTop.appendChild(document.createElement('th')).textContent='Year / Day';
    days.forEach(day=>{
      const th=document.createElement('th'); th.colSpan=slots; th.textContent=day; trTop.appendChild(th);
    }); thead.appendChild(trTop);
    const trSlots=document.createElement('tr'); trSlots.appendChild(document.createElement('th')).textContent='';
    days.forEach(()=>{const thForm=document.createElement('th'); thForm.className='slot-header'; thForm.textContent='Form'; trSlots.appendChild(thForm);
      for(let p=1;p<slots;p++){const thP=document.createElement('th'); thP.className='slot-header'; thP.textContent='P'+p; trSlots.appendChild(thP);}});
    thead.appendChild(trSlots); table.appendChild(thead);

    // Body per year
    const tbody=document.createElement('tbody');
    for(let y=0;y<years.length;y++){
      const tr=document.createElement('tr');
      const thYear=document.createElement('th'); thYear.className='year-label'; thYear.textContent=years[y]; tr.appendChild(thYear);

      for(let d=0;d<days.length;d++){
        for(let s=0;s<slots;s++){
          const td=document.createElement('td'); td.className='cell';
          td.dataset.yearIndex=y; td.dataset.dayIndex=d; td.dataset.slotIndex=s;
          td.addEventListener('dragover', cellDragOver);
          td.addEventListener('dragleave', cellDragLeave);
          td.addEventListener('drop', cellDrop);
          td.addEventListener('click', cellClick);

          const assigned = state.grid[y][d][s];
          if(assigned){
            // Merged lessons label
            const sub=state.subjects.find(x=>x.id===assigned.subjectId);
            const t=state.teachers.find(x=>x.id===assigned.teacherId);
            const r=state.rooms.find(x=>x.id===assigned.roomId);
            let yearsLabel=assigned.mergedYears? `Merged: ${assigned.mergedYears.join(", ")}<br>` : '';
            const wrapper=document.createElement('div');
            wrapper.innerHTML = `<div class="assign" style="background:${sub?sub.color:'#999'};color:white;padding:4px;border-radius:4px">
            ${yearsLabel}${sub?sub.name:'(Form)'}
            </div>
            <div class="meta">${t?('<span class="teacher-badge">'+t.code+'</span>'):''} ${r?(' • '+r.name):''}</div>`;
            td.appendChild(wrapper);

            const btn=document.createElement('button');
            btn.textContent='✕'; btn.title='Remove assignment';
            btn.style.position='absolute'; btn.style.right='6px'; btn.style.top='6px';
            btn.style.border='none'; btn.style.background='rgba(0,0,0,0.06)'; btn.style.borderRadius='4px'; btn.style.cursor='pointer';
            btn.addEventListener('click',ev=>{ev.stopPropagation(); state.grid[y][d][s]=null; renderAll();});
            td.appendChild(btn);
          }else{
            td.innerHTML='<div class="small">empty</div>';
          }

          // filter dim
          if(!cellMatchesFilter(assigned,y)) td.classList.add('dimmed');
          else td.classList.remove('dimmed');

          // teacher working day/period highlight
          if(assigned && assigned.teacherId){
            const t=state.teachers.find(tt=>tt.id===assigned.teacherId);
            const dayName=state.days[d];
            if(!t.working || !t.working[dayName] || !t.working[dayName].includes(s)){
              td.style.boxShadow='inset 0 0 0 1000px rgba(255,80,80,0.08)';
            }else td.style.boxShadow='';
          }

          tr.appendChild(td);
        }
      }
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    el.timetable.appendChild(table);
  }

  // Filtering
  function renderFilterOptions(){
    el.filterType.value = state.filter.type;
    el.filterValue.innerHTML='';
    if(state.filter.type==='all'||state.filter.type==='staff'){el.filterValue.style.display='none'; return;}
    el.filterValue.style.display='inline-block';
    if(state.filter.type==='year'){state.years.forEach(y=>{const o=document.createElement('option'); o.value=y;o.textContent=y;el.filterValue.appendChild(o);});}
    else if(state.filter.type==='teacher'){const blank=document.createElement('option');blank.value='';blank.textContent='(choose)';el.filterValue.appendChild(blank);state.teachers.forEach(t=>{const o=document.createElement('option');o.value=t.id;o.textContent=`${t.name} (${t.code})`;el.filterValue.appendChild(o);});}
    else if(state.filter.type==='room'){const blank=document.createElement('option');blank.value='';blank.textContent='(choose)';el.filterValue.appendChild(blank);state.rooms.forEach(r=>{const o=document.createElement('option');o.value=r.id;o.textContent=r.name;el.filterValue.appendChild(o);});}
    else if(state.filter.type==='subject'){const blank=document.createElement('option');blank.value='';blank.textContent='(choose)';el.filterValue.appendChild(blank);state.subjects.forEach(s=>{const o=document.createElement('option');o.value=s.id;o.textContent=s.name;el.filterValue.appendChild(o);});}
    if(state.filter.value) el.filterValue.value = state.filter.value;
  }
  function cellMatchesFilter(assigned, yearIndex){
    const f=state.filter;
    if(!f||f.type==='all') return true;
    if(f.type==='staff') return assigned && !!assigned.teacherId;
    if(f.type==='year'){const val=f.value;if(!val) return true; return state.years[yearIndex]===val;}
    if(!assigned) return false;
    if(f.type==='teacher') return assigned.teacherId === f.value;
    if(f.type==='room') return assigned.roomId === f.value;
    if(f.type==='subject') return assigned.subjectId === f.value;
    return true;
  }

  // Drag/drop
  let draggingLessonId=null;
  function tokenDragStart(e){
    draggingLessonId=e.currentTarget.dataset.lessonId;
    e.dataTransfer.setData('text/plain',draggingLessonId);
  }
  function cellDragOver(e){ e.preventDefault(); e.currentTarget.classList.add('drag-over'); }
  function cellDragLeave(e){ e.currentTarget.classList.remove('drag-over'); }

  async function cellDrop(e){
    e.preventDefault(); e.currentTarget.classList.remove('drag-over');
    const lessonId = e.dataTransfer.getData('text/plain')||draggingLessonId||state.pickedLessonId;
    if(!lessonId) return;
    const y=+e.currentTarget.dataset.yearIndex;
    const d=+e.currentTarget.dataset.dayIndex;
    const s=+e.currentTarget.dataset.slotIndex;
    const lesson=state.lessons.find(l=>l.id===lessonId); if(!lesson) return;
    if(!lesson.mergedYears && lesson.year !== state.years[y]){await alertDialog(`This lesson is for ${lesson.year}. Place in the matching row.`); return;}
    await assignLessonToCell(lessonId,y,d,s);
    draggingLessonId=null; state.pickedLessonId=null; renderAll();
  }

  async function assignLessonToCell(lessonId, yearIndex, dayIndex, slotIndex){
    const lesson = state.lessons.find(l=>l.id===lessonId);
    if(!lesson) return;
    // For merged lessons
    const years = lesson.mergedYears && lesson.mergedYears.length>0 ? lesson.mergedYears : [lesson.year];
    for(const ystr of years){
      const yidx = state.years.indexOf(ystr);
      if(yidx===-1) continue;
      // Prevent teacher/room/lesson double booking
      if(isSlotConflict(lesson,yidx,dayIndex,slotIndex)){
        await alertDialog("Slot is not available (teacher, room or double-book).");
        return;
      }
      if(lesson.teacherId){
        const teacher = state.teachers.find(t=>t.id===lesson.teacherId);
        const dayName = state.days[dayIndex];
        if(!teacher.working || !teacher.working[dayName] || !teacher.working[dayName].includes(slotIndex)){
          const ok = await confirmDialog(`Teacher ${teacher.name} does not work ${dayName} P${slotIndex}. Place anyway?`);
          if(!ok) return;
        }
      }
      state.grid[yidx][dayIndex][slotIndex]={ lessonId: lesson.id, subjectId: lesson.subjectId, teacherId: lesson.teacherId, roomId: lesson.roomId, year: ystr, mergedYears: lesson.mergedYears, placedId: uid('placed') };
    }
  }

  function isSlotConflict(lesson, yearIdx, dayIdx, slotIdx){
    // Room conflict — room cannot overlap
    for(let y=0;y<state.grid.length;y++)
      if(state.grid[y][dayIdx][slotIdx] && lesson.roomId && state.grid[y][dayIdx][slotIdx].roomId === lesson.roomId) return true;
    // Teacher conflict — teacher cannot teach 2 at once
    for(let y=0;y<state.grid.length;y++)
      if(state.grid[y][dayIdx][slotIdx] && lesson.teacherId && state.grid[y][dayIdx][slotIdx].teacherId === lesson.teacherId) return true;
    // Teacher per year more than 1 period/day (boring repeats), discourage
    let teacherPeriodCount=0;
    for(let s=0;s<state.grid[yearIdx][dayIdx].length;s++)
      if(state.grid[yearIdx][dayIdx][s] && lesson.teacherId && state.grid[yearIdx][dayIdx][s].teacherId === lesson.teacherId) teacherPeriodCount++;
    if(teacherPeriodCount > 0) return true;
    return false;
  }

  async function cellClick(e){
    const y=+e.currentTarget.dataset.yearIndex;
    const d=+e.currentTarget.dataset.dayIndex;
    const s=+e.currentTarget.dataset.slotIndex;
    if(state.pickedLessonId){
      const lesson=state.lessons.find(l=>l.id===state.pickedLessonId);
      if(!lesson.mergedYears && lesson.year!==state.years[y]){
        await alertDialog(`Picked lesson is for ${lesson.year}. Click a ${lesson.year} row.`); return;
      }
      await assignLessonToCell(state.pickedLessonId,y,d,s);
      state.pickedLessonId=null; renderAll(); return;
    }
    quickAssignPopup(e.currentTarget,y,d,s);
  }

  function quickAssignPopup(cellEl, y,d,s){
    const remainingCounts={};
    state.lessons.forEach(l=>remainingCounts[l.id]=(remainingCounts[l.id]||0)+l.count);
    for(let y=0;y<state.grid.length;y++)
      for(let d=0;d<state.grid[y].length;d++)
        for(let si=0;si<state.grid[y][d].length;si++)
          if(state.grid[y][d][si]?.lessonId) remainingCounts[state.grid[y][d][si].lessonId]=Math.max(0,(remainingCounts[state.grid[y][d][si].lessonId]||0)-1);

    const available = state.lessons.filter(l=>l.mergedYears ? l.mergedYears.includes(state.years[y]) : l.year === state.years[y] && (remainingCounts[l.id]||0)>0 );
    const menu=document.createElement('div');
    menu.style.position='absolute';menu.style.left='10px';menu.style.top='10px';menu.style.zIndex=1000;menu.style.background='white';menu.style.border='1px solid #e2e8f0';menu.style.borderRadius='6px';menu.style.padding='8px';menu.style.boxShadow='0 8px 24px rgba(0,0,0,0.08)';
    const title=document.createElement('div');title.textContent=`Assign for ${state.years[y]} (${state.days[d]} P${s})`;title.style.fontWeight='700';title.style.marginBottom='6px';menu.appendChild(title);

    if(available.length===0){
      const none=document.createElement('div');none.textContent='No assignable lessons for this year.';menu.appendChild(none);
    } else{
      available.slice(0,300).forEach(l=>{
        const sub=state.subjects.find(x=>x.id===l.subjectId);
        const t=state.teachers.find(x=>x.id===l.teacherId);
        const row=document.createElement('div');row.style.marginBottom='6px';
        const btn=document.createElement('button');btn.textContent=`${l.year} • ${sub?sub.name:'?'} ${t?(' • '+t.code):''}`;btn.style.width='100%';
        btn.addEventListener('click', async ()=>{await assignLessonToCell(l.id,y,d,s);document.body.removeChild(menu);renderAll();});
        row.appendChild(btn);menu.appendChild(row);
      });
    }
    const cancel=document.createElement('button');cancel.textContent='Close';cancel.style.marginTop='6px';cancel.addEventListener('click',()=>document.body.removeChild(menu));menu.appendChild(cancel);
    document.body.appendChild(menu);
    const rect=cellEl.getBoundingClientRect();
    menu.style.left=Math.min(window.innerWidth-320,rect.left+12)+'px';
    menu.style.top=Math.min(window.innerHeight-220,rect.top+12)+'px';
  }

  // Auto-scheduler
  async function autoSchedule(){
    if(!state.lessons.length){await alertDialog('No lessons to schedule. Create lessons first.');return;}
    if(!await confirmDialog('Auto-schedule will clear existing assignments. Continue?')) return;
    state.grid=createEmptyGrid();normalizeGrid();
    let tokens=[];
    state.lessons.forEach(l=>{
      for(let i=0;i<l.count;i++) tokens.push(Object.assign({tokenUid:uid('tk')},l));
    });
    // Sort: merged lessons first, then by least #slots possible
    tokens.sort((a,b)=>(b.mergedYears?.length||0)-(a.mergedYears?.length||0));

    let failed=[];
    for(const tok of tokens){
      let placed=false;
      let years=tok.mergedYears&&tok.mergedYears.length>0 ? tok.mergedYears : [tok.year];
      outer: for(let d=0;d<state.days.length;d++)
        for(let s=0;s<1+state.periods;s++)
        {
          for(const ystr of years){
            const yidx=state.years.indexOf(ystr);
            if(yidx===-1) continue;
            if(isSlotConflict(tok,yidx,d,s)) continue;
            const teacher = state.teachers.find(t=>t.id===tok.teacherId);
            if(teacher && (!teacher.working || !teacher.working[state.days[d]] || !teacher.working[state.days[d]].includes(s))) continue;
          }
          // No conflicts — place merged group
          for(const ystr of years)
            state.grid[state.years.indexOf(ystr)][d][s]={
              lessonId: tok.id, subjectId: tok.subjectId, teacherId: tok.teacherId, roomId: tok.roomId,
              year: ystr, placedId: tok.tokenUid, mergedYears: tok.mergedYears
            };
          placed=true; break outer;
        }
      if(!placed) failed.push(tok);
    }

    // If any failed, show summary
    if(failed.length){
      await alertDialog(`${failed.length} lessons couldn't be scheduled. Likely conflicts or lack of available period. Try adjusting teachers, rooms, period availability or reduce lessons/week.`);
    }
    renderAll();
  }

  // Export Excel (with xlsx.min.js)
  function exportExcel(){
    const slots = 1 + state.periods, days = state.days;
    let rows = [];
    for(let y=0;y<state.years.length;y++)
      for(let d=0;d<state.days.length;d++)
        for(let s=0;s<slots;s++){
          const cell=state.grid[y][d][s];
          if(cell){
            rows.push({
              lessonid: cell.lessonId,
              teacherid: cell.teacherId,
              teacher: state.teachers.find(t=>t.id===cell.teacherId)?.name || '',
              subjectid: cell.subjectId,
              subject: state.subjects.find(su=>su.id===cell.subjectId)?.name || '',
              room: state.rooms.find(r=>r.id===cell.roomId)?.name || '',
              year: cell.year,
              day: days[d],
              period: s===0?'Form':'P'+s,
              mergedYears: cell.mergedYears?.join(', ') || '',
            });
          }
        }
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Timetable");
    XLSX.writeFile(wb,"timetable.xlsx");
  }

  // Import/export JSON
  function exportJSON(){
    const data = {...state};
    const blob = new Blob([JSON.stringify(data,null,2)],{type:'application/json'});
    const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download='timetable.json';document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);
  }
  function importJSONFile(file){
    const reader=new FileReader();
    reader.onload = async (ev)=>{
      try{
        const d=JSON.parse(ev.target.result);
        Object.assign(state,d);normalizeGrid();renderAll();
      }catch(e){ await alertDialog('Invalid JSON');}
    };
    reader.readAsText(file);
  }

  // Modals
  function alertDialog(msg,title='Notice'){
    return new Promise(resolve=>{
      const m=createModalContainer(),box=createModalBox(title);
      const div=document.createElement('div');div.style.margin='8px 0 12px';div.textContent=msg;box.appendChild(div);
      const btn=document.createElement('button');btn.textContent='OK';stylePrimaryButton(btn);btn.addEventListener('click',()=>{document.body.removeChild(m);resolve();});box.appendChild(btn);
      m.appendChild(box);document.body.appendChild(m);btn.focus();
    });
  }
  function confirmDialog(msg,title='Confirm'){
    return new Promise(resolve=>{
      const m=createModalContainer(),box=createModalBox(title);
      const div=document.createElement('div');div.style.margin='8px 0 12px';div.textContent=msg;box.appendChild(div);
      const row=document.createElement('div');row.style.display='flex';row.style.justifyContent='flex-end';row.style.gap='8px';
      const cancel=document.createElement('button');cancel.textContent='Cancel';cancel.style.padding='8px 12px';cancel.style.borderRadius='6px';cancel.style.border='1px solid #e0e6ef';cancel.addEventListener('click',()=>{document.body.removeChild(m);resolve(false);});
      const ok=document.createElement('button');ok.textContent='OK';stylePrimaryButton(ok);ok.addEventListener('click',()=>{document.body.removeChild(m);resolve(true);});
      row.appendChild(cancel);row.appendChild(ok);box.appendChild(row);
      m.appendChild(box);document.body.appendChild(m);ok.focus();
    });
  }
  function createModalContainer(){const m=document.createElement('div');m.style.position='fixed';m.style.left='0';m.style.top='0';m.style.right='0';m.style.bottom='0';m.style.background='rgba(0,0,0,0.25)';m.style.zIndex=9999;return m;}
  function createModalBox(title){const box=document.createElement('div');box.style.width='440px';box.style.margin='70px auto';box.style.background='#fff';box.style.padding='16px';box.style.borderRadius='10px';box.style.boxShadow='0 10px 40px rgba(10,20,40,0.2)';const h=document.createElement('h3');h.textContent=title;h.style.margin='0 0 8px';box.appendChild(h);return box;}
  function stylePrimaryButton(btn){btn.style.padding='8px 12px';btn.style.borderRadius='6px';btn.style.border='none';btn.style.background='#1976d2';btn.style.color='#fff';btn.style.cursor='pointer';}

  // Teacher code generator
  function generateTeacherCode(fullName,existingCodes){
    const parts=fullName.trim().split(/\s+/);
    if(parts.length===0) return uid('T');
    const first=parts[0],surname=parts.length>1?parts[parts.length-1]:first;
    const firstInitial=first[0].toUpperCase(),lastLetter=surname.slice(-1).toUpperCase();
    for(let n=0;n<Math.max(3,surname.length);n++){
      const nth=(n<surname.length)?surname[n].toUpperCase():surname[0].toUpperCase();
      const code=`${firstInitial}${nth}${lastLetter}`;
      if(!existingCodes.includes(code)) return code;
    }
    let i=1;while(true){
      const code=`${firstInitial}${surname[0].toUpperCase()}${lastLetter}${i}`;
      if(!existingCodes.includes(code)) return code;i++;
    }
  }

  // Teacher edit modal (Advanced)
  function editTeacherModal(id){
    const teacher=state.teachers.find(t=>t.id===id);if(!teacher) return;
    const modal=document.createElement('div');modal.style.position='fixed';modal.style.left='0';modal.style.top='0';modal.style.right='0';modal.style.bottom='0';modal.style.background='rgba(0,0,0,0.22)';
    const box=document.createElement('div');box.style.width='440px';box.style.margin='80px auto';box.style.background='white';box.style.padding='16px';box.style.borderRadius='10px';
    box.innerHTML=`<h3>Edit ${teacher.name} • ${teacher.code}</h3>`;
    // Name/code
    const nameDiv=document.createElement('div'),nameInput=document.createElement('input');nameInput.value=teacher.name;nameDiv.appendChild(document.createTextNode('Name: '));nameDiv.appendChild(nameInput);
    const codeInput = document.createElement('input');codeInput.value=teacher.code;nameDiv.appendChild(document.createTextNode(' Code: '));nameDiv.appendChild(codeInput);
    box.appendChild(nameDiv);
    // Subjects
    const subjDiv=document.createElement('div');subjDiv.textContent='Subjects: ';
    const sel=document.createElement('select');sel.multiple=true;
    state.subjects.forEach(s=>{const o=document.createElement('option');o.value=s.id;o.textContent=s.name;if(teacher.subjects&&teacher.subjects.includes(s.id)) o.selected=true;sel.appendChild(o);});
    subjDiv.appendChild(sel); box.appendChild(subjDiv);

    // Working per day/period
    box.appendChild(document.createElement('hr'));
    box.appendChild(document.createTextNode('Teaching days/periods:'));
    const table=document.createElement('table');table.style.width='100%';
    for(const d of state.days){
      const tr=document.createElement('tr');
      const tdDay=document.createElement('td');tdDay.textContent=d;tr.appendChild(tdDay);
      const tdPeriods=document.createElement('td');
      for(let p=0;p<=state.periods;p++){
        const cb=document.createElement('input');cb.type='checkbox';cb.value=p;cb.checked=teacher.working&&teacher.working[d]&&teacher.working[d].includes(p);cb.dataset.day=d;cb.dataset.period=p;
        tdPeriods.appendChild(cb);tdPeriods.appendChild(document.createTextNode(' '+(p===0?'Form':'P'+p)+' '));
      }
      tr.appendChild(tdPeriods);table.appendChild(tr);
    }
    box.appendChild(table);
    // Save/cancel
    const saveBtn=document.createElement('button');saveBtn.textContent='Save';saveBtn.style.marginRight='8px';
    const cancelBtn=document.createElement('button');cancelBtn.textContent='Cancel';
    saveBtn.addEventListener('click',()=>{
      teacher.name=nameInput.value.trim()||teacher.name;
      teacher.code=codeInput.value.trim()||teacher.code;
      teacher.subjects=[];
      Array.from(sel.selectedOptions).forEach(o=>teacher.subjects.push(o.value));
      teacher.working={};for(const d of state.days) teacher.working[d]=[];
      table.querySelectorAll('input[type=checkbox]').forEach(cb=>{if(cb.checked) teacher.working[cb.dataset.day].push(+cb.dataset.period);});
      document.body.removeChild(modal); renderAll();
    });
    cancelBtn.addEventListener('click',()=>document.body.removeChild(modal));
    box.appendChild(saveBtn);box.appendChild(cancelBtn);
    modal.appendChild(box);document.body.appendChild(modal);
  }

  // UI events
  function wireEvents(){
    el.btnApplySettings.addEventListener('click',async ()=>{
      const days=Array.from(el.selectDays.selectedOptions).map(o=>o.value);
      if(days.length<1){await alertDialog('Pick at least one day');return;}
      const p=Math.max(1,Math.min(12,parseInt(el.inputPeriods.value)||5));
      state.days=days;state.periods=p;state.grid=createEmptyGrid();normalizeGrid();renderAll();
    });
    el.btnResetAll.addEventListener('click',async ()=>{
      if(await confirmDialog('Reset ALL data? This will delete subjects, teachers, lessons, assignments.')){
        localStorage.removeItem('school_timetable_v2'); load(); renderAll();
      }
    });
    el.btnAddTeacher.addEventListener('click',async ()=>{
      const nm=el.teacherFullName.value.trim();if(!nm) return;
      const code=generateTeacherCode(nm,state.teachers.map(t=>t.code));
      // By default allow all periods on all chosen days
      const wd={};state.days.forEach(d=>wd[d]=Array.from({length:1+state.periods},(_,i)=>i));
      state.teachers.push({id:uid('t'),name:nm,code,working:wd,subjects:[]});
      el.teacherFullName.value='';renderAll();
    });
    el.btnGenerateLessonsForYears.addEventListener('click',async ()=>{
      let selectedYears = Array.from(el.lessonYear.selectedOptions).map(o=>o.value);
      if(selectedYears.length===0){
        const nodes=Array.from(el.yearsList.children).filter(n=>n.classList.contains('active')); selectedYears=nodes.map(n=>n.firstChild.textContent);
      }
      if(selectedYears.length===0){await alertDialog('Select years (in Create Lesson or click years) first.');return;}
      selectedYears.forEach(y=>{
        state.subjects.forEach(s=>{
          const exist=state.lessons.find(l=>l.year===y&&l.subjectId===s.id);
          if(!exist){
            state.lessons.push({id:uid('les'),year:y,subjectId:s.id,teacherId:null,roomId:null,count:s.defaultCount,mergedYears:[]});
          }
        });
      });
      renderAll();
    });
    el.btnAddLesson.addEventListener('click',async ()=>{
      const selectedYears=Array.from(el.lessonYear.selectedOptions).map(o=>o.value);
      const mergeYears=Array.from(el.lessonMergeYears.selectedOptions).map(o=>o.value);
      if(selectedYears.length===0 && mergeYears.length===0){ await alertDialog('Select years to add.'); return; }
      const subjectId=el.lessonSubject.value,teacherId=el.lessonTeacher.value||null,roomId=el.lessonRoom.value||null,count=Math.max(1,Math.min(50,parseInt(el.lessonCount.value)||1);
      if(mergeYears.length>1){ // merged lesson (not per year)
        const exists=state.lessons.find(l=>l.subjectId===subjectId && l.teacherId===teacherId && l.roomId===roomId && JSON.stringify(l.mergedYears)===JSON.stringify(mergeYears));
        if(!exists)
          state.lessons.push({id:uid('les'),year:mergeYears[0],subjectId,teacherId,roomId,count,mergedYears:mergeYears});
      }else{
        selectedYears.forEach(year=>{
          const exists=state.lessons.find(l=>l.year===year && l.subjectId===subjectId && l.teacherId===teacherId && l.roomId===roomId);
          if(!exists)
            state.lessons.push({id:uid('les'),year,subjectId,teacherId,roomId,count,mergedYears:[]});
        });
      }
      renderAll();
    });
    el.btnAutoSchedule.addEventListener('click',autoSchedule);
    el.btnClear.addEventListener('click',async ()=>{if(await confirmDialog('Clear all assignments?')){state.grid=createEmptyGrid();normalizeGrid();renderAll();}});
    el.btnExportJSON.addEventListener('click',exportJSON);
    el.btnImportJSON.addEventListener('click',()=>el.fileImport.click());
    el.fileImport.addEventListener('change',(ev)=>{const f=ev.target.files[0];if(f) importJSONFile(f);ev.target.value='';});
    el.btnExportExcel.addEventListener('click',exportExcel);
    el.btnPrint.addEventListener('click',()=>window.print());
    el.filterType.addEventListener('change',(ev)=>{state.filter.type=ev.target.value;state.filter.value='';renderFilterOptions();});
    el.btnApplyFilter.addEventListener('click',()=>{state.filter.type=el.filterType.value;state.filter.value=el.filterValue.value||'';renderGrid();});
    el.btnClearFilter.addEventListener('click',()=>{state.filter={type:'all',value:''};renderFilterOptions();renderGrid();});
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
