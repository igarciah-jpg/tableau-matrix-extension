(function(){
  'use strict';

  // ===================== Utils =====================
  const $ = sel => document.querySelector(sel);

  const fmtDisplay = c =>
    (c && typeof c.formattedValue !== 'undefined') ? String(c.formattedValue ?? '')
    : (c && typeof c.value !== 'undefined') ? String(c.value ?? '')
    : String(c ?? '');

  const toNumber = v => {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    const n = parseFloat(String(v).replace(/[^\d.\-]/g,''));
    return isNaN(n) ? 0 : n;
  };

  const norm = s => (s||'')
    .toLowerCase()
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu,'')
    .trim();

  const aggWords = ['sum','suma','avg','average','prom','min','max','count','conteo','cnt','median','var','stdev','std','agg'];

  const cleanMeasureName = name => {
    if (!name) return '';
    const m = name.match(/\(([^)]+)\)/);
    return m ? m[1] : name;
  };

  // ===================== Estado =====================
  const state = {
    loading: true,
    columns: [],
    rows: [],
    dims: [],
    meas: [],
    dimOrder: [],
    expanded: {},                 // SOLO expansión manual (y ahora también programática cuando filtras)
    sort: { index: -1, dir: 'asc' },
    ts: null,

    pageSize: 10,
    currentPage: 1,
    showGrandTotal: true,
    search: '',

    pageOptions: [10,20,50],
    totalLevel1: 0
  };

  let worksheet = null;
  const settings = () => tableau.extensions.settings;

  // ===== Filtro por click (como extensión 1) =====
  let _summaryFieldNames = [];       // nombres “Summary” reales
  let _summaryDimNames   = [];       // dims detectadas en summary
  let _filterDimFieldNames = [];     // dims usadas en jerarquía → nombres Summary (en orden)
  let _lastAppliedPath = null;       // toggle
  let _suspendRefresh = false;       // evita refresh durante applyFilter

  // ✅ NUEVO: para NO perder expanded al refrescar cuando filtras
  let _preserveExpandedOnce = false;

  // ===================== Init =====================
  window.addEventListener('load', async () => {
    await tableau.extensions.initializeAsync();

    // ✅ estilos highlight + icono (sin tocar HTML)
    injectActiveStyles();

    state.pageSize       = Number(settings().get('pageSize')) || 10;
    state.currentPage    = Number(settings().get('currentPage')) || 1;
    state.showGrandTotal = settings().get('showGrandTotal') !== 'false';
    state.search         = settings().get('search') || '';

    worksheet = tableau.extensions.worksheetContent.worksheet;

    worksheet.addEventListener(
      tableau.TableauEventType.SummaryDataChanged,
      () => { if (!_suspendRefresh) refresh(); }
    );

    await refresh();
  });

  function saveSettings(){
    settings().set('pageSize', state.pageSize);
    settings().set('currentPage', state.currentPage);
    settings().set('showGrandTotal', state.showGrandTotal);
    settings().set('search', state.search);
    settings().saveAsync();
  }

  // ===================== Lectura =====================
  async function readSummary(){
    const reader = await worksheet.getSummaryDataReaderAsync();
    const table = await reader.getAllPagesAsync();
    await reader.releaseAsync();
    return table;
  }

  function isAggName(name){
    const low = norm(name);
    if (/\b(sum|avg|count|min|max|median|stdev|std|var)\b/.test(low)) return true;
    for (const w of aggWords) if (low.includes(w)) return true;
    return /\(.+\)/.test(low);
  }

  function detectDimsAndMeasures(cols){
    const dims=[], meas=[];
    cols.forEach((c,i)=> (isAggName(c)?meas:dims).push(i));
    return { dims, meas };
  }

  function findColumnIndex(colNorms, target){
    let p = colNorms.indexOf(target);
    if (p>=0) return p;
    return colNorms.findIndex(n => n.includes(target) || target.includes(n));
  }

  async function getOrdersFromMarks(cols, det){
    const spec = await worksheet.getVisualSpecificationAsync();
    const marks = spec.marksSpecifications[spec.activeMarksSpecificationIndex];
    const norms = cols.map(norm);

    const seenD=new Set(), seenM=new Set();
    const od=[], om=[];

    (marks.encodings||[]).forEach(e=>{
      if (!e.field) return;
      const n = norm(e.field.name);
      const i = findColumnIndex(norms,n);
      if (i<0) return;

      if (det.dims.includes(i) && !seenD.has(i)){ seenD.add(i); od.push(i); }
      if (det.meas.includes(i) && !seenM.has(i)){ seenM.add(i); om.push(i); }
    });

    return {
      dimOrder: od.concat(det.dims.filter(i=>!seenD.has(i))),
      measOrder: om.concat(det.meas.filter(i=>!seenM.has(i)))
    };
  }

  // ===================== Refresh =====================
  async function refresh(){
    const s = await readSummary();
    const cols = s.columns.map(c=>c.fieldName||c.name||'col');
    const rows = s.data.map(r=>r.map(fmtDisplay));
    const det = detectDimsAndMeasures(cols);
    const ord = await getOrdersFromMarks(cols, det);

    // ✅ mapeos Summary (para filtros)
    _summaryFieldNames = cols.slice();
    _summaryDimNames   = det.dims.map(i => cols[i]);

    // ✅ preservar expanded solo cuando vienes de filtrar
    const prevExpanded = state.expanded;

    Object.assign(state,{
      columns: cols,
      rows,
      dims: det.dims,
      meas: ord.measOrder,
      dimOrder: ord.dimOrder,
      expanded: _preserveExpandedOnce ? prevExpanded : {},  // ← cambio clave
      ts: new Date().toLocaleTimeString(),
      loading:false
    });

    _preserveExpandedOnce = false;

    // ✅ construir lista de nombres Summary por nivel (en el orden real dimOrder)
    const sumNorm = _summaryFieldNames.map(norm);
    _filterDimFieldNames = state.dimOrder.map(idx=>{
      const nameU = state.columns[idx];
      const nU = norm(nameU);

      // a) exacto
      let pos = sumNorm.indexOf(nU);
      if (pos>=0) return _summaryFieldNames[pos];

      // b) parcial
      pos = sumNorm.findIndex(n => n.includes(nU) || nU.includes(n));
      if (pos>=0) return _summaryFieldNames[pos];

      // c) fallback dims-summary
      const posDim = _summaryDimNames.map(norm).findIndex(n => n===nU);
      if (posDim>=0) return _summaryDimNames[posDim];

      // d) último recurso
      return nameU;
    });

    render();
  }

  // ===================== Árbol =====================
  function buildTree(){
    const root={key:'',depth:0,children:new Map(),agg:new Array(state.meas.length).fill(0),path:[]};

    state.rows.forEach(r=>{
      let node=root;
      state.dimOrder.forEach((d,i)=>{
        const v=r[d]??'';
        const k=(node.key?node.key+'||':'')+v;
        if(!node.children.has(v)){
          node.children.set(v,{
            key:k,name:v,depth:i+1,children:new Map(),
            agg:new Array(state.meas.length).fill(0),
            path:node.path.concat([v])
          });
        }
        node=node.children.get(v);
        state.meas.forEach((m,mi)=>node.agg[mi]+=toNumber(r[m]));
      });
    });

    root.children.forEach(c=>{
      state.meas.forEach((_,i)=>root.agg[i]+=c.agg[i]);
    });

    return root;
  }

  // ✅ NUEVO: expandir el path clickeado (para que NO colapse al filtrar)
  function expandPath(path){
    let key = '';
    path.forEach(v=>{
      key = key ? (key + '||' + v) : String(v);
      state.expanded[key] = true;
    });
  }

  // ===================== Flatten (tu lógica intacta) =====================
  function flatten(tree){
    const headers=['Jerarquía'].concat(
      state.meas.map(i=>cleanMeasureName(state.columns[i]))
    );
    const out=[];

    if(state.showGrandTotal){
      const row=new Array(headers.length).fill('');
      row[0]={depth:0,key:'::total',name:'Total',path:[], hasChildren:false, isTotal:true};
      state.meas.forEach((_,i)=>row[1+i]=tree.agg[i].toLocaleString());
      out.push({type:'total',row});
    }

    let level1=[...tree.children.values()];
    state.totalLevel1 = level1.length;

    if(state.sort.index>=0){
      level1.sort((a,b)=>{
        if(state.sort.index===0){
          return state.sort.dir==='desc'
            ? b.name.localeCompare(a.name)
            : a.name.localeCompare(b.name);
        }
        return state.sort.dir==='desc'
          ? b.agg[state.sort.index-1]-a.agg[state.sort.index-1]
          : a.agg[state.sort.index-1]-b.agg[state.sort.index-1];
      });
    }

    const start=(state.currentPage-1)*state.pageSize;
    const end=start+state.pageSize;
    const pageNodes=level1.slice(start,end);

    function hasMatch(n){
      if(!state.search) return true;
      if(norm(n.name).includes(norm(state.search))) return true;
      for(const c of n.children.values()){
        if(hasMatch(c)) return true;
      }
      return false;
    }

    function walk(n){
      if(!hasMatch(n)) return;

      const row=new Array(headers.length).fill('');
      row[0]={
        depth:n.depth,
        key:n.key,
        name:n.name,
        path:n.path,
        hasChildren: (n.children && n.children.size>0),
        isTotal:false
      };
      state.meas.forEach((_,i)=>row[1+i]=n.agg[i].toLocaleString());
      out.push({type:'group',row});

      const autoExpand =
        state.search &&
        [...n.children.values()].some(hasMatch);

      if(state.expanded[n.key] || autoExpand){
        [...n.children.values()].forEach(ch=>walk(ch));
      }
    }

    pageNodes.forEach(n=>walk(n));
    return {headers,rows:out};
  }

  // ===================== NUEVO: Filtro por click + toggle =====================
  async function applyPathFilters(path){
    try{
      _suspendRefresh = true;

      const samePath =
        Array.isArray(_lastAppliedPath) &&
        _lastAppliedPath.length === path.length &&
        _lastAppliedPath.every((v,i)=> v === path[i]);

      // 🔁 click en la misma fila => limpiar filtros
      if (samePath){
        for (const fname of _filterDimFieldNames){
          try { await worksheet.clearFilterAsync(fname); } catch(e) {}
        }
        _lastAppliedPath = null;
        return;
      }

      // limpiar antes de aplicar
      for (const fname of _filterDimFieldNames){
        try { await worksheet.clearFilterAsync(fname); } catch(e) {}
      }

      // aplicar por niveles (path)
      for (let i=0;i<path.length;i++){
        await worksheet.applyFilterAsync(
          _filterDimFieldNames[i],
          [path[i]],
          tableau.FilterUpdateType.Replace
        );
      }

      _lastAppliedPath = path.slice();

      // ✅ mantener expansión al filtrar (NO colapsa)
      expandPath(path);
      _preserveExpandedOnce = true;

    } catch (err){
      console.error('applyPathFilters', { fields:_filterDimFieldNames, path }, err);
    } finally {
      _suspendRefresh = false;
      refresh();
    }
  }

  // ===================== Render =====================
  function render(){
    const rootEl=$('#root');
    const activeId = document.activeElement?.id;

    rootEl.innerHTML='';
    const card=document.createElement('div'); card.className='card';

    // Toolbar
    const bar=document.createElement('div'); bar.className='toolbar';

    const input=document.createElement('input');
    input.id = 'search-input';
    input.placeholder='Buscar…';
    input.value=state.search;
    input.className='btn';
    input.style.minWidth='120px';
    input.oninput=e=>{
      state.search=e.target.value;
      state.currentPage=1;
      saveSettings();
      render();
    };

    const sel=document.createElement('select');
    sel.className='btn';
    state.pageOptions.forEach(v=>{
      const o=document.createElement('option');
      o.value=v;
      o.textContent=`Top ${v}`;
      if(v===state.pageSize) o.selected=true;
      sel.appendChild(o);
    });
    sel.onchange=e=>{
      state.pageSize=parseInt(e.target.value,10);
      state.currentPage=1;
      saveSettings();
      render();
    };

    const sp=document.createElement('div'); sp.className='spacer';
    const ts=document.createElement('div'); ts.className='status';
    ts.textContent='Actualizado: '+state.ts;

    bar.append(input,sel,sp,ts);
    card.appendChild(bar);

    const opt=document.createElement('div'); opt.className='options';
    const lbl=document.createElement('label');
    const cb=document.createElement('input');
    cb.type='checkbox'; cb.checked=state.showGrandTotal;
    cb.onchange=e=>{
      state.showGrandTotal=e.target.checked;
      saveSettings();
      render();
    };
    lbl.append(cb,document.createTextNode(' Mostrar total general'));
    opt.appendChild(lbl);
    card.appendChild(opt);

    const tree=buildTree();
    const {headers,rows}=flatten(tree);

    const table=document.createElement('table');
    const thead=document.createElement('thead');
    const trh=document.createElement('tr');

    headers.forEach((h,i)=>{
      const th=document.createElement('th');
      th.textContent=h;

      if(state.sort.index===i){
        th.classList.add('sorted');
      }

      const s=document.createElement('span');
      s.className='sort';
      s.textContent=state.sort.index===i
        ? (state.sort.dir==='asc'?'▲':'▼')
        : '↕';

      th.appendChild(s);
      th.onclick=()=>{
        state.sort.index===i
          ? state.sort.dir=state.sort.dir==='asc'?'desc':'asc'
          : (state.sort.index=i,state.sort.dir='asc');
        render();
      };
      trh.appendChild(th);
    });

    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody=document.createElement('tbody');

    rows.forEach(e=>{
      const tr=document.createElement('tr');

      // ✅ NUEVO: clase para TOTAL (solo diseño / sticky)
      if (e.type === 'total') tr.classList.add('mm-total-row');

      const entryMeta = e.row?.[0] || {};
      const entryPath = entryMeta?.path || [];

      // ✅ resaltar si es la fila activa (filtro aplicado)
      const isActive =
        Array.isArray(_lastAppliedPath) &&
        e.type === 'group' &&
        _lastAppliedPath.length === entryPath.length &&
        _lastAppliedPath.every((v,i)=> v === entryPath[i]);

      if (isActive) tr.classList.add('mm-active-row');

      // ✅ click en fila aplica filtro SOLO en group (Total NO filtra)
      if (e.type === 'group'){
        tr.style.cursor = 'pointer';
        tr.onclick = async () => { await applyPathFilters(entryPath); };
      } else {
        tr.style.cursor = 'default';
      }

      e.row.forEach((v,ci)=>{
        const td=document.createElement('td');

        if(ci===0){
          const m=v;
          td.className='rowhdr indent-'+Math.min((m.depth||0)-1,6);

          const showToggle =
            e.type === 'group' &&
            m &&
            m.hasChildren === true;

          if (showToggle){
            const o = state.expanded[m.key] === true;
            const t=document.createElement('span');
            t.className='toggle';
            t.textContent=o?'▾':'▸';
            t.onclick=x=>{
              x.stopPropagation();
              state.expanded[m.key]=!o;
              render();
            };
            td.appendChild(t);
          } else {
            const spc=document.createElement('span');
            spc.className='toggle';
            spc.style.visibility='hidden';
            td.appendChild(spc);
          }

          td.appendChild(document.createTextNode(m.name ?? ''));

          if (isActive){
            const ico = document.createElement('span');
            ico.className = 'mm-active-ico';
            ico.textContent = ' 🔄';
            td.appendChild(ico);
          }

        }else{
          td.className='right';
          td.textContent=v;
        }

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    card.appendChild(table);

    // Pager (igual a tu original)
    if(state.totalLevel1 > state.pageSize){
      const pager=document.createElement('div');
      pager.className='options';

      const totalPages=Math.ceil(state.totalLevel1/state.pageSize);

      const prev=document.createElement('button');
      prev.className='btn';
      prev.textContent='◀';
      prev.disabled=state.currentPage===1;
      prev.onclick=()=>{
        state.currentPage--;
        state.expanded={};
        saveSettings();
        render();
      };

      const next=document.createElement('button');
      next.className='btn';
      next.textContent='▶';
      next.disabled=state.currentPage===totalPages;
      next.onclick=()=>{
        state.currentPage++;
        state.expanded={};
        saveSettings();
        render();
      };

      const info=document.createElement('span');
      info.className='small';
      info.textContent=`Página ${state.currentPage} de ${totalPages}`;

      pager.append(prev,info,next);
      card.appendChild(pager);
    }

    rootEl.appendChild(card);

    if(activeId){
      const el=document.getElementById(activeId);
      if(el) el.focus();
    }
  }

  // ===================== CSS de fila activa + icono =====================
  function injectActiveStyles(){
    if (document.getElementById('mm-active-style')) return;
    const st = document.createElement('style');
    st.id = 'mm-active-style';
    st.textContent = `
      /* === TOTAL: sticky arriba + diseño === */
      tr.mm-total-row{
        position: sticky;
        top: 0;
        z-index: 3;
        background: rgba(15,23,42,0.06) !important; /* fondo ligeramente más oscuro */
        font-weight: 700;                           /* tipografía más fuerte */
      }
      tr.mm-total-row td{
        border-bottom: 2px solid rgba(15,23,42,0.18); /* línea inferior 2px (más delgada) */
      }

      /* === FILA ACTIVA (lo tuyo) === */
      tr.mm-active-row{
        background: rgba(37,99,235,0.14) !important;           /* fondo un poco más oscuro */
        border-left: 4px solid var(--accent, #2563eb);
        border-bottom: 2px solid var(--accent, #2563eb);       /* línea inferior 2px */
        font-weight: 700;                                      /* tipografía más fuerte */
      }
      tr.mm-active-row:hover{
        background: rgba(37,99,235,0.18) !important;
      }
      .mm-active-ico{
        margin-left: 6px;
        font-size: 12px;
        color: var(--accent, #2563eb);
        vertical-align: middle;
      }
    `;
    document.head.appendChild(st);
  }

})();
