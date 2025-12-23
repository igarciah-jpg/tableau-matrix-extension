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
    expanded: {},                 // colapsado por defecto
    sort: { index: -1, dir: 'asc' },
    ts: null,
    showGrandTotal: true
  };

  let worksheet = null;

  // ===================== Init =====================
  window.addEventListener('load', async () => {
    await tableau.extensions.initializeAsync();
    worksheet = tableau.extensions.worksheetContent.worksheet;
    worksheet.addEventListener(
      tableau.TableauEventType.SummaryDataChanged,
      refresh
    );
    await refresh();
  });

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

    Object.assign(state,{
      columns: cols,
      rows,
      dims: det.dims,
      meas: ord.measOrder,
      dimOrder: ord.dimOrder,
      expanded: {},
      ts: new Date().toLocaleTimeString(),
      loading:false
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

  // ===================== Flatten =====================
  function flatten(tree){
    const headers=['Jerarquía'].concat(state.meas.map(i=>cleanMeasureName(state.columns[i])));
    const out=[];

    // Total arriba
    if(state.showGrandTotal){
      const row=new Array(headers.length).fill('');
      row[0]={depth:0,key:'::total',name:'Total',path:[]};
      state.meas.forEach((_,i)=>row[1+i]=tree.agg[i].toLocaleString());
      out.push({type:'total',row});
    }

    function cmp(a,b){
      const i=state.sort.index,d=state.sort.dir;
      if(i===0) return d==='desc'?b.name.localeCompare(a.name):a.name.localeCompare(b.name);
      return d==='desc'?b.agg[i-1]-a.agg[i-1]:a.agg[i-1]-b.agg[i-1];
    }

    function walk(n){
      let ch=[...n.children.values()];
      if(state.sort.index>=0) ch.sort(cmp);
      ch.forEach(c=>{
        const row=new Array(headers.length).fill('');
        row[0]={depth:c.depth,key:c.key,name:c.name,path:c.path};
        state.meas.forEach((_,i)=>row[1+i]=c.agg[i].toLocaleString());
        out.push({type:'group',row});
        if(state.expanded[c.key]) walk(c);
      });
    }

    walk(tree);
    return {headers,rows:out};
  }

  // ===================== Render =====================
  function render(){
    const root=$('#root'); root.innerHTML='';
    const card=document.createElement('div'); card.className='card';

    // Toolbar
    const bar=document.createElement('div'); bar.className='toolbar';
    const sp=document.createElement('div'); sp.className='spacer';
    const ts=document.createElement('div'); ts.className='status';
    ts.textContent='Actualizado: '+state.ts;
    bar.append(sp,ts); card.appendChild(bar);

    // Checkbox total
    const opt=document.createElement('div'); opt.className='options';
    const lbl=document.createElement('label');
    const cb=document.createElement('input');
    cb.type='checkbox'; cb.checked=state.showGrandTotal;
    cb.onchange=e=>{state.showGrandTotal=e.target.checked; render();};
    lbl.append(cb,document.createTextNode(' Mostrar total general'));
    opt.appendChild(lbl); card.appendChild(opt);

    const tree=buildTree();
    const {headers,rows}=flatten(tree);

    const table=document.createElement('table');
    const thead=document.createElement('thead');
    const trh=document.createElement('tr');

    headers.forEach((h,i)=>{
      const th=document.createElement('th'); th.textContent=h;
      const s=document.createElement('span'); s.className='sort';

      if(state.sort.index===i){
        s.textContent=state.sort.dir==='asc'?'▲':'▼';
        s.style.color='#2563eb';
        s.style.fontWeight='600';
      } else {
        s.textContent='↕';
        s.style.color='#94a3b8';
      }

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

      // Diseño + hover elegante
      if(e.type==='total'){
        tr.style.background='#f8fafc';
        tr.style.fontWeight='600';
        tr.style.borderTop='2px solid #cbd5e1';
      } else {
        const d=e.row[0].depth||1;
        const baseBg=(d%2)?'#ffffff':'#f8fafc';
        tr.style.background=baseBg;
        tr.style.color='#1f2937';
        tr.onmouseenter=()=>{ tr.style.background='#eef2f7'; };
        tr.onmouseleave=()=>{ tr.style.background=baseBg; };
      }

      e.row.forEach((v,ci)=>{
        const td=document.createElement('td');
        if(ci===0){
          const m=v;
          td.className='rowhdr indent-'+Math.min(m.depth-1,6);
          if(e.type==='group'){
            const o=state.expanded[m.key]===true;
            const t=document.createElement('span');
            t.className='toggle';
            t.textContent=o?'▾':'▸';
            t.onclick=x=>{
              x.stopPropagation();
              state.expanded[m.key]=!o;
              render();
            };
            td.appendChild(t);
          }
          td.appendChild(document.createTextNode(m.name));
        } else {
          td.className='right';
          td.textContent=v;
        }
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    card.appendChild(table);
    root.appendChild(card);
  }

})();
