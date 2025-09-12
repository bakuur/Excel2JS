(function(){
  const inputEl = document.getElementById('json-input');
  const arraySelect = document.getElementById('array-path-select');
  const analyzeBtn = document.getElementById('analyze-btn');
  const renderBtn = document.getElementById('render-btn');
  const clearBtn = document.getElementById('clear-btn');
  const flattenChk = document.getElementById('flatten');
  const explodeArraysChk = document.getElementById('explode-arrays');
  const tableContainer = document.getElementById('result-table-container');
  const copyCsvBtn = document.getElementById('copy-csv-btn');
  const downloadXlsxBtn = document.getElementById('download-xlsx-btn');

  let parsedRoot = null;
  let lastTableData = null; // { headers:[], rows:[][] }

  analyzeBtn.addEventListener('click', () => {
    const text = inputEl.value.trim();
    if(!text){ alert('Please paste JSON first'); return; }
    try {
      parsedRoot = JSON.parse(text);
    } catch(e){
      alert('Invalid JSON: ' + e.message);
      return;
    }
    const paths = findArrayPaths(parsedRoot, 5, 200); // limit depth & items scanned
    populateArraySelect(paths);
  });

  renderBtn.addEventListener('click', () => {
    if(!parsedRoot){ alert('Analyze JSON first.'); return; }
    const path = arraySelect.value;
    if(!path){ alert('Select an array path.'); return; }
    const targetArray = resolvePath(parsedRoot, path);
    if(!Array.isArray(targetArray)) { alert('Selected path is not an array.'); return; }

    const flatten = flattenChk.checked;
    const explode = explodeArraysChk.checked;

    const { headers, rows } = arrayToTable(targetArray, { flatten, explode });
    lastTableData = { headers, rows };
    renderTable(headers, rows);
  });

  clearBtn.addEventListener('click', () => {
    inputEl.value='';
    arraySelect.innerHTML='';
    tableContainer.innerHTML='';
    parsedRoot=null; lastTableData=null;
  });

  copyCsvBtn.addEventListener('click', () => {
    if(!lastTableData) { alert('Nothing to copy'); return; }
    const csv = toCSV([lastTableData.headers, ... lastTableData.rows]);
    navigator.clipboard.writeText(csv).then(()=>alert('CSV copied'));
  });

  downloadXlsxBtn.addEventListener('click', () => {
    if(!lastTableData){ alert('Nothing to download'); return; }
    const aoa = [lastTableData.headers, ... lastTableData.rows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    XLSX.writeFile(wb, 'json_table.xlsx');
  });

  function populateArraySelect(paths){
    arraySelect.innerHTML='';
    if(paths.length===0){
      const opt = document.createElement('option');
      opt.value=''; opt.textContent='(No arrays of objects found)';
      arraySelect.appendChild(opt);
      return;
    }
    paths.forEach(p => {
      const opt = document.createElement('option');
      opt.value=p.path; opt.textContent=`${p.path}  (len=${p.length}, type=${p.elementType})`;
      arraySelect.appendChild(opt);
    });
  }

  function findArrayPaths(root, maxDepth=5, sampleLimit=100){
    const results=[];
    const stack=[{value:root, path:'$'}];
    const seen=new WeakSet();
    while(stack.length){
      const {value, path} = stack.pop();
      if(value && typeof value==='object'){
        if(seen.has(value)) continue; seen.add(value);
      }
      if(Array.isArray(value)){
        // classify array
        let elementType='mixed';
        let objCount=0, primCount=0;
        for(let i=0;i<value.length && i<sampleLimit;i++){
          const el=value[i];
            if(el && typeof el==='object' && !Array.isArray(el)) objCount++; else primCount++;
        }
        if(objCount>0 && primCount===0) elementType='object';
        else if(primCount>0 && objCount===0) elementType='primitive';
        results.push({path, length:value.length, elementType});
        // continue traversal only if objects inside
        if(elementType==='object'){
          value.slice(0, sampleLimit).forEach((el, idx)=>{
            if(el && typeof el==='object') stack.push({value:el, path: path+`[${idx}]`});
          });
        }
        continue; // do not push array as object below
      }
      if(value && typeof value==='object'){
        if(path.split(/\.|\[/).length-1 >= maxDepth) continue; // depth limit
        Object.keys(value).forEach(k => {
          stack.push({value:value[k], path: path+(path==='$$'?'':'.')+k.replace(/\./g,'_')});
        });
      }
    }
    // filter only arrays containing objects OR primitives (we'll let user decide) but prefer object arrays first
    return results.filter(r=>r.elementType==='object' || r.elementType==='primitive');
  }

  function resolvePath(root, path){
    if(path==='$') return root;
    // parse something like $.users[0].roles
    let current = root; 
    const propRegex = /\.([^\[\.]+)|(\[(\d+)\])/g; // matches .prop or [index]
    let m; let startIndex=1; // skip leading $
    while(m = propRegex.exec(path)){
      if(m[1]){ // property
        const prop = m[1];
        if(current==null) return undefined;
        current = current[prop];
      } else if(m[3]) {
        const idx = parseInt(m[3],10);
        if(!Array.isArray(current)) return undefined;
        current = current[idx];
      }
    }
    return current;
  }

  function arrayToTable(arr, {flatten=true, explode=false}={}){
    // gather headers
    const flattenedRows = [];

    for(const item of arr){
      if(item && typeof item==='object' && !Array.isArray(item)){
        const flat = flatten ? flattenObject(item) : shallowStringify(item);
        flattenedRows.push(flat);
      } else {
        flattenedRows.push({ value: serializePrimitive(item) });
      }
    }

    // if explode: find first key whose value is an array of primitives; expand
    let rowsToProcess = flattenedRows;
    if(explode){
      const explodeKey = detectExplodeKey(flattenedRows);
      if(explodeKey){
        const expanded=[];
        for(const r of flattenedRows){
          const val = r[explodeKey];
            if(Array.isArray(val)){
              if(val.length===0) expanded.push({...r, [explodeKey]:''});
              else val.forEach(v => expanded.push({...r, [explodeKey]: serializePrimitive(v)}));
            } else {
              expanded.push(r);
            }
        }
        rowsToProcess = expanded;
      }
    }

    const headersSet = new Set();
    rowsToProcess.forEach(r => Object.keys(r).forEach(k => headersSet.add(k)));
    const headers = Array.from(headersSet);
    const rows = rowsToProcess.map(r => headers.map(h => r[h] == null ? '' : (Array.isArray(r[h])||typeof r[h]==='object'? JSON.stringify(r[h]) : r[h])));
    return { headers, rows };
  }

  function detectExplodeKey(rows){
    for(const r of rows){
      for(const k of Object.keys(r)){
        const v = r[k];
        if(Array.isArray(v) && v.every(el => el==null || typeof el!=='object')) return k;
      }
    }
    return null;
  }

  function flattenObject(obj, prefix='', out={}){
    for(const [k,v] of Object.entries(obj)){
      const key = prefix? prefix + '.' + k : k;
      if(v && typeof v==='object'){
        if(Array.isArray(v)){
          // keep arrays as is (maybe explode later) but if array of objects -> stringify
          if(v.every(el => el==null || typeof el!=='object')){
            out[key]= v.slice();
          } else {
            out[key]= JSON.stringify(v);
          }
        } else {
          flattenObject(v, key, out);
        }
      } else {
        out[key]=v;
      }
    }
    return out;
  }

  function shallowStringify(obj){
    const out={};
    for(const [k,v] of Object.entries(obj)){
      if(v && typeof v==='object') out[k]= JSON.stringify(v); else out[k]=v;
    }
    return out;
  }

  function serializePrimitive(v){
    return v==null? '' : typeof v==='string'? v : String(v);
  }

  function renderTable(headers, rows){
    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const hr = document.createElement('tr');
    headers.forEach(h=>{ const th=document.createElement('th'); th.textContent=h; hr.appendChild(th); });
    thead.appendChild(hr); table.appendChild(thead);
    const tbody = document.createElement('tbody');
    rows.forEach(r => { const tr=document.createElement('tr'); r.forEach(c=>{ const td=document.createElement('td'); td.textContent=c; tr.appendChild(td); }); tbody.appendChild(tr); });
    table.appendChild(tbody);
    tableContainer.innerHTML='';
    tableContainer.appendChild(table);
  }

  function toCSV(aoa){
    return aoa.map(row => row.map(cell => {
      const s = cell==null? '' : String(cell);
      if(/[",\n]/.test(s)) return '"'+s.replace(/"/g,'""')+'"';
      return s;
    }).join(',')).join('\n');
  }
})();
