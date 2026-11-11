const fs=require('fs');const R='/Users/egecan/Code/accessible-spreadsheet-component/src/engine/__tests__/conformance/';
const cases=JSON.parse(fs.readFileSync('cases.json'));const res=JSON.parse(fs.readFileSync('results.json'));
const files={};const norm=g=>g===''?null:g;
const same=(a,e)=>{ if(Array.isArray(e)) return Array.isArray(a)&&a.length===e.length&&e.every((r,i)=>r.length===a[i].length&&r.every((x,j)=>same(a[i][j],x)));
 if(typeof e==='number'&&typeof a==='number') return e===0?Math.abs(a)<=1e-12:Math.abs(a-e)<=1e-9*Math.abs(e);
 if(e===null) return a===''||a===null||a===0; return a===e;};
const buckets={match:[],nameErr:[],diff:[],err:[]};
cases.forEach((c,k)=>{files[c.file]??=JSON.parse(fs.readFileSync(R+c.file,'utf8'));const fx=files[c.file][c.i];const got=res[k];
 const flat=JSON.stringify(got);
 if(got&&got.err){buckets.err.push([c.file,c.i,c.f,got.err]);return;}
 if(flat.includes('"#NAME?"')&&!JSON.stringify(fx.e).includes('#NAME?')){buckets.nameErr.push([c.file,c.i,c.f]);return;}
 if(same(got,fx.e)){buckets.match.push([c.file,c.i]);} else buckets.diff.push([c.file,c.i,c.f,fx.e,got]);});
for(const k in buckets)console.log(k,buckets[k].length);
fs.writeFileSync('buckets.json',JSON.stringify(buckets));
const byFile={};buckets.diff.forEach(d=>byFile[d[0]]=(byFile[d[0]]||0)+1);console.log(byFile);
const nn={};buckets.nameErr.forEach(d=>nn[d[0]]=(nn[d[0]]||0)+1);console.log('nameErr by file',nn);
