// Copyright (c) 2026 CAP Index. Licensed under the MIT License.
// Deterministic vector artwork; PNGs are rendered for Office RibbonX.
const fs=require('fs'),path=require('path'),os=require('os');
let sharp;
try { sharp=require('sharp'); }
catch(error) {
 if(error.code!=='MODULE_NOT_FOUND')throw error;
 sharp=require(path.join(os.homedir(),'.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/sharp'));
}
const root=path.resolve(__dirname,'..'), dir=path.join(root,'src/ribbon/icons');fs.mkdirSync(dir,{recursive:true});
const navy='#211c65', lime='#d7f56b', pale='#eef0f8';
const rect=(x,y,w,h,fill=pale)=>`<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="1.5" fill="${fill}"/>`;
const line=(d)=>`<path d="${d}" fill="none"/>`;
const circ=(x,y,r,fill=pale)=>`<circle cx="${x}" cy="${y}" r="${r}" fill="${fill}"/>`;
const check=line('M7 12l3 3 7-8');
const sheet=rect(4,2,16,20)+line('M8 7h8M8 11h8M8 15h5');
const calendar=rect(2,4,20,18)+line('M2 9h20M7 2v5M17 2v5');
const clock=circ(15.5,15.5,7,lime)+line('M15.5 11v4.5l3 2');
const pin=`<path d="M12 22s8-8 8-13a8 8 0 0 0-16 0c0 5 8 13 8 13Z" fill="${pale}"/>`+circ(12,9,3,lime);
const grid=rect(3,3,18,18)+line('M3 9h18M3 15h18M9 3v18M15 3v18');
const bars=line('M3 2v20h19')+rect(6,12,3,7)+rect(12,8,3,11,lime)+rect(18,4,3,15);
const defs={
 check:sheet+circ(17,17,6,lime)+line('M14 17l2 2 4-5'),
 review:rect(2,4,12,16)+rect(7,2,12,16)+circ(16,15,5,lime)+line('M19.5 18.5l3 3'),
 timestamp:calendar+clock,
 hour:circ(12,12,10)+line('M12 5v7l5 3')+circ(12,12,1,lime),
 block:rect(2,5,20,14)+line('M2 12h20M9 5v14M15 5v14')+rect(10,6,4,5,lime),
 random:rect(3,3,18,18)+circ(7,7,1,lime)+circ(17,7,1,lime)+circ(12,12,1,lime)+circ(7,17,1,lime)+circ(17,17,1,lime),
 geocode:pin,
 crosswalk:rect(2,3,8,18)+rect(14,3,8,18,lime)+line('M6 8h1M6 12h1M6 16h1M10 12h4m-2-2 2 2-2 2'),
 classification:`<path d="M2 3h11l9 9-10 10L2 12Z" fill="${pale}"/>`+circ(7,8,1.5,lime)+line('M11 13l3 3 4-5'),
 scopeall:grid,
 scopevisible:grid+`<path d="M4 4h16l-6 7v7l-4 3V11Z" fill="${lime}"/>`,
 site:circ(12,12,10)+circ(12,12,6)+line('M12 1v4M12 19v4M1 12h4M19 12h4')+circ(12,12,2,lime),
 concentration:bars,
 heatmap:rect(2,2,20,20)+rect(3,3,5,5,lime)+rect(9.5,3,5,5,'#ffe1b2')+rect(16,3,5,5,'#e97950')+rect(3,9.5,5,5,'#ffe1b2')+rect(9.5,9.5,5,5,'#e97950')+rect(16,9.5,5,5,lime)+rect(3,16,5,5,lime)+rect(9.5,16,5,5,pale)+rect(16,16,5,5,'#ffe1b2'),
 repeat:pin+line('M3 20H1v-5M1 20l4-4M21 3h2v5M23 3l-4 4'),
 hotstreets:line('M3 2v20h19M5 18l5-7 4 3 6-10')+circ(20,4,2,lime),
 trends:line('M3 2v20h19M6 17l4-4 4 2 7-9')+line('M17 6h4v4'),
 yoy:rect(2,3,8,17)+rect(14,3,8,17)+line('M5 7h2M17 7h2M10 12h4m-2-2 2 2-2 2')+rect(5,15,2,3,lime)+rect(17,11,2,7,lime),
 seasonality:calendar+line('M6 13h3M6 17h3M13 13h3M13 17h3')+circ(18,18,4,lime),
 compare:line('M3 22h18')+rect(5,9,5,11)+rect(14,3,5,17,lime)+line('M4 4h6m-2-2 2 2-2 2'),
 profile:sheet+circ(16,16,6,lime)+line('M16 10v6h6'),
 report:sheet+rect(7,14,3,5,lime)+rect(12,12,3,7)+rect(17,16,2,3,lime),
 selectreport:sheet+circ(17,17,6,lime)+line('M14 17l2 2 4-5'),
 export:sheet+line('M13 16h10m-3-3 3 3-3 3'),
 image:rect(2,4,20,16)+circ(7,9,2,lime)+`<path d="M3 19l6-6 4 4 4-7 4 9Z" fill="${lime}"/>`,
 powerpoint:rect(3,3,18,14)+line('M12 17v5M8 22h8M7 12l4-4 3 2 4-4'),
 pdf:sheet+line('M8 17c4-7 3-11 2-10s0 10 7 11c4 0-2-5-9-1Z'),
 refresh:line('M20 9a8 8 0 0 0-14-4L3 8m0-5v5h5M4 15a8 8 0 0 0 14 4l3-3m0 5v-5h-5'),
 reset:line('M7 4l-5 5 5 5M2 9h11a7 7 0 0 1 0 14')+circ(17,8,4,lime)+line('M15 8h4'),
 guide:rect(3,2,18,20)+line('M8 2v20M12 7h5M12 11h5M12 15h3'),
 about:circ(12,12,10)+circ(12,7,1,lime)+line('M12 11v7')
};
(async()=>{
 const preview=[];
 for(const [name,body] of Object.entries(defs)){
  const outline=body.replace(/fill="[^"]*"/g,'fill="none"');
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="-2 -2 28 28"><g stroke="white" stroke-width="4" fill="none" stroke-linecap="round" stroke-linejoin="round">${outline}</g><g stroke="${navy}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${body}</g></svg>`;
  fs.writeFileSync(path.join(dir,name+'.svg'),svg);
  for(const size of [16,32])await sharp(Buffer.from(svg)).resize(size,size).png().toFile(path.join(dir,`${name}-${size}.png`));
  preview.push({name,svg});
 }
 // Preserve the official artwork. An SVG viewport isolates the circular mark
 // from the left of the full wordmark for a square ribbon control.
 const brandBytes=fs.readFileSync(path.join(root,'src/ribbon/brand/CAP-Index-Full-Color-Logo.png'));
 const brandSVG=`<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="-110 -110 2504 2504"><circle cx="1142" cy="1142" r="1242" fill="white"/><image x="0" y="0" width="6040" height="2284" href="data:image/png;base64,${brandBytes.toString('base64')}"/></svg>`;
 fs.writeFileSync(path.join(dir,'cap-logo.svg'),brandSVG);
 for(const size of [16,32])await sharp(Buffer.from(brandSVG)).resize(size,size).png().toFile(path.join(dir,`cap-logo-${size}.png`));
 const w=840,h=Math.ceil((preview.length+1)/7)*105;
 let cells='';preview.forEach(({name,svg},i)=>{const x=(i%7)*120,y=Math.floor(i/7)*105;cells+=`<g transform="translate(${x+40},${y+12})">${svg.replace('width="32" height="32"','width="40" height="40"')}</g><text x="${x+60}" y="${y+78}" text-anchor="middle">${name}</text>`;});
 const logo=fs.readFileSync(path.join(dir,'cap-logo-32.png')).toString('base64');const i=preview.length,x=(i%7)*120,y=Math.floor(i/7)*105;
 cells+=`<image x="${x+40}" y="${y+12}" width="40" height="40" href="data:image/png;base64,${logo}"/><text x="${x+60}" y="${y+78}" text-anchor="middle">CAP logo</text>`;
 await sharp(Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><rect width="100%" height="100%" fill="white"/><g font-family="Arial" font-size="13" fill="${navy}">${cells}</g></svg>`)).png().toFile(path.join(root,'docs/ribbon-icon-preview.png'));
 await sharp(Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><rect width="100%" height="100%" fill="#292929"/><g font-family="Arial" font-size="13" fill="#f4f4f4">${cells}</g></svg>`)).png().toFile(path.join(root,'docs/ribbon-icon-preview-dark.png'));
 fs.writeFileSync(path.join(dir,'manifest.json'),JSON.stringify(Object.keys(defs).concat('cap-logo'),null,2)+'\n');
 console.log(`${Object.keys(defs).length+1} icon families rendered at 16 and 32 pixels.`);
})();
