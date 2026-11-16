const fs = require('fs');
const JSZip = require('jszip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const { fixPackage } = require('../src/fixer.js');
(async () => {
  const [src, dst] = process.argv.slice(2);
  const t0 = Date.now();
  const out = await fixPackage(fs.readFileSync(src), { JSZip, DOMParser, XMLSerializer });
  fs.writeFileSync(dst, out.data);
  console.log(JSON.stringify(out.report), (Date.now() - t0) + 'ms');
})().catch((e) => { console.error('ERR', e.stack); process.exit(1); });
