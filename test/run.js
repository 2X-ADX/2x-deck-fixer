const fs = require('fs');
const JSZip = require('jszip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const { fixPackage } = require('../src/fixer.js');
(async () => {
  const args = process.argv.slice(2); const clean = args.includes('--clean'); const [src, dst] = args.filter((a) => !a.startsWith('--'));
  const t0 = Date.now();
  const out = await fixPackage(fs.readFileSync(src), { JSZip, DOMParser, XMLSerializer }, { clean });
  fs.writeFileSync(dst, out.data);
  for (const c of out.report.cleaned) console.log(`  slide ${c.slide}: removed "${c.name}" (${c.why})`);
  console.log(JSON.stringify({ ...out.report, cleaned: out.report.cleaned.length }), (Date.now() - t0) + 'ms');
})().catch((e) => { console.error('ERR', e.stack); process.exit(1); });
