// HTML 版で段階2のシナリオを操作し、各時点の MRP 手配案・製造指図・発注を JSON で出す
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage();
  await p.goto('file://' + process.argv[2]); await p.evaluate(() => localStorage.clear()); await p.reload();
  const out = {};
  const tab = async t => p.click(`#tabs [data-tab="${t}"]`);
  const div = async d => p.click('#div-' + d);
  const mrp = async () => { await tab('mrp');
    return p.$$eval('#view > .card', cs => cs.filter(c => c.querySelector('[data-act="mrpgo"]')).map(c => {
      const pills = [...c.querySelectorAll('.pill')].map(x => x.textContent);
      const calc = [...c.querySelectorAll('.calc')].map(x => x.textContent);
      const btn = c.querySelector('[data-act="mrpgo"]').textContent;
      const m = btn.match(/ ([\d.,]+)\D*を作成/);
      const when = calc[1];
      let st, du, nd;
      let w = when.match(/着手 (\S+) → 完成 (\S+?) ／/); if (w) { st = w[1]; du = w[2]; }
      w = when.match(/入荷 (\S+?)（/); if (w) { st = '今日'; du = w[1]; }
      w = when.match(/必要日 (\S+)/); nd = w ? w[1] : '補充';
      return { kind: pills[0], code: c.querySelector('.rcode').textContent, qty: Number(m[1].replace(/,/g, '')), st, du, nd,
               so: (pills.find(x => / 向け$/.test(x)) || '').replace(' 向け', '').replace('親品目', ''), late: pills.includes('必要日に間に合わない') };
    })); };
  const mos = async () => { await tab('mo');
    return p.$$eval('#view .row', rs => rs.map(r => { const c = r.querySelector('.rcode').textContent; const s = r.querySelector('.rsub').textContent;
      const d = s.match(/(\S+) → (\S+)$/);
      return { no: c.split(' · ')[0], so: (c.match(/(SO-\d+) 向け/) || [, ''])[1], qty: Number(r.querySelector('.rnum b').textContent.replace(/,/g, '')),
               st: r.querySelector('.rname .pill').textContent, start: d ? d[1] : '', due: d ? d[2] : '', name: r.querySelector('.rname').childNodes[0].textContent.trim() }; })); };
  const pos = async () => { await tab('po');
    return p.$$eval('#view .row', rs => rs.map(r => { const c = r.querySelector('.rcode').textContent;
      return { no: c.split(' · ')[0], eta: c.split('入荷予定 ')[1], qty: Number(r.querySelector('.rnum b').textContent.replace(/,/g, '')), name: r.querySelector('.rname').childNodes[0].textContent.trim() }; })); };
  const stats = async () => p.$$eval('#stats .stat b', s => s.map(x => Number(x.textContent)));
  const form = async (vals, i) => { for (const [k, v] of Object.entries(vals)) { const el = await p.$('#f-' + k);
      if ((await el.evaluate(e => e.tagName)) === 'SELECT') await el.selectOption(String(v)); else await el.fill(String(v)); }
    await p.click(`#sh-acts [data-act="shact"][data-i="${i || 0}"]`); };
  const closeIf = async () => { if (!(await p.$eval('#panel', e => e.hidden))) await p.click('[data-act="shclose"]'); };
  const load = async () => { await tab('plan');
    return p.$$eval('table.load tbody tr', rs => rs.map(r => [...r.querySelectorAll('td')].map((c, i) => i === 0 ? c.textContent : (c.textContent === '·' ? 0 : Number(c.textContent))))); };
  const snap = async k => { out[k] = { mrp: await mrp(), mo: await mos(), po: await pos(), stats: await stats(), load: await load() }; };

  await div('kako'); await snap('A_加工_初期');
  await tab('mrp'); await p.click('[data-act="mrpall"]'); await snap('B_加工_一括手配後');
  await tab('so'); await p.click('[data-act="sonew"]'); await form({ i: 'P-103', q: 5, c: '北斗精工', d: 4 }); await snap('C_加工_受注P103');
  await tab('so'); await p.click('[data-act="so"][data-no="SO-0005"]'); await p.click('#sh-acts [data-act="shact"][data-i="1"]'); await closeIf(); await snap('D_加工_受注向け指図');
  await tab('fc'); await p.click('[data-act="fcnew"]'); await form({ i: 'P-101', q: 50, d: 20 }); await snap('E_加工_見込追加');
  await tab('plan'); await p.click('[data-act="plan"][data-no="MO-0001"]'); await form({ s: 5 }); await closeIf(); await snap('F_加工_着手日変更');
  await div('kumi'); await snap('G_組立_初期');
  await tab('mo'); await p.click('[data-act="mo"][data-no="MO-0002"]');
  const acts = await p.$$eval('#sh-acts [data-act="shact"]', a => a.map(x => x.textContent)); await p.click(`#sh-acts [data-act="shact"][data-i="${acts.indexOf('製造指図を取り消す')}"]`); await snap('H_組立_指図取消');
  await tab('so'); await p.click('[data-act="sonew"]'); await form({ i: 'A-301', q: 2, c: '東和機械', d: 30 }); await snap('I_組立_受注A301');
  await tab('mrp'); await p.click('[data-act="mrpall"]'); await snap('J_組立_一括手配後');
  await tab('po'); await p.click('[data-act="ponew"]'); await form({ i: 'B-004', q: 30 }); await snap('K_組立_手動発注');
  await tab('so'); await p.click('[data-act="so"][data-no="SO-0003"]'); const a2 = await p.$$eval('#sh-acts [data-act="shact"]', a => a.map(x => x.textContent));
  await p.click(`#sh-acts [data-act="shact"][data-i="${a2.indexOf('引当を解除')}"]`); out.L_引当解除_エラー = await p.$eval('#sh-err', e => e.textContent); await closeIf();
  console.log(JSON.stringify(out));
  await b.close();
})();
