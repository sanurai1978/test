// HTML 版を自動操作して、在庫タブと品目タブの数字を取り出す（段階1の期待値）
const { chromium } = require('playwright');
(async () => {
  const file = process.argv[2];
  const b = await chromium.launch(); const p = await b.newPage();
  await p.goto('file://' + file); await p.evaluate(() => localStorage.clear()); await p.reload();
  for (const div of ['kako', 'kumi']) {
    await p.click('#div-' + div);
    await p.click('[data-tab="stock"]');
    const rows = await p.$$eval('#view .row', rs => rs.map(r => r.querySelector('.rcode').textContent.split(' · ')[0] + ' 在庫' + r.querySelector('.rnum b').textContent + ' ' + r.querySelector('.rsub').textContent + ' ' + [...r.querySelectorAll('.rname .pill')].map(x=>x.textContent).join('')));
    console.log('== 在庫', div); rows.forEach(r => console.log(r));
    await p.click('[data-tab="im"]');
    const im = await p.$$eval('#view .row', rs => rs.map(r => r.querySelector('.rcode').textContent.split(' · ')[0] + ' ' + r.querySelector('.rnum b').textContent));
    console.log('== 標準原価', div, im.join(' | '));
    const st = await p.$$eval('#stats .stat', s => s.map(x => x.textContent).join(' '));
    console.log('== 件数', st);
  }
  await b.close();
})();
