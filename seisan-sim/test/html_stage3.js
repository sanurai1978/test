// HTML 版で段階3のシナリオを操作し、各時点のデータ（localStorage の中身）を JSON で出す
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage();
  await p.goto('file://' + process.argv[2]); await p.evaluate(() => localStorage.clear()); await p.reload();
  const out = {}, errs = {};
  const tab = async t => p.click(`#tabs [data-tab="${t}"]`);
  const div = async d => p.click('#div-' + d);
  const act = async label => { const as = await p.$$eval('#sh-acts [data-act="shact"]', a => a.map(x => x.textContent));
    const i = as.findIndex(x => x.startsWith(label)); if (i < 0) throw new Error('no action ' + label + ' in ' + as.join('|'));
    await p.click(`#sh-acts [data-act="shact"][data-i="${i}"]`); };
  const fill = async vals => { for (const [k, v] of Object.entries(vals)) { const el = await p.$('#f-' + k);
      if ((await el.evaluate(e => e.tagName)) === 'SELECT') await el.selectOption(String(v)); else await el.fill(String(v)); } };
  const err = async () => p.$eval('#sh-err', e => e.textContent);
  const close = async () => { if (!(await p.$eval('#panel', e => e.hidden))) await p.click('[data-act="shclose"]'); };
  const openMO = async no => { await close(); await tab('mo'); await p.click(`[data-act="mo"][data-no="${no}"]`); };
  const opDefault = async (label) => { await act(label); await act('実績を登録'); };      // 既定値のまま登録
  const snap = async k => { await close(); out[k] = await p.evaluate(() => JSON.parse(localStorage.getItem('seisan-sim-v6'))); };

  await div('kako');
  await tab('sc'); await p.click('[data-act="sc"][data-no="SC-0001"]'); await fill({ g: 38, n: 2 }); await act('受入を計上'); await snap('S1_外注受入');
  await openMO('MO-0002'); await act('工程 40'); await fill({ g: 37, n: 1, t: 15 }); await act('実績を登録'); await snap('S2_実績で完成');
  await openMO('MO-0001'); await act('ピッキング指示を出す'); await act('残りをまとめて'); await act('払出を確定して着手');
  await act('工程 10'); await fill({ g: 20, n: 0, t: 95 }); await act('実績を登録');
  await act('工程 20「熱処理」を外注に出す'); await act('工程 20「熱処理」の外注受入'); await fill({ g: 19, n: 1 }); await act('受入を計上');
  await act('工程 30'); await fill({ g: 19, n: 0, t: 15 }); await act('実績を登録'); await snap('S3_払出から完成まで');
  await close(); await tab('pk'); await p.click('[data-act="pk"][data-no="PK-0001"]'); await act('残りをまとめて'); await act('出荷を確定'); await snap('S4_出荷ピッキング確定');
  await close(); await tab('so'); await p.click('[data-act="sonew"]'); await fill({ i: 'P-101', q: 30, c: '北斗精工', d: 5 }); await act('受注を登録');
  await tab('so'); await p.click('[data-act="so"][data-no="SO-0005"]'); await act('在庫を引き当てる');
  await tab('ship'); await p.click('[data-act="shipgo"][data-no="SO-0005"]'); await fill({ q: 20, c: '宅配便' }); await act('ピッキング指示を出す');
  await act('残りをまとめて'); await act('出荷を確定'); await snap('S5_分納出荷');
  await close(); await tab('ship'); await p.click('[data-act="sh"][data-no="SH-0002"]'); await act('この出荷を取り消す'); await snap('S6_出荷取消');
  await close(); await tab('stock'); await p.click('[data-act="item"][data-code="M-001"]'); await act('在庫調整'); await fill({ q: 390 }); await act('在庫を更新'); await snap('S7_在庫調整');
  await close(); await tab('so'); await p.click('[data-act="so"][data-no="SO-0004"]'); await act('この受注向けの製造指図を作る');
  await openMO('MO-0003'); await act('ピッキング指示を出す'); await act('残りをまとめて'); await act('払出を確定して着手');
  await opDefault('工程 10'); await opDefault('工程 20'); await opDefault('工程 30'); await snap('S8_受注生産の完成と自動引当');
  await div('kumi');
  await tab('pk'); await p.click('[data-act="pk"][data-no="PK-0001"]'); await act('残りをまとめて'); await act('払出を確定して着手');
  await opDefault('工程 10'); await opDefault('工程 20'); await opDefault('工程 30'); await snap('S9_組立_払出から完成');
  await openMO('MO-0001'); await act('工程 20'); await fill({ g: 9, n: 1, t: 30 }); await act('実績を登録'); await snap('S10_組立_実績');
  await close(); await tab('po'); await p.click('[data-act="po"][data-no="PO-0001"]'); await act('入荷を計上'); await snap('S11_入荷');
  await close(); await tab('ship'); await p.click('[data-act="shipgo"][data-no="SO-0003"]'); await fill({ q: 4, c: '路線便' }); await act('ピッキング指示を出す');
  await act('残りをまとめて'); await act('出荷を確定'); await snap('S12_組立_出荷');
  // エラーになる操作
  await close(); await tab('mo'); await p.click('[data-act="monew"]'); await fill({ i: 'A-201', q: 100, s: 0 }); await act('製造指図を作成');
  await openMO('MO-0003'); await act('ピッキング指示を出す'); errs.E1_部材不足 = await err();
  console.log(JSON.stringify({ out, errs }));
  await b.close();
})();
