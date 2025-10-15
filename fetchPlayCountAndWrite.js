// fetchPlayCountAndWrite.js — ヘッダーは文字列、100行チャンク版
const { GoogleSpreadsheet } = require('google-spreadsheet');
const axios = require('axios');

// ===== 設定 =====
const SHEET_ID   = '1vm6JyX8a8Bt5FX4xE6CGgtRUYMh7gXzHhPAxnEooH_8';
const SHEET_NAME = 'Vin計測ツール';
const CHUNK_SIZE = 100;        // 100行ごとに処理
// =================

function columnToLetter(col) {
  let temp = '', letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

function getJstTodayStrings() {
  // JSTの今日。ヘッダーは単純に "M/D" の文字列で書く
  const now = new Date();
  const utc = now.getTime() + now.getTimezoneOffset() * 60000;
  const jst = new Date(utc + 9 * 3600000);
  const y = jst.getFullYear();
  const m = jst.getMonth() + 1;
  const d = jst.getDate();

  const md  = `${m}/${d}`;                         // 例: "8/22"
  const ymd = `${y}/${String(m).padStart(2,'0')}/${String(d).padStart(2,'0')}`; // 例: "2025/08/22"
  const iso = `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`; // 例: "2025-08-22"
  return { md, ymd, iso };
}

async function fetchPlayCount(url) {
  try {
    const res = await axios.get(url, {
      headers: { 'User-Agent': 'Mozilla/5.0' },
      timeout: 15000,
      maxContentLength: 20 * 1024 * 1024, // 20MBガード
    });
    const html = res.data;
    const match = html.match(/["']?playCount["']?\s*[:=]\s*(\d+)/i);
    const n = match ? Number(match[1]) : 0;
    return Number.isFinite(n) ? n : 0;
  } catch (err) {
    console.error(`❌ ${url}: ${err.message}`);
    return 0;
  }
}

(async () => {
  // 認証（環境変数 GOOGLE_CREDS_BASE64 を想定）
  const creds = JSON.parse(
    Buffer.from(process.env.GOOGLE_CREDS_BASE64, 'base64').toString('utf-8')
  );

  const doc = new GoogleSpreadsheet(SHEET_ID);
  await doc.useServiceAccountAuth(creds);
  await doc.loadInfo();

  const sheet = doc.sheetsByTitle[SHEET_NAME];
  if (!sheet) {
    console.error(`❌ シート「${SHEET_NAME}」が見つかりません`);
    process.exit(1);
  }

  const rowCount = sheet.rowCount;
  const colCount = sheet.columnCount;

  // 1) ヘッダー1行だけ読み込む（軽量）
  await sheet.loadCells(`A1:${columnToLetter(colCount)}1`);

  const { md, ymd, iso } = getJstTodayStrings();

  // 2) 今日の列（targetCol）を探す（A列=URLは除外、B列=1から検索）
  let targetCol = null;
  for (let col = 1; col < colCount; col++) {
    const c = sheet.getCell(0, col);
    const raw  = (c.value ?? '').toString().trim();
    const disp = (c.displayValue ?? '').toString().trim();
    if ([raw, disp].some(v => v === md || v === ymd || v === iso)) {
      targetCol = col;
      break;
    }
  }

  // 3) なければ最初の空き列に "M/D" 文字列で作成（形式指定なし）
  if (targetCol === null) {
    for (let col = 1; col < colCount; col++) {
      const c = sheet.getCell(0, col);
      const hasVal = c.value !== null && c.value !== undefined && c.value !== '';
      if (!hasVal) {
        c.value = md; // ただの文字列でOK（後でGASが整形する想定）
        targetCol = col;
        break;
      }
    }
    if (targetCol === null) {
      console.error('❌ 空き列がありません（列数を増やしてください）');
      process.exit(1);
    }
    await sheet.saveUpdatedCells(); // ヘッダー書き込みを反映
  }

  const targetColLetter = columnToLetter(targetCol + 1);
  console.log(`🗓  書き込み先ヘッダー列: ${targetColLetter} (index=${targetCol})`);

  // 4) 本体は100行ずつ、A列と書き込み列だけ読み書き
  for (let startRow = 1; startRow < rowCount; startRow += CHUNK_SIZE) {
    const endRow = Math.min(rowCount - 1, startRow + CHUNK_SIZE - 1);

    const aStart = startRow + 1; // A1基準に変換
    const aEnd   = endRow + 1;

    const urlRange = `A${aStart}:A${aEnd}`;
    const outRange = `${targetColLetter}${aStart}:${targetColLetter}${aEnd}`;

    await sheet.loadCells(urlRange);
    await sheet.loadCells(outRange);

    let wrote = 0;

    for (let r = startRow; r <= endRow; r++) {
      const urlCell = sheet.getCell(r, 0);         // A列（URL）
      const outCell = sheet.getCell(r, targetCol); // 今日の列
      const url     = (urlCell.value || '').toString().trim();

      let playCount = 0;
      if (url && url.startsWith('http') && url.includes('tiktok.com')) {
        playCount = await fetchPlayCount(url);
      } else {
        playCount = 0; // 無効URL/空白は 0 記録
      }

      if (!Number.isFinite(playCount)) playCount = 0;

      outCell.value = playCount; // 数値で書く
      outCell.numberFormat = { type: 'NUMBER', pattern: '0' };
      wrote++;
      console.log(`✅ 行${r + 1} → ${playCount}`);
    }

    await sheet.saveUpdatedCells();
    console.log(`💾 保存: 行${aStart}-${aEnd}（${wrote}件更新）`);
  }

  console.log('🎉 完了');
})().catch(err => {
  console.error('❌ Fatal:', err?.stack || err);
  process.exit(1);
});
