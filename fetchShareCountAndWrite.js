// fetchshareCountAndWrite.js — C列URLを対象にシェア回数取得して記録
const { GoogleSpreadsheet } = require('google-spreadsheet');
const axios = require('axios');

// ===== 設定 =====
const SHEET_ID   = '1wVFefWuElsq7krWpZjTVcerYOHX7SeBTQujVXI7bdXk';
const SHEET_NAME = '投稿シェア回数データ';
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
  const now = new Date();
  const utc = now.getTime() + now.getTimezoneOffset() * 60000;
  const jst = new Date(utc + 9 * 3600000);
  const y = jst.getFullYear();
  const m = jst.getMonth() + 1;
  const d = jst.getDate();

  const md  = `${m}/${d}`;
  const ymd = `${y}/${String(m).padStart(2,'0')}/${String(d).padStart(2,'0')}`;
  const iso = `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
  return { md, ymd, iso };
}

async function fetchShareCount(url) {
  try {
    const res = await axios.get(url, {
      headers: { 'User-Agent': 'Mozilla/5.0' },
      timeout: 15000,
      maxContentLength: 20 * 1024 * 1024,
    });
    const html = res.data;

    // 🔍 shareCount の直後の数値を取得
    const match = html.match(/["']?shareCount["']?\s*[:=]\s*(\d+)/i);
    const n = match ? Number(match[1]) : 0;
    return Number.isFinite(n) ? n : 0;

  } catch (err) {
    console.error(`❌ ${url}: ${err.message}`);
    return 0;
  }
}

(async () => {
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

  await sheet.loadCells(`A1:${columnToLetter(colCount)}1`);
  const { md, ymd, iso } = getJstTodayStrings();

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

  if (targetCol === null) {
    for (let col = 1; col < colCount; col++) {
      const c = sheet.getCell(0, col);
      const hasVal = c.value !== null && c.value !== undefined && c.value !== '';
      if (!hasVal) {
        c.value = md;
        targetCol = col;
        break;
      }
    }
    if (targetCol === null) {
      console.error('❌ 空き列がありません（列数を増やしてください）');
      process.exit(1);
    }
    await sheet.saveUpdatedCells();
  }

  const targetColLetter = columnToLetter(targetCol + 1);
  console.log(`🗓 書き込み先ヘッダー列: ${targetColLetter} (index=${targetCol})`);

  // === 💡 C列（インデックス2）をURL列として処理 ===
  const URL_COL_INDEX = 2;

  for (let startRow = 1; startRow < rowCount; startRow += CHUNK_SIZE) {
    const endRow = Math.min(rowCount - 1, startRow + CHUNK_SIZE - 1);
    const startA1 = startRow + 1;
    const endA1   = endRow + 1;

    const urlRange = `${columnToLetter(URL_COL_INDEX + 1)}${startA1}:${columnToLetter(URL_COL_INDEX + 1)}${endA1}`;
    const outRange = `${targetColLetter}${startA1}:${targetColLetter}${endA1}`;

    await sheet.loadCells(urlRange);
    await sheet.loadCells(outRange);

    let wrote = 0;
    for (let r = startRow; r <= endRow; r++) {
      const urlCell = sheet.getCell(r, URL_COL_INDEX);
      const outCell = sheet.getCell(r, targetCol);
      const url     = (urlCell.value || '').toString().trim();

      let shareCount = 0;
      if (url && url.startsWith('http') && url.includes('tiktok.com')) {
        shareCount = await fetchShareCount(url);
      }

      outCell.value = shareCount;
      outCell.numberFormat = { type: 'NUMBER', pattern: '0' };
      wrote++;
      console.log(`✅ 行${r + 1}: ${shareCount}`);
    }

    await sheet.saveUpdatedCells();
    console.log(`💾 保存: 行${startA1}-${endA1}（${wrote}件更新）`);
  }

  console.log('🎉 完了');
})().catch(err => {
  console.error('❌ Fatal:', err?.stack || err);
  process.exit(1);
});
