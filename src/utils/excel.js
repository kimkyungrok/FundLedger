// src/utils/excel.js
import ExcelJS from 'exceljs';

/* ===== 공용 유틸 ===== */

// Excel XML에서 허용되지 않는 문자 제거
function sanitizeXmlText(v) {
  if (v == null) return '';
  let o = '';
  for (const ch of String(v)) {
    const cp = ch.codePointAt(0);
    if (
      (cp >= 0x20 && cp <= 0xd7ff) ||
      (cp >= 0xe000 && cp <= 0xfffd) ||
      (cp >= 0x10000 && cp <= 0x10ffff) ||
      cp === 0x09 ||
      cp === 0x0a ||
      cp === 0x0d
    ) {
      o += ch;
    }
  }
  return o;
}

const toNum = (v, d = 0) => {
  const n = Number(v);
  return Number.isFinite(n) ? n : d;
};

// 열 폭 (A~G) - 엑셀 시트 열 너비 기준
// A = 4.38, B = 3.38, C = 3.38, D = 19.5, E = 12.5, F = 12.5, G = 15.5
const COL_WIDTHS = [4.38, 3.38, 3.38, 19.5, 12.5, 12.5, 15.5];

// 원화 회계 서식 ("₩", 0일 때 "-")
const ACCOUNT_FMT =
  '_-[$₩-ko-KR]* #,##0_-;_-[$₩-ko-KR]* #,##0_-;_-[$₩-ko-KR]* "-"_-;_-@_-';

function applyThinBorder(cell) {
  cell.border = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } },
  };
}

function applyMoney(cell) {
  cell.numFmt = ACCOUNT_FMT;
  cell.alignment = { horizontal: 'right', vertical: 'center' };
  cell.font = cell.font || { name: '맑은 고딕', size: 10 };
}

/**
 * summary: /entries 에서 계산한 summary 객체
 * rows   : DB에서 가져온 entries 배열
 * options: { start, end, q, order } (필터/정렬 정보)
 */
export async function createWorkbookFromEntries(summary, rows, options = {}) {
  console.log('[excel.js] createWorkbookFromEntries 호출', {
    rowsLen: rows?.length ?? 0,
    summarySnapshot: {
      income: summary?.income,
      expense: summary?.expense,
      balance: summary?.balance,
      prevYear: summary?.detail?.prevYear,
      prevCarry: summary?.detail?.prevCarry,
      year: summary?.detail?.year,
    },
    options,
  });

  const wb = new ExcelJS.Workbook();
  wb.creator = 'fund-ledger';
  wb.created = new Date();

  const ws = wb.addWorksheet('공금수불부');

  // 열 설정
  ws.columns = [
    { header: '년', key: 'year', width: COL_WIDTHS[0] },
    { header: '월', key: 'month', width: COL_WIDTHS[1] },
    { header: '일', key: 'day', width: COL_WIDTHS[2] },
    { header: '적요', key: 'desc', width: COL_WIDTHS[3] },
    { header: '수입', key: 'income', width: COL_WIDTHS[4] },
    { header: '지출', key: 'expense', width: COL_WIDTHS[5] },
    { header: '잔액', key: 'balance', width: COL_WIDTHS[6] },
  ];

  // 2행까지 고정 (1: 제목, 2: 헤더)
  ws.views = [
    {
      state: 'frozen',
      xSplit: 0,
      ySplit: 2,
      topLeftCell: 'A3',
    },
  ];

  const detail = Object.assign(
    {
      prevYear: new Date().getFullYear() - 1,
      year: new Date().getFullYear(),
      prevCarry: 0,
      yearIncome: 0,
      yearExpense: 0,
      yearBalance: 0,
    },
    (summary && summary.detail) || {}
  );

  const totalIncome = toNum(summary?.income);
  const totalExpense = toNum(summary?.expense);
  const totalBalance = toNum(summary?.balance);

  /* ===== 1행: 제목 ===== */
  const titleRow = ws.getRow(1);
  titleRow.getCell(1).value = `${detail.year}년 공금 수불부`;
  titleRow.getCell(1).font = { name: '맑은 고딕', size: 14, bold: true };
  titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'center' };
  ws.mergeCells('A1:G1');

  /* ===== 2행: 머리글 ===== */
  const headerRow = ws.getRow(2);
  headerRow.getCell(1).value = '년';
  headerRow.getCell(2).value = '월';
  headerRow.getCell(3).value = '일';
  headerRow.getCell(4).value = '적요';
  headerRow.getCell(5).value = '수입';
  headerRow.getCell(6).value = '지출';
  headerRow.getCell(7).value = '잔액';

  for (let c = 1; c <= 7; c++) {
    const cell = headerRow.getCell(c);
    cell.font = { name: '맑은 고딕', size: 10, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'center' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE7E6E6' }, // 헤더 배경
    };
    applyThinBorder(cell);
  }

  /* ===== 3행부터: 입출금 내역 ===== */
  const dataStartRow = 3;
  rows = Array.isArray(rows) ? rows : [];
  let running = toNum(detail.prevCarry);

  // 날짜 병합을 위한 메타 저장
  const dateMeta = [];

  rows.forEach((r, idx) => {
    const rowIndex = dataStartRow + idx;
    const row = ws.getRow(rowIndex);

    let y = '';
    let m = '';
    let d = '';
    if (r.date) {
      const raw = String(r.date);
      const m1 = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m1) {
        y = m1[1];
        m = String(Number(m1[2]));
        d = String(Number(m1[3]));
      } else {
        try {
          const dt = new Date(raw);
          if (!Number.isNaN(dt.getTime())) {
            y = String(dt.getFullYear());
            m = String(dt.getMonth() + 1);
            d = String(dt.getDate());
          }
        } catch (_) {}
      }
    }

    const income = toNum(r.income);
    const expense = toNum(r.expense);
    running += income - expense;

    row.getCell(1).value = y || null;
    row.getCell(2).value = m || null;
    row.getCell(3).value = d || null;
    row.getCell(4).value = sanitizeXmlText(r.desc || '');
    row.getCell(5).value = income || null;
    row.getCell(6).value = expense || null;
    row.getCell(7).value = running || null;

    // 연/월/일: 항상 가운데 정렬
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'center' };
    row.getCell(2).alignment = { horizontal: 'center', vertical: 'center' };
    row.getCell(3).alignment = { horizontal: 'center', vertical: 'center' };

    row.getCell(4).alignment = {
      horizontal: 'left',
      vertical: 'center',
      wrapText: true,
    };
    applyMoney(row.getCell(5));
    applyMoney(row.getCell(6));
    applyMoney(row.getCell(7));

    for (let c = 1; c <= 7; c++) {
      const cell = row.getCell(c);
      cell.font = cell.font || { name: '맑은 고딕', size: 10 };
      applyThinBorder(cell);
    }

    dateMeta.push({ rowIndex, y, m, d });
  });

  const lastDataRow = rows.length ? dataStartRow + rows.length - 1 : dataStartRow - 1;

  /* ===== 같은 날짜(연/월/일) 병합 ===== */
  if (dateMeta.length > 1) {
    const sameDate = (a, b) =>
      a.y === b.y && a.m === b.m && a.d === b.d && a.y && a.m && a.d;

    let groupStart = dateMeta[0].rowIndex;
    let prev = dateMeta[0];

    const flushGroup = (startRow, endRow) => {
      if (startRow >= endRow) return; // 한 줄이면 병합 안 함
      try {
        // 년(A), 월(B), 일(C) 열 각각 세로 병합
        for (let col = 1; col <= 3; col++) {
          ws.mergeCells(startRow, col, endRow, col);
          const cell = ws.getCell(startRow, col);
          cell.alignment = { horizontal: 'center', vertical: 'center' };
          applyThinBorder(cell);
        }
      } catch (err) {
        console.error('[excel.js] date merge 실패:', { startRow, endRow }, err);
      }
    };

    for (let i = 1; i < dateMeta.length; i++) {
      const cur = dateMeta[i];
      if (sameDate(prev, cur)) {
        prev = cur;
        continue;
      }
      flushGroup(groupStart, prev.rowIndex);
      groupStart = cur.rowIndex;
      prev = cur;
    }
    flushGroup(groupStart, prev.rowIndex);
  }

  /* ===== 내역 이후: 요약 블록 배치 ===== */
  // 내역 마지막 줄 다음에 한 줄 비우고, 그 아래부터 요약 작성
  let r = lastDataRow + 2; // 한 줄 띄운 후 시작

  // --- 총 수입 / 총 지출 / 총 잔액 (2줄) ---
  const sumHeaderRowIdx = r;
  const sumValueRowIdx = r + 1;

  const sumHeaderRow = ws.getRow(sumHeaderRowIdx);
  sumHeaderRow.getCell(5).value = '총 수입';
  sumHeaderRow.getCell(6).value = '총 지출';
  sumHeaderRow.getCell(7).value = '총 잔액';
  for (let c = 5; c <= 7; c++) {
    const cell = sumHeaderRow.getCell(c);
    cell.font = { name: '맑은 고딕', size: 10, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'center' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9E1F2' }, // 파란 헤더
    };
    applyThinBorder(cell);
  }

  const sumValueRow = ws.getRow(sumValueRowIdx);
  sumValueRow.getCell(5).value = totalIncome;
  sumValueRow.getCell(6).value = totalExpense;
  sumValueRow.getCell(7).value = totalBalance;
  for (let c = 5; c <= 7; c++) {
    const cell = sumValueRow.getCell(c);
    applyMoney(cell);
    applyThinBorder(cell);
  }

  // 다음 블록 시작 위치 계산
  r = sumValueRowIdx + 2; // 한 줄 비우고 다음 블록

  // --- 이월/연도별 수입·지출·총 잔액 (4줄) ---
  const rowPrev = ws.getRow(r);
  rowPrev.getCell(5).value = `${detail.prevYear}년 이월금`;
  ws.mergeCells(`E${r}:F${r}`);
  rowPrev.getCell(5).font = { name: '맑은 고딕', size: 10 };
  rowPrev.getCell(5).alignment = { horizontal: 'center', vertical: 'center' };
  applyThinBorder(rowPrev.getCell(5));
  applyThinBorder(rowPrev.getCell(6));
  rowPrev.getCell(7).value = toNum(detail.prevCarry);
  applyMoney(rowPrev.getCell(7));
  applyThinBorder(rowPrev.getCell(7));

  const rowYearIncome = ws.getRow(r + 1);
  rowYearIncome.getCell(5).value = `${detail.year}년 수입`;
  ws.mergeCells(`E${r + 1}:F${r + 1}`);
  rowYearIncome.getCell(5).font = { name: '맑은 고딕', size: 10 };
  rowYearIncome.getCell(5).alignment = {
    horizontal: 'center',
    vertical: 'center',
  };
  applyThinBorder(rowYearIncome.getCell(5));
  applyThinBorder(rowYearIncome.getCell(6));
  rowYearIncome.getCell(7).value = toNum(detail.yearIncome);
  applyMoney(rowYearIncome.getCell(7));
  applyThinBorder(rowYearIncome.getCell(7));

  const rowYearExpense = ws.getRow(r + 2);
  rowYearExpense.getCell(5).value = `${detail.year}년 지출`;
  ws.mergeCells(`E${r + 2}:F${r + 2}`);
  rowYearExpense.getCell(5).font = { name: '맑은 고딕', size: 10 };
  rowYearExpense.getCell(5).alignment = {
    horizontal: 'center',
    vertical: 'center',
  };
  applyThinBorder(rowYearExpense.getCell(5));
  applyThinBorder(rowYearExpense.getCell(6));
  rowYearExpense.getCell(7).value = toNum(detail.yearExpense);
  applyMoney(rowYearExpense.getCell(7));
  applyThinBorder(rowYearExpense.getCell(7));

  const rowYearBalance = ws.getRow(r + 3);
  rowYearBalance.getCell(5).value = `${detail.year}년 총 잔액`;
  ws.mergeCells(`E${r + 3}:F${r + 3}`);
  rowYearBalance.getCell(5).font = { name: '맑은 고딕', size: 10 };
  rowYearBalance.getCell(5).alignment = {
    horizontal: 'center',
    vertical: 'center',
  };
  applyThinBorder(rowYearBalance.getCell(5));
  applyThinBorder(rowYearBalance.getCell(6));
  rowYearBalance.getCell(7).value = toNum(detail.yearBalance);
  applyMoney(rowYearBalance.getCell(7));
  applyThinBorder(rowYearBalance.getCell(7));

  console.log(
    '[excel.js] 워크북 생성 완료',
    'rows:',
    rows.length,
    'summaryStartRow:',
    sumHeaderRowIdx
  );

  return wb;
}
