// src/utils/excel.js
import ExcelJS from 'exceljs';

const COL_WIDTHS = [5, 4, 4, 20.13, 13.13, 13.13, 16.13];

function sanitizeXmlText(v) {
  if (v == null) return '';
  let o = '';
  for (const ch of String(v)) {
    const cp = ch.codePointAt(0);
    if (
      (cp >= 0x20 && cp <= 0xD7FF) ||
      (cp >= 0xE000 && cp <= 0xFFFD) ||
      (cp >= 0x10000 && cp <= 0x10FFFF) ||
      cp === 0x09 || cp === 0x0A || cp === 0x0D
    ) o += ch;
  }
  return o;
}
const toNum = (v, d = 0) => (Number.isFinite(Number(v)) ? Number(v) : d);

// 통일 테두리 프리셋
const B_THIN   = { style: 'thin',   color: { argb: 'FFBFBFBF' } };
const B_MEDIUM = { style: 'medium', color: { argb: 'FF808080' } };
const B_THICK  = { style: 'thick',  color: { argb: 'FF404040' } };

// 범위 전체에 동일 테두리 적용
function setRangeBorder(ws, r1, c1, r2, c2, { top, right, bottom, left }) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      // 각 셀 기준으로 해당 변만 지정 (겹치는 곳도 일관되게 지정)
      const cell = ws.getCell(r, c);
      const border = {};
      if (r === r1 && top) border.top = top;
      if (r === r2 && bottom) border.bottom = bottom;
      if (c === c1 && left) border.left = left;
      if (c === c2 && right) border.right = right;

      // 다른 변은 유지하려면 얇은 선으로 통일
      if (!border.top) border.top = B_THIN;
      if (!border.right) border.right = B_THIN;
      if (!border.bottom) border.bottom = B_THIN;
      if (!border.left) border.left = B_THIN;

      cell.border = border;
    }
  }
}

export async function buildLedgerExcel({
  rows = [],
  carry,
  pageSummary = {},
} = {}) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('공금수불부', {
    pageSetup: { paperSize: 9, orientation: 'portrait' },
    properties: { defaultRowHeight: 18 },
    views: [{ state: 'frozen', ySplit: 2 }],
  });

  // 열너비(엑셀 문자기반 단위)
  COL_WIDTHS.forEach((w, i) => (ws.getColumn(i + 1).width = w));

  // ===== 제목 =====
  const lastColLetter = String.fromCharCode(64 + COL_WIDTHS.length); // 'G'
  ws.mergeCells(`A1:${lastColLetter}1`);
  const yearByData = (() => {
    if (safeRows.length && safeRows[0]?.date) {
      const d = new Date(safeRows[0].date);
      if (!isNaN(d)) return d.getFullYear();
    }
    return new Date().getFullYear();
  })();
  const title = ws.getCell('A1');
  title.value = `${yearByData}년 공금 수불부`;
  title.font = { name: 'Malgun Gothic', size: 16, bold: true };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 24;

  // ===== 헤더 =====
  const header = ws.addRow(['년', '월', '일', '적요', '수입', '지출', '잔액']);
  header.eachCell((c) => {
    c.font = { name: 'Malgun Gothic', size: 11, bold: true };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } };
    c.border = { top: B_THIN, left: B_THIN, right: B_THIN, bottom: B_THIN };
  });
  ws.getRow(header.number).height = 20;

  // ===== 회계 포맷 =====
  const ACCOUNTING =
    '_-[$₩-ko-KR]* #,##0_-;_-[$₩-ko-KR]* #,##0_-;_-[$₩-ko-KR]* "-"_-;_-@_-';
  [5, 6, 7].forEach((col) => {
    ws.getColumn(col).numFmt = ACCOUNTING;
    ws.getColumn(col).alignment = { horizontal: 'right', vertical: 'middle' };
  });

  // ===== 데이터 작성 + 구간 계산 =====
  let running = toNum(carry?.prevCarry, 0);
  const dataStartRow = 3; // 1:제목, 2:헤더

  // 병합/테두리 구간
  const ymSegments = [];   // A,B 병합용 (년-월)
  let ymStart = null, prevYM = null;

  const ymdSegments = [];  // C 병합용 (년-월-일)
  let ymdStart = null, prevYMD = null;

  const monthBands = [];   // 월 테두리(thick)
  let mbStart = null, prevMonthKey = null;

  const dayBands = [];     // 일 테두리(medium)
  let dbStart = null, prevDayKey = null;

  safeRows.forEach((r, idx) => {
    const d = new Date(r?.date);
    const ok = Number.isFinite(d.getTime());
    const y = ok ? d.getFullYear() : null;
    const m = ok ? d.getMonth() + 1 : null;
    const day = ok ? d.getDate() : null;

    const inc = toNum(r?.income, 0);
    const exp = toNum(r?.expense, 0);
    running += inc - exp;

    const row = ws.addRow([
      y ?? '',
      m ?? '',
      day ?? '',
      sanitizeXmlText(r?.desc ?? ''),
      inc,
      exp,
      running,
    ]);
    row.eachCell((c, i) => {
      c.font = { name: 'Malgun Gothic', size: 10 };
      if (i <= 4) c.alignment = { horizontal: 'center', vertical: 'middle' };
      else c.alignment = { horizontal: 'right', vertical: 'middle' };
      c.border = { top: B_THIN, left: B_THIN, right: B_THIN, bottom: B_THIN };
    });

    const rowNo = dataStartRow + idx;

    // YM (년-월) 키
    const ymKey = y != null && m != null ? `${y}-${m}` : null;
    if (ymKey && prevYM === ymKey) {
      // 진행
    } else {
      if (ymStart != null && prevYM != null) {
        const end = rowNo - 1;
        if (end > ymStart) ymSegments.push({ start: ymStart, end });
      }
      ymStart = ymKey ? rowNo : null;
      prevYM = ymKey;
    }

    // YMD (년-월-일) 키
    const ymdKey = y != null && m != null && day != null ? `${y}-${m}-${day}` : null;
    if (ymdKey && prevYMD === ymdKey) {
      // 진행
    } else {
      if (ymdStart != null && prevYMD != null) {
        const end = rowNo - 1;
        if (end > ymdStart) ymdSegments.push({ start: ymdStart, end });
      }
      ymdStart = ymdKey ? rowNo : null;
      prevYMD = ymdKey;
    }

    // 월 테두리 밴드
    if (ymKey && prevMonthKey === ymKey) {
      // 진행
    } else {
      if (mbStart != null && prevMonthKey != null) {
        const end = rowNo - 1;
        if (end >= mbStart) monthBands.push({ start: mbStart, end });
      }
      mbStart = ymKey ? rowNo : null;
      prevMonthKey = ymKey;
    }

    // 일 테두리 밴드
    if (ymdKey && prevDayKey === ymdKey) {
      // 진행
    } else {
      if (dbStart != null && prevDayKey != null) {
        const end = rowNo - 1;
        if (end >= dbStart) dayBands.push({ start: dbStart, end });
      }
      dbStart = ymdKey ? rowNo : null;
      prevDayKey = ymdKey;
    }
  });

  const lastDataRow = ws.rowCount; // 데이터의 실제 마지막 행 번호

  // 구간 마감 (단일 행이면 병합/굵은선 생략)
  if (ymStart != null && prevYM != null && lastDataRow > ymStart)
    ymSegments.push({ start: ymStart, end: lastDataRow });
  if (ymdStart != null && prevYMD != null && lastDataRow > ymdStart)
    ymdSegments.push({ start: ymdStart, end: lastDataRow });
  if (mbStart != null && prevMonthKey != null && lastDataRow > mbStart - 1)
    monthBands.push({ start: mbStart, end: lastDataRow });
  if (dbStart != null && prevDayKey != null && lastDataRow > dbStart - 1)
    dayBands.push({ start: dbStart, end: lastDataRow });

  // ===== 병합 적용 =====
  for (const seg of ymSegments) {
    if (seg.end > seg.start) {
      ['A', 'B'].forEach((col) => {
        ws.mergeCells(`${col}${seg.start}:${col}${seg.end}`);
        ws.getCell(`${col}${seg.start}`).alignment = { horizontal: 'center', vertical: 'middle' };
      });
    }
  }
  for (const seg of ymdSegments) {
    if (seg.end > seg.start) {
      ws.mergeCells(`C${seg.start}:C${seg.end}`);
      ws.getCell(`C${seg.start}`).alignment = { horizontal: 'center', vertical: 'middle' };
    }
  }

  // ===== 큰 틀 외곽선(A2:G<lastDataRow>) - 굵게 =====
  if (lastDataRow >= 2) {
    setRangeBorder(ws, 2, 1, lastDataRow, 7, {
      top: B_THICK, right: B_THICK, bottom: B_THICK, left: B_THICK,
    });
  }

  // ===== 월/일 구간 구분선 =====
  for (const b of monthBands) {
    if (b.end >= b.start) {
      // 시작 상단 / 종료 하단 thick
      setRangeBorder(ws, b.start, 1, b.start, 7, { top: B_THICK, right: null, bottom: null, left: null });
      setRangeBorder(ws, b.end,   1, b.end,   7, { top: null, right: null, bottom: B_THICK, left: null });
    }
  }
  for (const b of dayBands) {
    if (b.end >= b.start) {
      // 시작 상단 / 종료 하단 medium
      setRangeBorder(ws, b.start, 1, b.start, 7, { top: B_MEDIUM, right: null, bottom: null, left: null });
      setRangeBorder(ws, b.end,   1, b.end,   7, { top: null, right: null, bottom: B_MEDIUM, left: null });
    }
  }

  // ===== 합계 요약표 =====
  const sumIncome = toNum(pageSummary?.sumIncome, 0);
  const sumExpense = toNum(pageSummary?.sumExpense, 0);
  const sumBalance = toNum(pageSummary?.sumBalance, running);
  ws.addRow([]);

  const headRow = ws.addRow([null, null, null, null, '총 수입', '총 지출', '총 잔액']);
  ['E', 'F', 'G'].forEach((col) => {
    const c = ws.getCell(`${col}${headRow.number}`);
    c.font = { name: 'Malgun Gothic', size: 11, bold: true };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
    c.border = { top: B_THIN, left: B_THIN, right: B_THIN, bottom: B_THIN };
  });

  const valRow = ws.addRow([null, null, null, null, sumIncome, sumExpense, sumBalance]);
  ['E', 'F', 'G'].forEach((col) => {
    const c = ws.getCell(`${col}${valRow.number}`);
    c.font = { name: 'Malgun Gothic', size: 10, bold: true };
    c.alignment = { horizontal: 'right', vertical: 'middle' };
    c.numFmt = ACCOUNTING;
    c.border = { top: B_THIN, left: B_THIN, right: B_THIN, bottom: B_THIN };
  });

  ws.addRow([]);
  const prevYear = Number(carry?.prevYear) || yearByData - 1;
  const kvRows = [
    [`${prevYear}년 이월금`, toNum(carry?.prevCarry, 0)],
    [`${yearByData}년 수입`, sumIncome],
    [`${yearByData}년 지출`, sumExpense],
    [`${yearByData}년 총 잔액`, sumBalance],
  ];
  for (const [label, value] of kvRows) {
    const r = ws.addRow([null, null, null, null, label, null, value]);
    ws.mergeCells(`E${r.number}:F${r.number}`);
    const l = ws.getCell(`E${r.number}`);
    const v = ws.getCell(`G${r.number}`);
    l.font = { name: 'Malgun Gothic', size: 10, bold: true };
    v.font = { name: 'Malgun Gothic', size: 10, bold: true };
    l.alignment = { horizontal: 'center', vertical: 'middle' };
    v.alignment = { horizontal: 'right', vertical: 'middle' };
    v.numFmt = ACCOUNTING;
    l.border = { top: B_THIN, left: B_THIN, right: B_THIN, bottom: B_THIN };
    v.border = { top: B_THIN, left: B_THIN, right: B_THIN, bottom: B_THIN };
  }

  // 제목 정렬 재보강
  ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };

  const buf = await wb.xlsx.writeBuffer();
  // writeBuffer()가 이미 Uint8Array를 반환 → Buffer로 감싸서 반환
  return Buffer.from(buf);
}
