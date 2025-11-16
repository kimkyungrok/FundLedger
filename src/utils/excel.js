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

// 공용 테두리 프리셋
const B_THIN   = { style: 'thin',   color: { argb: 'FFBFBFBF' } };
const B_MEDIUM = { style: 'medium', color: { argb: 'FF808080' } };
const B_THICK  = { style: 'thick',  color: { argb: 'FF404040' } };

export async function buildLedgerExcel({
  rows = [],
  carry,
  pageSummary = {},
} = {}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('공금수불부', {
    pageSetup: { paperSize: 9, orientation: 'portrait' },
    properties: { defaultRowHeight: 18 },
    views: [{ state: 'frozen', ySplit: 2 }],
  });

  // 열너비
  COL_WIDTHS.forEach((w, i) => (ws.getColumn(i + 1).width = w));

  // ===== 제목 =====
  const lastColLetter = String.fromCharCode(64 + COL_WIDTHS.length); // 'G'
  ws.mergeCells(`A1:${lastColLetter}1`);
  const yearByData = (() => {
    if (rows.length && rows[0]?.date) {
      const d = new Date(rows[0].date);
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

  // ===== 데이터 작성 + 병합 구간 기록 =====
  let running = toNum(carry?.prevCarry, 0);
  const dataStartRow = 3; // 1:제목, 2:헤더, 3부터 데이터

  // YM 병합(A,B) 구간
  const ymSegments = []; // {start,end}
  let ymStart = null, prevYM = null;

  // YMD 병합(C) 구간
  const ymdSegments = []; // {start,end}
  let ymdStart = null, prevYMD = null;

  // 월 구간 테두리용
  const monthBands = []; // {start,end} for same Y-M
  let mbStart = null, prevMonthKey = null;

  // 일 구간 테두리용
  const dayBands = []; // {start,end} for same Y-M-D
  let dbStart = null, prevDayKey = null;

  rows.forEach((r, idx) => {
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

    // YM 병합
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

    // YMD 병합
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

    // 월 구간 테두리
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

    // 일 구간 테두리
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

  // 구간 마감
  const lastDataRow = ws.rowCount;
  if (ymStart != null && prevYM != null && lastDataRow > ymStart)
    ymSegments.push({ start: ymStart, end: lastDataRow });
  if (ymdStart != null && prevYMD != null && lastDataRow > ymdStart)
    ymdSegments.push({ start: ymdStart, end: lastDataRow });
  if (mbStart != null && prevMonthKey != null && lastDataRow >= mbStart)
    monthBands.push({ start: mbStart, end: lastDataRow });
  if (dbStart != null && prevDayKey != null && lastDataRow >= dbStart)
    dayBands.push({ start: dbStart, end: lastDataRow });

  // ===== 병합 적용 =====
  // A,B = 년·월 단위 병합
  for (const seg of ymSegments) {
    ['A', 'B'].forEach((col) => {
      ws.mergeCells(`${col}${seg.start}:${col}${seg.end}`);
      ws.getCell(`${col}${seg.start}`).alignment = { horizontal: 'center', vertical: 'middle' };
    });
  }
  // C = 년·월·일 단위 병합
  for (const seg of ymdSegments) {
    ws.mergeCells(`C${seg.start}:C${seg.end}`);
    ws.getCell(`C${seg.start}`).alignment = { horizontal: 'center', vertical: 'middle' };
  }

  // ===== 큰 틀 외곽 테두리(A2:G<lastDataRow>) =====
  if (lastDataRow >= 2) {
    // 위/아래 굵게
    for (let c = 1; c <= 7; c++) {
      const topCell = ws.getCell(lastDataRow >= 2 ? 2 : 2, c); // 헤더행
      topCell.border = {
        top: B_THICK, left: topCell.border?.left ?? B_THIN,
        right: topCell.border?.right ?? B_THIN, bottom: topCell.border?.bottom ?? B_THIN
      };
      const botCell = ws.getCell(lastDataRow, c);
      botCell.border = {
        top: botCell.border?.top ?? B_THIN, left: botCell.border?.left ?? B_THIN,
        right: botCell.border?.right ?? B_THIN, bottom: B_THICK
      };
    }
    // 좌/우 굵게
    for (let r = 2; r <= lastDataRow; r++) {
      const leftCell = ws.getCell(r, 1);
      leftCell.border = {
        top: leftCell.border?.top ?? B_THIN, left: B_THICK,
        right: leftCell.border?.right ?? B_THIN, bottom: leftCell.border?.bottom ?? B_THIN
      };
      const rightCell = ws.getCell(r, 7);
      rightCell.border = {
        top: rightCell.border?.top ?? B_THIN, left: rightCell.border?.left ?? B_THIN,
        right: B_THICK, bottom: rightCell.border?.bottom ?? B_THIN
      };
    }
  }

  // ===== 월 구간 두꺼운 가로선(thick) =====
  for (const band of monthBands) {
    // 시작 행 상단, 종료 행 하단
    for (let c = 1; c <= 7; c++) {
      const startCell = ws.getCell(band.start, c);
      startCell.border = {
        top: B_THICK,
        left: startCell.border?.left ?? B_THIN,
        right: startCell.border?.right ?? B_THIN,
        bottom: startCell.border?.bottom ?? B_THIN,
      };
      const endCell = ws.getCell(band.end, c);
      endCell.border = {
        top: endCell.border?.top ?? B_THIN,
        left: endCell.border?.left ?? B_THIN,
        right: endCell.border?.right ?? B_THIN,
        bottom: B_THICK,
      };
    }
  }

  // ===== 일 구간 중간 가로선(medium) =====
  for (const band of dayBands) {
    for (let c = 1; c <= 7; c++) {
      const startCell = ws.getCell(band.start, c);
      startCell.border = {
        top: B_MEDIUM,
        left: startCell.border?.left ?? B_THIN,
        right: startCell.border?.right ?? B_THIN,
        bottom: startCell.border?.bottom ?? B_THIN,
      };
      const endCell = ws.getCell(band.end, c);
      endCell.border = {
        top: endCell.border?.top ?? B_THIN,
        left: endCell.border?.left ?? B_THIN,
        right: endCell.border?.right ?? B_THIN,
        bottom: B_MEDIUM,
      };
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

  // 제목 정렬 보강
  ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}
