// src/routes/entries.js
import express from 'express';

const router = express.Router();

/* ===== 유틸 ===== */
const isYmd = (v) => typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v);
const num = (v) => Number(v || 0) || 0;

/**
 * 어떤 타입이든 'YYYY-MM-DD' 문자열로 정규화
 * - 문자열이면 앞부분의 YYYY-MM-DD를 사용
 * - Date 객체면 toISOString().slice(0,10)
 * - 그 외는 Date로 한번 감싸 보고 실패 시 ''
 */
const toYmd = (v) => {
  if (!v) return '';
  if (typeof v === 'string') {
    const m = v.match(/^(\d{4}-\d{2}-\d{2})/);
    if (m) return m[1];
    try {
      const d = new Date(v);
      if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    } catch (e) {
      return '';
    }
    return '';
  }
  if (v instanceof Date) {
    if (!isNaN(v.getTime())) return v.toISOString().slice(0, 10);
    return '';
  }
  try {
    const d = new Date(v);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  } catch (e) {
    return '';
  }
  return '';
};

// DB row의 date를 안전한 YYYY-MM-DD 문자열로
const dateStringOf = (row) => {
  if (!row) return '';
  return toYmd(row.date);
};

console.log('[entries router] 로드 완료');

/* ===== 홈 → /entries ===== */
router.get('/', (req, res) => {
  console.log('[GET /] redirect → /entries');
  res.redirect('/entries');
});

/* ===== 리스트 + 요약 렌더 ===== */
router.get('/entries', async (req, res) => {
  console.log('[GET /entries] query:', req.query);
  try {
    const db = req.app.locals.db;

    const { start, end, q, order } = req.query || {};
    const query = {};

    // 날짜 필터
    if (isYmd(start) || isYmd(end)) {
      const cond = {};
      if (isYmd(start)) cond.$gte = start;
      if (isYmd(end)) cond.$lte = end;
      query.date = cond;
    }

    // 적요 검색
    if (q && q.trim()) {
      const keyword = q.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      query.desc = { $regex: keyword, $options: 'i' };
    }

    const rowsRaw = await db.collection('entries').find(query).toArray();

    const rows = rowsRaw.map((r) => ({
      ...r,
      date: toYmd(r.date),
    }));

    const sortDir = order === 'desc' ? -1 : 1;
    rows.sort((a, b) => {
      const da = a.date || '';
      const db_ = b.date || '';
      if (!da && !db_) return 0;
      if (!da) return -sortDir;
      if (!db_) return sortDir;
      if (da < db_) return -sortDir;
      if (da > db_) return sortDir;
      const ida = String(a._id || '');
      const idb = String(b._id || '');
      if (ida < idb) return -sortDir;
      if (ida > idb) return sortDir;
      return 0;
    });

    const setting =
      (await db.collection('settings').findOne({ _id: 'carry' })) || {};
    const prevYear = Number(setting.prevYear || new Date().getFullYear() - 1);
    const prevCarry = Number(setting.prevCarry || 0);

    const income = rows.reduce((t, r) => t + num(r.income), 0);
    const expense = rows.reduce((t, r) => t + num(r.expense), 0);
    const balance = prevCarry + income - expense;

    const year = prevYear + 1;
    const yearPrefix = String(year);
    const rowsYear = rows.filter((r) => {
      const d = r.date || '';
      return typeof d === 'string' && d.startsWith(yearPrefix);
    });

    const yearIncome = rowsYear.reduce((t, r) => t + num(r.income), 0);
    const yearExpense = rowsYear.reduce((t, r) => t + num(r.expense), 0);
    const yearBalance = prevCarry + yearIncome - yearExpense;

    const summary = {
      income,
      expense,
      balance,
      detail: {
        prevYear,
        year,
        prevCarry,
        yearIncome,
        yearExpense,
        yearBalance,
      },
    };

    console.log('[GET /entries] rows:', rows.length, 'summary:', summary);

    res.render('entries/index', {
      title: '입출금 내역',
      rows,
      summary,
      start: isYmd(start) ? start : '',
      end: isYmd(end) ? end : '',
      q: q || '',
      order: order === 'desc' ? 'desc' : 'asc',
    });
  } catch (e) {
    console.error('[GET /entries] ERROR', e);
    res.status(500).send('Server Error');
  }
});

/* ===== 생성 ===== */
router.post('/entries', async (req, res) => {
  console.log('[POST /entries] body:', req.body);
  try {
    const db = req.app.locals.db;
    const { date, desc, income = 0, expense = 0 } = req.body || {};

    const ymd = toYmd(date);
    if (!isYmd(ymd) || !desc || !String(desc).trim()) {
      console.warn('[POST /entries] invalid payload:', { date, desc });
      return res.status(400).json({ ok: false, error: 'invalid payload' });
    }

    const now = new Date();
    const doc = {
      date: ymd,
      desc: String(desc).trim(),
      income: num(income),
      expense: num(expense),
      createdAt: now,
      updatedAt: now,
    };

    const r = await db.collection('entries').insertOne(doc);
    console.log('[POST /entries] insertedId:', r.insertedId);
    return res.json({ ok: true, id: String(r.insertedId) });
  } catch (e) {
    console.error('[POST /entries] ERROR', e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

/* ===== 수정 ===== */
router.put('/entries/:id', async (req, res) => {
  console.log('[PUT /entries/:id] params:', req.params, 'body:', req.body);
  try {
    const db = req.app.locals.db;
    const { ObjectId } = req.app.locals;
    const { id } = req.params;

    const payload = {};

    if (req.body.date) {
      const ymd = toYmd(req.body.date);
      if (isYmd(ymd)) payload.date = ymd;
    }
    if (typeof req.body.desc === 'string') {
      payload.desc = req.body.desc.trim();
    }
    if (req.body.income != null) {
      payload.income = num(req.body.income);
    }
    if (req.body.expense != null) {
      payload.expense = num(req.body.expense);
    }

    if (!Object.keys(payload).length) {
      console.warn('[PUT /entries/:id] empty update');
      return res.status(400).json({ ok: false, error: 'empty update' });
    }
    payload.updatedAt = new Date();

    const r = await db
      .collection('entries')
      .updateOne({ _id: new ObjectId(id) }, { $set: payload });

    if (!r.matchedCount) {
      console.warn('[PUT /entries/:id] not found:', id);
      return res.status(404).json({ ok: false, error: 'not found' });
    }
    console.log(
      '[PUT /entries/:id] matched:',
      r.matchedCount,
      'modified:',
      r.modifiedCount
    );
    res.json({ ok: true });
  } catch (e) {
    console.error('[PUT /entries/:id] ERROR', e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

/* ===== 삭제 ===== */
router.delete('/entries/:id', async (req, res) => {
  console.log('[DELETE /entries/:id] params:', req.params);
  try {
    const db = req.app.locals.db;
    const { ObjectId } = req.app.locals;
    const { id } = req.params;

    const r = await db
      .collection('entries')
      .deleteOne({ _id: new ObjectId(id) });
    if (!r.deletedCount) {
      console.warn('[DELETE /entries/:id] not found:', id);
      return res.status(404).json({ ok: false, error: 'not found' });
    }
    console.log('[DELETE /entries/:id] deletedCount:', r.deletedCount);
    res.json({ ok: true });
  } catch (e) {
    console.error('[DELETE /entries/:id] ERROR', e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

/* ===== 이월 설정 저장 ===== */
router.put('/settings/carry', async (req, res) => {
  console.log('[PUT /settings/carry] body:', req.body);
  try {
    const db = req.app.locals.db;
    const prevYear = Number(req.body.prevYear);
    const prevCarry = Number(req.body.prevCarry);
    if (!Number.isFinite(prevYear) || !Number.isFinite(prevCarry)) {
      console.warn('[PUT /settings/carry] invalid payload:', req.body);
      return res.status(400).json({ ok: false, error: 'invalid payload' });
    }
    await db.collection('settings').updateOne(
      { _id: 'carry' },
      { $set: { prevYear, prevCarry, updatedAt: new Date() } },
      { upsert: true }
    );
    console.log('[PUT /settings/carry] upsert OK:', { prevYear, prevCarry });
    res.json({ ok: true });
  } catch (e) {
    console.error('[PUT /settings/carry] ERROR', e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

/* ===== 영수증 이미지 저장 (선택) ===== */
router.post('/receipts', async (req, res) => {
  console.log('[POST /receipts] body keys:', Object.keys(req.body || {}));
  try {
    const db = req.app.locals.db;
    const { imageData, date, total, desc } = req.body || {};
    if (!imageData || !/^data:image\/(png|jpe?g);base64,/.test(imageData)) {
      console.warn('[POST /receipts] invalid image header');
      return res.status(400).json({ ok: false, error: 'invalid image' });
    }
    const doc = {
      date: isYmd(toYmd(date)) ? toYmd(date) : null,
      total: num(total),
      desc: typeof desc === 'string' ? desc.trim() : '',
      imageData,
      createdAt: new Date(),
    };
    const r = await db.collection('receipts').insertOne(doc);
    console.log('[POST /receipts] insertedId:', r.insertedId);
    res.json({ ok: true, id: String(r.insertedId) });
  } catch (e) {
    console.error('[POST /receipts] ERROR', e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

/* ===== 엑셀 내보내기 (동적 import) ===== */

let createWorkbookFromEntriesFn = null;

router.get('/entries/export', async (req, res, next) => {
  console.log('[GET /entries/export] query:', req.query);
  try {
    const db = req.app.locals.db;
    const { start, end, q } = req.query || {};
    const orderParam = req.query.order || req.query.sort || 'asc';
    const sortDir = orderParam === 'desc' ? -1 : 1;

    // 1) excel.js 동적 import + 디버깅
    if (!createWorkbookFromEntriesFn) {
      try {
        console.log('[GET /entries/export] excel.js 동적 import 시도');
        const mod = await import('../utils/excel.js');
        console.log(
          '[GET /entries/export] excel.js import 성공, export keys:',
          Object.keys(mod)
        );

        if (typeof mod.createWorkbookFromEntries !== 'function') {
          console.error(
            '[GET /entries/export] createWorkbookFromEntries 함수가 아님:',
            typeof mod.createWorkbookFromEntries
          );
          return res
            .status(500)
            .send(
              '엑셀 모듈에 createWorkbookFromEntries 함수가 없습니다. export 형식을 확인하세요.'
            );
        }
        createWorkbookFromEntriesFn = mod.createWorkbookFromEntries;
      } catch (impErr) {
        console.error(
          '[GET /entries/export] excel.js import 실패:',
          impErr
        );
        return res
          .status(500)
          .send(
            '엑셀 모듈 로드 중 오류: ' +
              (impErr?.message || String(impErr))
          );
      }
    }

    const query = {};
    if (isYmd(start) || isYmd(end)) {
      const cond = {};
      if (isYmd(start)) cond.$gte = start;
      if (isYmd(end)) cond.$lte = end;
      query.date = cond;
    }
    if (q && q.trim()) {
      const keyword = q.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      query.desc = { $regex: keyword, $options: 'i' };
    }

    console.log(
      '[GET /entries/export] Mongo query:',
      JSON.stringify(query),
      'sortDir:',
      sortDir
    );

    const rowsRaw = await db
      .collection('entries')
      .find(query)
      .sort({ date: sortDir, _id: sortDir })
      .toArray();

    const rows = rowsRaw.map((r) => ({
      ...r,
      date: dateStringOf(r),
      income: num(r.income),
      expense: num(r.expense),
    }));

    const setting =
      (await db.collection('settings').findOne({ _id: 'carry' })) || {};
    const prevYear = Number(
      setting.prevYear || new Date().getFullYear() - 1
    );
    const prevCarry = Number(setting.prevCarry || 0);

    const income = rows.reduce((t, r) => t + num(r.income), 0);
    const expense = rows.reduce((t, r) => t + num(r.expense), 0);
    const balance = prevCarry + income - expense;

    const year = prevYear + 1;
    const yearPrefix = String(year);
    const rowsYear = rows.filter((r) => {
      const d = r.date || '';
      return typeof d === 'string' && d.startsWith(yearPrefix);
    });

    const yearIncome = rowsYear.reduce((t, r) => t + num(r.income), 0);
    const yearExpense = rowsYear.reduce((t, r) => t + num(r.expense), 0);
    const yearBalance = prevCarry + yearIncome - yearExpense;

    const summary = {
      income,
      expense,
      balance,
      detail: {
        prevYear,
        year,
        prevCarry,
        yearIncome,
        yearExpense,
        yearBalance,
      },
    };

    console.log('[GET /entries/export] summary/rows 계산 완료', {
      rowsLen: rows.length,
      income,
      expense,
      balance,
      orderParam,
    });

    let workbook;
    try {
      workbook = await createWorkbookFromEntriesFn(summary, rows, {
        start: isYmd(start) ? start : '',
        end: isYmd(end) ? end : '',
        q: q || '',
        sort: orderParam === 'desc' ? 'desc' : 'asc',
      });
    } catch (excelErr) {
      console.error(
        '[GET /entries/export] createWorkbookFromEntries 호출 에러:',
        excelErr
      );
      return res
        .status(500)
        .send(
          '엑셀 워크북 생성 중 오류: ' +
            (excelErr?.message || String(excelErr))
        );
    }

    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    const filename = encodeURIComponent(`공금수불부_${y}${m}${d}.xlsx`);

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${filename}"; filename*=UTF-8''${filename}`
    );

    try {
      console.log('[GET /entries/export] workbook.xlsx.write 시작');
      await workbook.xlsx.write(res);
      console.log('[GET /entries/export] workbook.xlsx.write 완료');
      res.end();
    } catch (writeErr) {
      console.error(
        '[GET /entries/export] workbook.xlsx.write 에러:',
        writeErr
      );
      if (!res.headersSent) {
        return res
          .status(500)
          .send(
            '엑셀 파일 전송 중 오류: ' +
              (writeErr?.message || String(writeErr))
          );
      }
      try {
        res.end();
      } catch (_) {}
    }
  } catch (err) {
    console.error('[GET /entries/export] 외부 try/catch 에러:', err);
    next(err);
  }
});

export default router;
