// src/routes/entries.js
import express from 'express';

const router = express.Router();

/* ===== 유틸 ===== */
const isYmd = (v) => typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v);
const num   = (v) => Number(v || 0) || 0;

/** 어떤 입력이 와도 YYYY-MM-DD 문자열로 정규화 */
function toYmd(v) {
  try {
    if (typeof v === 'string') {
      if (isYmd(v)) return v;
      const d = new Date(v);
      if (!isNaN(d)) return d.toISOString().slice(0, 10);
      return '';
    }
    if (v instanceof Date) {
      if (!isNaN(v)) return v.toISOString().slice(0, 10);
      return '';
    }
    const d = new Date(v);
    if (!isNaN(d)) return d.toISOString().slice(0, 10);
    return '';
  } catch {
    return '';
  }
}

/** 문자열/Date 혼재 컬럼을 모두 커버하는 날짜범위 쿼리 */
function buildDateRangeQuery(start, end) {
  const hasStart = isYmd(start);
  const hasEnd   = isYmd(end);
  if (!hasStart && !hasEnd) return {};

  const strRange = {};
  const dateRange = {};
  if (hasStart) { strRange.$gte = start; dateRange.$gte = new Date(start); }
  if (hasEnd)   { strRange.$lte = end;   dateRange.$lte = new Date(end);   }

  return {
    $or: [
      { date: strRange },                                   // 문자열로 저장된 케이스
      { $and: [ { date: { $type: 9 } }, { date: dateRange } ] }, // Date 타입(9)
    ],
  };
}

/* 루트 → /entries */
router.get('/', (req, res) => res.redirect('/entries'));

/* 리스트 + 요약 렌더 */
router.get('/entries', async (req, res) => {
  try {
    const db = req.app.locals.db;

    // 쿼리 파라미터
    const { start, end, q, order } = req.query || {};
    const query = {};

    // 날짜 범위
    Object.assign(query, buildDateRangeQuery(start, end));

    // 적요 검색
    if (q && String(q).trim()) {
      query.desc = {
        $regex: String(q).trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
        $options: 'i',
      };
    }

    // 정렬: asc(최초 날짜순) / desc(최신 날짜순)
    const sortDir = (String(order).toLowerCase() === 'desc') ? -1 : 1;

    const rows = await db
      .collection('entries')
      .find(query)
      .sort({ date: sortDir, _id: sortDir }) // 타입 혼재시 보조키
      .toArray();

    // 설정(이월)
    const setting = (await db.collection('settings').findOne({ _id: 'carry' })) || {};
    const prevYear  = Number(setting.prevYear || new Date().getFullYear() - 1);
    const prevCarry = Number(setting.prevCarry || 0);

    // 합계
    const income  = rows.reduce((t, r) => t + num(r.income), 0);
    const expense = rows.reduce((t, r) => t + num(r.expense), 0);
    const balance = prevCarry + income - expense;

    // 연도별 합계(기준연도+1)
    const year = prevYear + 1;
    const yearStr = String(year) + '-';
    const rowsYear = rows.filter((r) => toYmd(r?.date).startsWith(yearStr));
    const yearIncome  = rowsYear.reduce((t, r) => t + num(r.income), 0);
    const yearExpense = rowsYear.reduce((t, r) => t + num(r.expense), 0);
    const yearBalance = prevCarry + yearIncome - yearExpense;

    const summary = {
      income,
      expense,
      balance,
      detail: { prevYear, year, prevCarry, yearIncome, yearExpense, yearBalance },
    };

    res.render('entries/index', {
      title: '입출금 내역',
      rows,
      summary,
      start: isYmd(start) ? start : '',
      end:   isYmd(end)   ? end   : '',
      q: q || '',
      order: (sortDir === -1 ? 'desc' : 'asc'),
    });
  } catch (e) {
    console.error('[GET /entries]', e);
    res.status(500).send('Server Error');
  }
});

/* 생성: 항상 YYYY-MM-DD 문자열로 저장 */
router.post('/entries', async (req, res) => {
  try {
    const db = req.app.locals.db;
    const { date, desc, income = 0, expense = 0 } = req.body || {};
    const ymd = toYmd(date);
    if (!ymd || !desc || !String(desc).trim()) {
      return res.status(400).json({ ok: false, error: 'invalid payload' });
    }
    const doc = {
      date: ymd,
      desc: String(desc).trim(),
      income: num(income),
      expense: num(expense),
      createdAt: new Date(),
      updatedAt: new Date(),
    };
    const r = await db.collection('entries').insertOne(doc);
    return res.json({ ok: true, id: String(r.insertedId) });
  } catch (e) {
    console.error('[POST /entries]', e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

/* 수정 */
router.put('/entries/:id', async (req, res) => {
  try {
    const db = req.app.locals.db;
    const { ObjectId } = req.app.locals;
    const { id } = req.params;

    const payload = {};
    if (req.body.date) {
      const ymd = toYmd(req.body.date);
      if (ymd) payload.date = ymd;
    }
    if (typeof req.body.desc === 'string') payload.desc = req.body.desc.trim();
    if (req.body.income != null)  payload.income  = num(req.body.income);
    if (req.body.expense != null) payload.expense = num(req.body.expense);

    if (!Object.keys(payload).length) {
      return res.status(400).json({ ok: false, error: 'empty update' });
    }
    payload.updatedAt = new Date();

    const r = await db
      .collection('entries')
      .updateOne({ _id: new ObjectId(id) }, { $set: payload });

    if (!r.matchedCount) return res.status(404).json({ ok: false, error: 'not found' });
    res.json({ ok: true });
  } catch (e) {
    console.error('[PUT /entries/:id]', e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

/* 삭제 */
router.delete('/entries/:id', async (req, res) => {
  try {
    const db = req.app.locals.db;
    const { ObjectId } = req.app.locals;
    const { id } = req.params;

    const r = await db.collection('entries').deleteOne({ _id: new ObjectId(id) });
    if (!r.deletedCount) return res.status(404).json({ ok: false, error: 'not found' });
    res.json({ ok: true });
  } catch (e) {
    console.error('[DELETE /entries/:id]', e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

/* 이월 설정 저장 */
router.put('/settings/carry', async (req, res) => {
  try {
    const db = req.app.locals.db;
    const prevYear  = Number(req.body.prevYear);
    const prevCarry = Number(req.body.prevCarry);
    if (!Number.isFinite(prevYear) || !Number.isFinite(prevCarry)) {
      return res.status(400).json({ ok: false, error: 'invalid payload' });
    }
    await db.collection('settings').updateOne(
      { _id: 'carry' },
      { $set: { prevYear, prevCarry, updatedAt: new Date() } },
      { upsert: true }
    );
    res.json({ ok: true });
  } catch (e) {
    console.error('[PUT /settings/carry]', e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

/* (옵션) 영수증 이미지 저장 */
router.post('/receipts', async (req, res) => {
  try {
    const db = req.app.locals.db;
    const { imageData, date, total, desc } = req.body || {};
    if (!imageData || !/^data:image\/(png|jpe?g);base64,/.test(imageData)) {
      return res.status(400).json({ ok: false, error: 'invalid image' });
    }
    const doc = {
      date: toYmd(date) || null,
      total: num(total),
      desc: typeof desc === 'string' ? desc.trim() : '',
      imageData,
      createdAt: new Date(),
    };
    const r = await db.collection('receipts').insertOne(doc);
    res.json({ ok: true, id: String(r.insertedId) });
  } catch (e) {
    console.error('[POST /receipts]', e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

export default router;
