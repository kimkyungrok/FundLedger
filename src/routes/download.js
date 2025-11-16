// routes/download.js
app.get('/ledger.xlsx', async (req, res) => {
  const buffer = await buildLedgerExcel({ rows, carry, yearSummary, pageSummary });
  res.set({
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': 'attachment; filename="입출금내역.xlsx"',
    'Content-Length': buffer.length,
  });
  res.end(buffer); // res.send(buffer)도 가능하지만 end가 안전
});
