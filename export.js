const express = require('express');
const router = express.Router();
const ExcelJS = require('exceljs');
const supabase = require('../db');
const { requireAuth, requireRole } = require('../auth');

const HEADER = [
  { header: 'Date', key: 'transaction_date', width: 14 },
  { header: 'Category', key: 'category', width: 14 },
  { header: 'Description', key: 'description', width: 40 },
  { header: 'Amount', key: 'amount', width: 14 },
  { header: 'Status', key: 'status', width: 12 },
  { header: 'Created At', key: 'created_at', width: 20 },
  { header: 'Paid At', key: 'paid_at', width: 20 }
];

function styleSheet(sheet) {
  sheet.columns = HEADER;
  sheet.getRow(1).font = { bold: true };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
}

router.get('/', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  const { site_id } = req.query;

  let query = supabase.from('cst_transactions').select('*').order('transaction_date', { ascending: false });
  if (site_id) query = query.eq('site_id', site_id);

  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Construction Site Transactions';
  workbook.created = new Date();

  const paid = data.filter(t => t.status === 'paid');
  const pending = data.filter(t => t.status === 'pending' || t.status === 'approved');
  const rejected = data.filter(t => t.status === 'rejected');

  const paidSheet = workbook.addWorksheet('Paid');
  styleSheet(paidSheet);
  paidSheet.addRows(paid);

  const pendingSheet = workbook.addWorksheet('Pending');
  styleSheet(pendingSheet);
  pendingSheet.addRows(pending);

  const rejectedSheet = workbook.addWorksheet('Rejected');
  styleSheet(rejectedSheet);
  rejectedSheet.addRows(rejected);

  const summarySheet = workbook.addWorksheet('Summary');
  summarySheet.columns = [{ header: 'Metric', key: 'metric', width: 24 }, { header: 'Value', key: 'value', width: 18 }];
  summarySheet.getRow(1).font = { bold: true };
  summarySheet.addRows([
    { metric: 'Total Paid', value: paid.reduce((s, t) => s + Number(t.amount), 0) },
    { metric: 'Total Pending', value: pending.reduce((s, t) => s + Number(t.amount), 0) },
    { metric: 'Total Rejected', value: rejected.reduce((s, t) => s + Number(t.amount), 0) },
    { metric: 'Count Paid', value: paid.length },
    { metric: 'Count Pending', value: pending.length }
  ]);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="site-transactions-${new Date().toISOString().slice(0, 10)}.xlsx"`);

  await workbook.xlsx.write(res);
  res.end();
});

module.exports = router;
