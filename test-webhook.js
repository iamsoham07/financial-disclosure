/**
 * Test script: generates a filled xlsx and sends it to the n8n webhook.
 * Run: node test-webhook.js
 */

const XlsxPopulate = require('xlsx-populate');
const fetch = require('node-fetch');
const FormData = require('form-data');

const WEBHOOK_URL = 'https://n8n.amicablerd.uk/webhook-test/910c3d20-faee-41b9-96d0-e5716c011c5b';

async function run() {
  // ── Build a minimal test xlsx with some recognisable values ──────────────
  const wb = await XlsxPopulate.fromBlankAsync();
  const ws = wb.sheet(0).name('Financial Disclosure');

  ws.cell('A1').value('Test Financial Disclosure');
  ws.cell('A2').value('Petitioner');
  ws.cell('B2').value('Alice');
  ws.cell('A3').value('Respondent');
  ws.cell('B3').value('Bob');
  ws.cell('A4').value('Case Number');
  ws.cell('B4').value('TEST-2024-001');
  ws.cell('A5').value('FMH Value');
  ws.cell('B5').value(500000);
  ws.cell('A6').value('Petitioner Bank');
  ws.cell('B6').value(25000);
  ws.cell('A7').value('Respondent Bank');
  ws.cell('B7').value(18000);
  ws.cell('A8').value('Petitioner Pension');
  ws.cell('B8').value(120000);
  ws.cell('A9').value('Respondent Pension');
  ws.cell('B9').value(85000);
  ws.cell('A10').value('Generated');
  ws.cell('B10').value(new Date().toISOString());

  const buffer = await wb.outputAsync();

  const hs_object_id = '999999999';
  const service = 'Assisted';
  const fileName = `Financial_Disclosure_${service}_Test_Bob_${new Date().toISOString().slice(0, 10)}.xlsx`;

  // ── Send as multipart form ────────────────────────────────────────────────
  const form = new FormData();
  form.append('file', buffer, {
    filename: fileName,
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  form.append('hs_object_id', hs_object_id);
  form.append('service', service);
  form.append('file_name', fileName);

  console.log(`Sending test payload to: ${WEBHOOK_URL}`);
  console.log(`File: ${fileName} (${buffer.length} bytes)`);

  const res = await fetch(WEBHOOK_URL, {
    method: 'POST',
    body: form,
    headers: form.getHeaders(),
  });

  const text = await res.text();
  console.log(`\nWebhook response status: ${res.status}`);
  console.log('Response body:', text);
}

run().catch(err => {
  console.error('Error:', err.message);
  process.exit(1);
});
