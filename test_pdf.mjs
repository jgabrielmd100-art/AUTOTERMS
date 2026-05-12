// Quick test script for the PDF generation API
import fs from 'fs';

const payload = {
  filledHtml: `<h1 class="ql-align-center"><strong>TERMO DE TESTE</strong></h1><p><br></p><p><strong>EMPRESA:</strong> Simples Assessoria Contábil</p><p><strong>CNPJ:</strong> 12.345.678/0001-99</p><p><br></p><p>Este documento confirma que o Sr. <strong>João da Silva</strong>, portador do CPF 123.456.789-00, recebeu os equipamentos necessários.</p><p><br></p><p><strong>Data:</strong> 12/05/2026</p><p><br></p><p><strong>Assinatura:</strong> __________________________</p>`
};

console.log('Testing PDF generation...');

try {
  const response = await fetch('http://localhost:3001/api/generate-pdf', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });

  console.log(`Status: ${response.status}`);
  console.log(`Content-Type: ${response.headers.get('content-type')}`);

  if (response.ok) {
    const buffer = await response.arrayBuffer();
    fs.writeFileSync('test_output.pdf', Buffer.from(buffer));
    console.log(`✅ PDF saved! Size: ${buffer.byteLength} bytes`);
    console.log('File: test_output.pdf');
  } else {
    const text = await response.text();
    console.log(`❌ Error: ${text}`);
  }
} catch (err) {
  console.error('❌ Connection error:', err);
}
