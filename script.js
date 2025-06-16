document.getElementById('upload').addEventListener('change', handleFile);
document.getElementById('concluir').addEventListener('click', exportToExcel);

let workbookGlobal;

function handleFile(e) {
  const reader = new FileReader();
  const file = e.target.files[0];
  
  reader.onload = ev => {
    const data = new Uint8Array(ev.target.result);
    workbookGlobal = XLSX.read(data, { type: 'array' });
    
    const html = XLSX.utils.sheet_to_html(workbookGlobal.Sheets[workbookGlobal.SheetNames[0]]);
    const container = document.getElementById('tabela');
    container.innerHTML = html;
    
    const table = container.querySelector('table');
// Remove todos os estilos inline da tabela e seus elementos internos
table.removeAttribute('style');

// Remove estilos dos elementos que o XLSX insere: td, th, tr, thead, tbody
table.querySelectorAll('td, th, tr, thead, tbody').forEach(el => {
  el.removeAttribute('style');
  el.className = ''; // Remove classes automáticas (se houver)
});
    // ✅ Aplica a sua classe com estilo personalizado
    table.classList.add('custom-table');
    
    // 🔁 Continua com a edição das células
    makeTableEditable(table);
  };
  
  reader.readAsArrayBuffer(file);
}

function makeTableEditable(table) {
  table.querySelectorAll('td').forEach(cell => {
    cell.addEventListener('click', () => showEditOptions(cell));
  });
}

function showEditOptions(cell) {
  const choice = confirm("Clique Ok para usar o Scanner, ou Cancelar para digitar manualmente.");
  if (choice) startScanner(cell);
  else {
    const val = prompt("Digite o valor:", cell.innerText);
    if (val !== null) cell.innerText = val;
  }
}

function startScanner(cell) {
  const popup = document.getElementById('scanner-popup');
  const video = document.getElementById('scanner-video');
  popup.style.display = 'block';
  
  const codeReader = new ZXing.BrowserBarcodeReader(); // só código de barras
  let resultText = "";
  
  codeReader.decodeFromVideoDevice(null, video, (result, err) => {
    if (result) {
      resultText = result.text;
    }
  });
  
  function stopScanner() {
    codeReader.reset();
    popup.style.display = 'none';
  }
  
  document.getElementById('scan-ok').onclick = () => {
    stopScanner();
    if (resultText) {
      const digitos = resultText.replace(/\D/g, '');
      if (digitos.length >= 6) {
        const ultimos6 = digitos.slice(-6);
        const cincoFinal = ultimos6.slice(0, 5);
        cell.innerText = cincoFinal;
      } else {
        alert("Código muito curto ou inválido.");
      }
    } else {
      alert("Nenhum código foi lido.");
    }
  };
  
  document.getElementById('scan-cancel').onclick = () => {
    stopScanner();
    const manual = prompt("Digite o valor:");
    if (manual !== null) cell.innerText = manual;
  };
}

function exportToExcel() {
  const table = document.querySelector('#tabela table');
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(table);
  XLSX.utils.book_append_sheet(wb, ws, "Planilha");
  
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = "planilha_editada.xlsx";
  a.innerText = "Clique aqui para baixar a planilha";
  
  const dl = document.getElementById('download');
  dl.innerHTML = '';
  dl.appendChild(a);
}