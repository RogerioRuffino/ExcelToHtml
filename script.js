const inputExcel = document.getElementById('input-excel');
const selectSheet = document.getElementById('select-sheet');
const tablesContainer = document.getElementById('tables');

inputExcel.addEventListener('change', () => {
  const files = inputExcel.files;
  if(files.length === 0) return;
  const file = files[0];

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = e.target.result;
    const workbook = XLSX.read(data, {type: 'binary'});

    // Popula o select com as opções de abas
    selectSheet.innerHTML = '';
    workbook.SheetNames.forEach(sheetName => {
      const option = document.createElement('option');
      option.textContent = sheetName;
      selectSheet.appendChild(option);
    });

    // Exibe os dados da primeira aba ao carregar o arquivo
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    exibeDados(sheet);

    // Altera os dados exibidos na tabela ao selecionar uma nova aba
    selectSheet.addEventListener('change', () => {
      const sheetName = selectSheet.value;
      const sheet = workbook.Sheets[sheetName];
      exibeDados(sheet);
    });
  };
  reader.readAsBinaryString(file);
});

function exibeDados(sheet) {
  const table = XLSX.utils.sheet_to_html(sheet);
  tablesContainer.innerHTML = table;
  
  const tableElem = tablesContainer.querySelector('table');
  tableElem.classList.add('table');
  
  const headerRow = tableElem.querySelector('tr');

 
const numCols = headerRow.children.length;
if(numCols < 5) {
tableElem.classList.add('less-than-5-cols');
}
else {
tableElem.classList.remove('less-than-5-cols');
}
}