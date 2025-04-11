const xlsx = require('xlsx');
const axios = require('axios');
const cheerio = require('cheerio');

const inputFile = 'urlFormulario.xlsx';            // Arquivo de entrada com as URLs
const outputFile = 'urls_with_form.xlsx';   // Arquivo de saída com as URLs que possuem formulário

// Função que verifica se a URL possui um formulário (<form>)
async function checkForm(url) {
  try {
    const axiosConfig = {
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
      timeout: 10000  // timeout de 10 segundos
    };
    const response = await axios.get(url, axiosConfig);
    const html = response.data;
    const $ = cheerio.load(html);
    // Retorna true se houver pelo menos um elemento <form>
    return $('form').length > 0;
  } catch (error) {
    console.error(`Erro ao acessar ${url}: ${error.message}`);
    return false;
  }
}

// Função que processa todas as URLs do arquivo
async function processUrls() {
  // Lê o arquivo Excel
  const workbook = xlsx.readFile(inputFile);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Converte a planilha para um array (assumindo que as URLs estão na primeira coluna)
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  const urls = data
    .map(row => row[0])
    .filter(url => url); // remove linhas vazias

  const results = [];
  for (const url of urls) {
    console.log(`Verificando: ${url}`);
    const hasForm = await checkForm(url);
    if (hasForm) {
      console.log(`Formulário encontrado em: ${url}`);
      results.push([url]);
    } else {
      console.log(`Nenhum formulário em: ${url}`);
    }
  }

  // Cria uma nova planilha com as URLs que possuem formulário
  const newSheet = xlsx.utils.aoa_to_sheet([["URL"], ...results]);
  const newWorkbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(newWorkbook, newSheet, "URLs com Formulário");
  xlsx.writeFile(newWorkbook, outputFile);

  console.log(`Processamento concluído. Arquivo salvo como: ${outputFile}`);
}

// Executa a função principal
processUrls();
