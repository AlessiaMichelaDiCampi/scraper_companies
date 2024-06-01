const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const readline = require('readline');

(async () => {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();

  await page.setRequestInterception(true);
  page.on('request', (request) => {
    if (
      request.resourceType() === 'document' ||
      request.resourceType() === 'script' ||
      request.resourceType() === 'xhr' ||
      request.resourceType() === 'fetch'
    ) {
      request.continue();
    } else {
      request.abort();
    }
  });

  await page.goto('https://servizi.anticorruzione.it/RicercaAttestazioniWebApp/#/', { timeout: 60000 });

  console.log("Seleziona la regione e le province manualmente, poi premi Invio qui.");
  await new Promise(resolve => process.stdin.once('data', () => resolve()));

  const selectRegionAndProvince = async (regionValue, provinceValues) => {
    await page.select('mat-form-field[colspan="9"] mat-select[formcontrolname="regione"]', regionValue);
    await page.waitForTimeout(1000);
    await page.select('mat-form-field[colspan="9"] mat-select[formcontrolname="provincia"]', provinceValues);
    await page.click('button[mat-flat-button][color="primary"]');
  };

  // Inserisci i valori della regione e delle province che desideri
  // await selectRegionAndProvince('3', ['CO', 'BS']);

  const extractNames = async () => {
    const names = await page.evaluate(() => {
      const nameElements = document.querySelectorAll('p.font-weight-bold.text-uppercase.mb-0');
      return Array.from(nameElements, element => element.textContent.trim());
    });
    return names;
  };

  let allNames = [];

  while (true) {
    try {
      const names = await extractNames();
      const newNames = names.filter(name => !allNames.includes(name));
      if (newNames.length > 0) {
        console.log(`Nomi trovati nella pagina ${await getCurrentPage()}:`, newNames);
        console.log(`Totale nomi finora: ${allNames.length + newNames.length}`);
      } else {
        console.log('Nessun nuovo nome trovato nella pagina', await getCurrentPage());
      }
      allNames = allNames.concat(newNames);


      const pageInfo = await page.$('.mat-paginator-range-label');
      if (pageInfo) {
        const pageInfoText = await page.evaluate(el => el.textContent.trim(), pageInfo);
        const lastPageInfo = pageInfoText.split('–')[1];
        const [currentPage, total] = lastPageInfo.split('of').map(item => item.trim());
        totalPages = parseInt(total);

        if (currentPage === total) {
          console.log('Ultima pagina raggiunta:', lastPageInfo);
          break;
        }
      } else {
        console.log('Informazioni sulla pagina non trovate.');
        break;
      }


      const nextButton = await page.$('.mat-paginator-navigation-next');
      if (nextButton) {
        await nextButton.click();
        await page.waitForTimeout(10000);
      } else {
        break;
      }
    } catch (error) {
      console.error('Si è verificato un errore durante l\'estrazione dei nomi:', error);
      break;
    }
  }

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  rl.question('Come desideri chiamare il file di output? (inclusa l\'estensione .xlsx): ', async (answer) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Nomi');
    allNames.forEach((name, index) => {
      worksheet.getCell(`A${index + 1}`).value = name;
    });

    const filePath = answer.trim();
    await workbook.xlsx.writeFile(filePath);
    console.log(`File Excel salvato con successo: ${filePath}`);

    await browser.close();
    rl.close();
  });

  async function getCurrentPage() {
    const pageInfo = await page.$('.mat-paginator-range-label');
    if (pageInfo) {
      const pageInfoText = await page.evaluate(el => el.textContent.trim(), pageInfo);
      const currentPage = pageInfoText.split('–')[0];
      return currentPage.trim();
    } else {
      return 'Informazioni sulla pagina non trovate.';
    }
  }
})();
