const { chromium } = require("playwright");
const XLSX = require("xlsx");
const fs = require("fs-extra");

const INPUT_FILE = "./magalu.xlsx";
const OUTPUT_FILE = "./cotacoes.xlsx";

// ---------- UTILS ----------
function normalizeCEP(cep) {
    const d = String(cep || "").replace(/\D/g, "");
    return d.length === 8 ? `${d.slice(0, 5)}-${d.slice(5)}` : d;
}

function sanitize(text) {
    return String(text || "").replace(/\s+/g, " ").trim();
}

function nowLabel() {
    return new Date().toLocaleString("pt-BR");
}

// ---------- EXCEL ----------
function readExcel(file) {
    const wb = XLSX.readFile(file);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet);
}

function updateExcel(results) {
    let wb;

    if (fs.existsSync(OUTPUT_FILE)) {
        wb = XLSX.readFile(OUTPUT_FILE);
    } else {
        wb = XLSX.utils.book_new();
    }

    const sheetName = "cota";
    let sheet = wb.Sheets[sheetName];

    if (!sheet) {
        sheet = XLSX.utils.json_to_sheet([]);
        XLSX.utils.book_append_sheet(wb, sheet, sheetName);
    }

    let data = XLSX.utils.sheet_to_json(sheet);
    const columnName = nowLabel();

    const map = {};
    data.forEach((row, index) => {
        map[sanitize(row.Destino)] = index;
    });

    results.forEach(r => {
        const destino = sanitize(r.destino);
        if (!destino) return;

        if (map[destino] !== undefined) {
            data[map[destino]][columnName] = r.prazo || "";
        } else {
            data.push({
                Destino: destino,
                [columnName]: r.prazo || ""
            });
        }
    });

    const newSheet = XLSX.utils.json_to_sheet(data);
    wb.Sheets[sheetName] = newSheet;

    XLSX.writeFile(wb, OUTPUT_FILE);
}

// ---------- PLAYWRIGHT ----------
async function openShipping(page) {
    const selectors = [
        'text=Alterar',
        'text=Calcular frete e prazo'
    ];

    for (let sel of selectors) {
        try {
            const el = page.locator(sel).first();
            await el.click({ timeout: 4000 });
            await page.waitForSelector('[data-testid="zipcode-input"]');
            return;
        } catch { }
    }

    throw new Error("Não abriu frete");
}

async function fillCEP(page, cep) {
    const input = page.locator('[data-testid="zipcode-input"]');
    await input.fill(cep);

    const btn = page.locator('[data-testid="confirm-zipcode-button"]');
    await btn.click();

    await page.waitForTimeout(2000);
}

async function getShipping(page) {
    const items = page.locator('[data-testid="shipping-item"]');
    const count = await items.count();

    if (!count) return null;

    const text = sanitize(await items.first().innerText());

    const price = text.match(/R\$\s?[\d\.,]+/)?.[0] || "";
    const prazo = text.match(/Receba até.*?\d+/)?.[0] || "";

    return { price, prazo };
}

async function processRow(browser, row) {
    const page = await browser.newPage();

    await page.goto(row.LINK);
    await page.waitForTimeout(2000);

    const ceps = [row.CEP1, row.CEP2, row.CEP3]
        .map(normalizeCEP)
        .filter(Boolean);

    for (let cep of ceps) {
        try {
            await openShipping(page);
            await fillCEP(page, cep);

            const data = await getShipping(page);

            if (data) {
                return {
                    destino: row.Destino,
                    cep,
                    ...data,
                    status: "OK"
                };
            }

        } catch (e) {
            console.log("Erro CEP:", cep);
        }
    }

    return {
        destino: row.Destino,
        status: "ERRO"
    };
}

// ---------- MAIN ----------
(async () => {
    const data = readExcel(INPUT_FILE);
    const browser = await chromium.launch({ headless: true });

    const results = [];

    for (let row of data) {
        const result = await processRow(browser, row);
        results.push(result);
        console.log(result);
    }

    await browser.close();

    updateExcel(results);

    fs.writeJsonSync("resultado.json", results, { spaces: 2 });

    console.log("Finalizado!");
})();