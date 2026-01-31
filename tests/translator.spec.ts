import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';
import * as XLSX from 'xlsx';

// --- CONFIGURATION ---
const FILE_NAME = 'Test_cases.xlsx';

// Storage for results (Key: TC ID, Value: Result)
const resultsMap = new Map<string, { actual: string; status: string }>();

// --- HELPER: Safely get text from ANY Excel cell ---
function getSafeCellText(value: any): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object' && 'richText' in value) {
    return value.richText.map((p: any) => p.text).join('').trim();
  }
  if (typeof value === 'object' && 'text' in value) {
    return value.text.toString().trim();
  }
  return String(value).trim();
}

// --- 1. SYNC LOADER ---
function loadTestsSync() {
  const filePath = path.join(process.cwd(), FILE_NAME);
  if (!fs.existsSync(filePath)) throw new Error(`File not found: ${filePath}`);

  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false }) as any[][];

  let headerRowIndex = -1;
  // Find header row (dynamically finds the row with "TC ID")
  for (let i = 0; i < rawData.length; i++) {
    if (rawData[i]?.some((c: any) => String(c).trim() === 'TC ID')) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) throw new Error('No "TC ID" header found in Excel file.');

  const headers = rawData[headerRowIndex];
  const tests = [];

  for (let i = headerRowIndex + 1; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row || row.length === 0) continue;

    const tc: any = {};
    headers.forEach((h: string, idx: number) => {
      if (h) tc[String(h).trim()] = row[idx];
    });

    const tcId = getSafeCellText(tc['TC ID']);
    
    // Only process rows with a valid TC ID
    if (tcId && tcId.length > 0) {
      tc['TC ID'] = tcId;
      // Handle potentially empty names or names with newlines
      tc['Test case name'] = getSafeCellText(tc['Test case name']) || 'Unnamed Test'; 
      tc['Input'] = getSafeCellText(tc['Input']);
      tc['Expected output'] = getSafeCellText(tc['Expected output']);
      
      tests.push(tc);
    }
  }

  // ★ CRITICAL FIX 1: Sort Tests ★
  // Ensures the Master and Worker processes see the exact same list in the same order.
  tests.sort((a, b) => a['TC ID'].localeCompare(b['TC ID']));

  return tests;
}

const testCases = loadTestsSync();
console.log(`Loaded ${testCases.length} tests from Excel.`);

test.describe('Singlish Translator', () => {

  test.beforeEach(async ({ page }) => {
    await page.goto('https://www.swifttranslator.com/');
    await expect(page.locator('text=Singlish ↔ English Translator')).toBeVisible({ timeout: 10000 });
  });

  for (const tc of testCases) {
    // ★ CRITICAL FIX 2: Sanitize Title ★
    // We remove newlines (\n) from the title because they cause "Test not found" errors in Playwright workers.
    const cleanName = tc['Test case name'].replace(/[\n\r]+/g, ' ');
    const testTitle = `${tc['TC ID']}: ${cleanName}`;

    test(testTitle, async ({ page }) => {
      const input = tc['Input'];
      
      // We trim the expected output to remove accidental Excel cell spacing,
      // BUT we do NOT remove internal spaces, so "A  B" vs "A B" will properly fail.
      let expected = tc['Expected output'];
      if (expected.includes('display:')) expected = expected.split('display:')[1].trim();
      else expected = expected.trim();

      const inputBox = page.getByPlaceholder('Input Your Singlish Text Here.');
      await expect(inputBox).toBeVisible();
      await inputBox.clear();
      await inputBox.fill(input);

      // Verify the output selector matches your site structure
      const outputBox = page.locator('.card', { hasText: 'Sinhala' }).locator('.bg-slate-50');
      
      let actual = '';
      let status = 'Fail';

      try {
        // ★ STRICT CHECKING ★
        // expect.poll waits until the text STRICTLY matches.
        // If the site has "A  B" and expected is "A B", this will FAIL (as requested).
        await expect.poll(async () => {
          return await outputBox.innerText();
        }, { 
          timeout: 8000,
          message: `Expected strict match: "${expected}"`
        }).toBe(expected);

        status = 'Pass';
        actual = await outputBox.innerText();
      } catch (e) {
        status = 'Fail';
        
        // ★ FIX FOR NULL OUTPUT ★
        // If the test fails, we forcefully grab the text so it appears in the Excel report.
        try { 
            if (await outputBox.isVisible()) {
                actual = await outputBox.innerText(); 
            } else {
                actual = "Element not visible";
            }
        } catch { 
            actual = "Error retrieving text"; 
        }
        
        // Re-throw so Playwright marks it red
        throw e; 
      } finally {
        resultsMap.set(tc['TC ID'], {
          actual: actual.trim(),
          status: status
        });
      }
    });
  }

  // --- 2. WRITER (Write results back to Excel) ---
  test.afterAll(async () => {
    if (resultsMap.size === 0) return;

    console.log(`\n💾 Saving results for ${resultsMap.size} test cases...`);
    const filePath = path.join(process.cwd(), FILE_NAME);

    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(filePath);
    } catch (e) {
        console.error("❌ ERROR: Could not open Excel file. Is it open? Please close it.");
        return;
    }

    const sheet = workbook.worksheets[0];

    // Find Header Columns dynamically
    let tcIdCol = -1, actualCol = -1, statusCol = -1, headerRowIdx = -1;

    sheet.eachRow((row, idx) => {
      if (headerRowIdx !== -1) return;
      row.eachCell((cell) => {
        if (getSafeCellText(cell.value) === 'TC ID') headerRowIdx = idx;
      });
    });

    if (headerRowIdx === -1) {
        console.error("❌ Could not find TC ID header row.");
        return;
    }

    const headerRow = sheet.getRow(headerRowIdx);
    headerRow.eachCell((cell, colIdx) => {
        const txt = getSafeCellText(cell.value).toLowerCase();
        if (txt === 'tc id') tcIdCol = colIdx;
        if (txt.includes('actual output')) actualCol = colIdx;
        if (txt === 'status') statusCol = colIdx;
    });

    // Update Rows
    let updatesCount = 0;
    sheet.eachRow((row, rowIdx) => {
        if (rowIdx <= headerRowIdx) return; // Skip headers

        const cellVal = getSafeCellText(row.getCell(tcIdCol).value);
        
        if (resultsMap.has(cellVal)) {
            const res = resultsMap.get(cellVal)!;
            
            // Update Actual Output
            row.getCell(actualCol).value = res.actual;
            
            // Update Status with Color
            const statusCell = row.getCell(statusCol);
            statusCell.value = res.status;

            if (res.status === 'Pass') {
                statusCell.font = { color: { argb: 'FF008000' }, bold: true }; // Green
            } else {
                statusCell.font = { color: { argb: 'FFFF0000' }, bold: true }; // Red
            }

            row.commit();
            updatesCount++;
        }
    });

    await workbook.xlsx.writeFile(filePath);
    console.log(`✅ Success! Updated ${updatesCount} rows in Excel.`);
  });
});