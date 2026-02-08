import * as XLSX from "xlsx";
import * as path from "path";

const SAMPLE_DIR = path.join(__dirname, "sample"); // Corrected path

function inspectFile(filename: string) {
  console.log(`--- Inspecting ${filename} ---`);
  try {
    const filePath = path.join(SAMPLE_DIR, filename);
    if (!require("fs").existsSync(filePath)) {
      console.log(`File not found: ${filePath}`);
      return;
    }
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
    // print first 5 rows
    data
      .slice(0, 5)
      .forEach((row, i) => console.log(`Row ${i}:`, JSON.stringify(row)));
  } catch (e) {
    console.error(e);
  }
}

inspectFile("tms.xlsx");
inspectFile("Operation-sample.xlsx");
