import * as fs from "fs";
import * as path from "path";
import * as XLSX from "xlsx";

export const SAMPLE_DIR = path.join(__dirname, "../sample");
export const DATA_DIR = path.join(__dirname, "../data");

export type MappingItem = string | { [key: string]: string[] };

export function ensureDataDir() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR);
    console.log(`Created directory: ${DATA_DIR}`);
  }
}

export function processSheetWithMapping(
  worksheet: XLSX.WorkSheet,
  mapping: MappingItem[],
  dataStartOffset: number = 1,
): any[] {
  const jsonData = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
  }) as any[][];

  if (!jsonData || jsonData.length === 0) {
    return [];
  }

  // Find the header row containing "STT"
  let headerRowIndex = -1;
  let sttColIndex = -1;

  for (let i = 0; i < jsonData.length; i++) {
    const row = jsonData[i];
    const index = row.findIndex(
      (cell: any) => cell && String(cell).trim().toUpperCase().includes("STT"),
    );
    if (index !== -1) {
      headerRowIndex = i;
      sttColIndex = index;
      break;
    }
  }

  if (headerRowIndex === -1 || sttColIndex === -1) {
    console.log(
      "STT column not found. Defaulting to starting at column 0, row 0.",
    );
    headerRowIndex = 0;
    sttColIndex = -1; // So dataStartColIndex becomes 0 below
  }

  // Logic: User wants to ignore columns before STT and STT itself.
  let dataStartColIndex = sttColIndex + 1;
  const headerRow = jsonData[headerRowIndex];

  if (
    headerRow[dataStartColIndex] &&
    String(headerRow[dataStartColIndex]).trim().toUpperCase() === "CAT"
  ) {
    console.log(
      "Detected 'CAT' column after STT, skipping it to align mapping.",
    );
    dataStartColIndex++;
  }

  const result: any[] = [];

  // Apply dataStartOffset
  const startRowIndex = headerRowIndex + dataStartOffset;

  // Start reading data from the specified start row
  for (let i = startRowIndex; i < jsonData.length; i++) {
    const row = jsonData[i];

    // Check if row is empty
    if (!row || row.length === 0) continue;

    const item: Record<string, any> = {};
    let hasData = false;
    let currentDataColIndex = dataStartColIndex;

    // Map columns
    for (const mapItem of mapping) {
      if (typeof mapItem === "string") {
        const value =
          currentDataColIndex < row.length
            ? row[currentDataColIndex]
            : undefined;
        if (value !== undefined && value !== null && value !== "") {
          hasData = true;
        }
        item[mapItem] = value;
        currentDataColIndex++;
      } else {
        // It's an object mapping like { trip_price: [...] }
        const groupKey = Object.keys(mapItem)[0];
        const subKeys = mapItem[groupKey];

        const groupObj: Record<string, any> = {};

        for (const subKey of subKeys) {
          const value =
            currentDataColIndex < row.length
              ? row[currentDataColIndex]
              : undefined;
          if (value !== undefined && value !== null && value !== "") {
            hasData = true;
          }
          groupObj[subKey] = value;
          currentDataColIndex++;
        }
        item[groupKey] = groupObj;
      }
    }

    if (hasData) {
      result.push(item);
    }
  }

  return result;
}

export function readExcelData<T>(
  inputFileName: string,
  mapping?: MappingItem[],
  dataStartOffset: number = 1,
): T[] {
  const inputPath = path.join(SAMPLE_DIR, inputFileName);

  if (!fs.existsSync(inputPath)) {
    throw new Error(`Error: Source file not found at ${inputPath}`);
  }

  console.log(`Reading ${inputPath}...`);
  const workbook = XLSX.readFile(inputPath);
  let allData: T[] = [];

  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    let sheetData: T[] = [];
    if (mapping) {
      sheetData = processSheetWithMapping(
        worksheet,
        mapping,
        dataStartOffset,
      ) as T[];
    } else {
      sheetData = XLSX.utils.sheet_to_json(worksheet) as T[];
    }
    allData = allData.concat(sheetData);
  });

  return allData;
}

export function convertExcelToJson(
  inputFileName: string,
  outputFileName: string,
  mapping?: MappingItem[],
  dataStartOffset: number = 1,
) {
  const outputPath = path.join(DATA_DIR, outputFileName);

  try {
    const result = readExcelData(inputFileName, mapping, dataStartOffset);

    fs.writeFileSync(outputPath, JSON.stringify(result, null, 2));
    console.log(`Successfully generated ${outputPath}`);
  } catch (err) {
    console.error(`Error processing ${inputFileName}:`, err);
  }
}
