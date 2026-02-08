import path from "path";
import * as fs from "fs";
import { DATA_DIR, ensureDataDir, MappingItem, readExcelData } from "./utils";
import { Cbm } from "./types";

const CBM_FIELDS_MAPPING: MappingItem[] = [
  "materialNo",
  "materialDesp",
  "length",
  "width",
  "height",
  "cbm",
  "category",
];

// Ensure data directory exists
ensureDataDir();

export function convertExcelToJson(
  inputFileName: string,
  outputFileName: string,
  mapping?: MappingItem[],
  dataStartOffset: number = 1,
) {
  const outputPath = path.join(DATA_DIR, outputFileName);

  try {
    const result = readExcelData<Cbm>(inputFileName, mapping, dataStartOffset);
    const final = result.map((item) => {
      return {
        materialNo: item.materialNo,
        materialDesp: item.materialDesp,
        length: Number(item.length),
        width: Number(item.width),
        height: Number(item.height),
        cbm:
          (Number(item.length) * Number(item.width) * Number(item.height)) /
          1000000000,
        category: item.category,
      };
    });

    fs.writeFileSync(outputPath, JSON.stringify(final, null, 2));
    console.log(`Successfully generated ${outputPath}`);
  } catch (err) {
    console.error(`Error processing ${inputFileName}:`, err);
  }
}

// Execute conversions
convertExcelToJson("CBM.xlsx", "cbm.json", CBM_FIELDS_MAPPING);
