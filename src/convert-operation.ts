import * as fs from "fs";
import path from "path";
import { DATA_DIR, ensureDataDir, MappingItem, readExcelData } from "./utils";
import { Operation } from "./types";

const OPERATION_JSON_FIELDS: MappingItem[] = [
  "vehicleNo",
  "waybillNo",
  "pickupDate",
  "soNo",
  "soldToCode",
  "dnNo",
  "materialNo",
  "materialDesp",
  "deliveryQuantity",
  "storageLocation",
  "poNo",
  "shipToName",
  "soldToName",
  "shipToAddress",
  "rdd",
];

// Ensure data directory exists
// ensureDataDir();

export const readOperationExcelData = (fileName: string) => {
  const result = readExcelData<Operation>(fileName, OPERATION_JSON_FIELDS);
  return result.map((item) => {
    return {
      ...item,
      pickupDate: item.pickupDate.replace(/\./g, "/"),
    };
  });
};

// export function convertOperationExcelToJson(
//   inputFileName: string,
//   outputFileName: string,
//   mapping?: MappingItem[],
//   dataStartOffset: number = 1,
// ) {
//   const outputPath = path.join(DATA_DIR, outputFileName);

//   try {
//     const result = readExcelData<Operation>(
//       inputFileName,
//       mapping,
//       dataStartOffset,
//     );
//     const finalData = result.map((item) => {
//       return {
//         ...item,
//         pickupDate: item.pickupDate.replace(/\./g, "/"),
//       };
//     });

//     fs.writeFileSync(outputPath, JSON.stringify(finalData, null, 2));
//     console.log(`Successfully generated ${outputPath}`);
//   } catch (err) {
//     console.error(`Error processing ${inputFileName}:`, err);
//   }
// }

// Execute conversions
// convertOperationExcelToJson(
//   "Operation-sample.xlsx",
//   "operation.json",
//   OPERATION_JSON_FIELDS,
// );
