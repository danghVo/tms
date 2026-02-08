import { convertExcelToJson, ensureDataDir, MappingItem } from "./utils";

const APPENDIX_PRICE_FIELDS_MAPPING: MappingItem[] = [
  "area",
  "city",
  "ward",
  "code",
  { trip_price: ["2.5T", "3.5T", "5T", "7T", "8T", "11T", "15T"] },
  { cbm_price: ["2cbm", "4cbm", "6bm"] },
];

// Ensure data directory exists
ensureDataDir();

// Execute conversions
// Offset 2 because user says header has 2 rows
convertExcelToJson(
  "appendix-price.xlsx",
  "appendix-price.json",
  APPENDIX_PRICE_FIELDS_MAPPING,
  2,
);
