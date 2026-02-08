import { convertExcelToJson, ensureDataDir, MappingItem } from "./utils";

const TMS_JSON_FIELDS: MappingItem[] = [
  "status",
  "waybillNo",
  "TMSDNs",
  "SAPDN",
  "addingPoint",
  "loadingRate",
  "budgetNumber",
  "issuingWarehouse",
  "receivingWarehouse",
  "carrierName",
  "truckType",
  "estimatePickUpTime",
  "refuseToOrder",
  "rejectReasonDetail",
  "createBy",
  "creationDate",
];

// Ensure data directory exists
ensureDataDir();

// Execute conversions
convertExcelToJson("tms.xlsx", "tms.json", TMS_JSON_FIELDS);
