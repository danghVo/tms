import * as XLSX from "xlsx-js-style";
import * as path from "path";
import { readExcelData, SAMPLE_DIR, MappingItem } from "./utils";
import {
  getCbmDb,
  getAppendixPriceDb,
  getAddressMockDb,
  searchAddressDistance,
  getAddressInfo,
} from "./db";
import { AppendixPrice, Cbm, Operation, Tms } from "./types";
import { readOperationExcelData } from "./convert-operation";
import * as fs from "fs";

// Define Mappings
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

const OUTPUT_HEADERS = [
  "No\nSTT",
  "Shipped date\nNgày xuất",
  "Delivery No\nSố lệnh",
  "PO No\nSố đơn hàng",
  "Waybill No\nSố vận đơn",
  "Warehouse code\nMã kho",
  "Vehical No\nSố xe",
  "Tonnage\nTrọng tải",
  "Categorize\nChủng loại",
  "Material code\nMã hàng",
  "Models\nSản phẩm",
  "Quantities\nSố lượng",
  "CBM\nThể tích",
  "Total CBM\nTổng thể tích",
  "Sold to code\nMã đại lý",
  "Sold to top name\nNhóm đại lý",
  "Ship to name\nTên đại lý",
  "Ship to old address\nĐịa chỉ giao hàng cũ",
  "Old ward\nPhường xã cũ",
  "Old province\nTĩnh cũ",
  "Ship to New address\nĐịa chỉ giao hàng Mới",
  "New ward\nPhường xã Mới",
  "New province\nTĩnh Mới",
  "Route Code\nVùng",
  "Charging points\nĐiểm tính cước",
  "Abnormal km\nKm trái tuyến",
  "Longest distance in km\nSố km điểm xa nhất",
  "Point code\nMã điểm tính xa nhất",
  "Unit loading price\nĐơn giá đưa hàng xuống đất",
  "Unit price\nĐơn giá chuyến",
  "Adding point\nĐiểm ghép",
  "Loading fee\nPhí đưa hàng xuống đất",
  "Transfer fee\nPhí Chuyển Tải",
  "Overnight fee\nNeo đêm",
  "Return fee\nPhí mang hàng về",
  "Abnormal fee\nPhí trái tuyến",
  "Waiting fee\nPhí chờ",
  "Others fee\nPhí khác",
  "Total Amount\nTổng tiền",
  "Total Amount by trip\nTổng tiền theo chuyến",
  "Note\nGhi chú",
];

const readImportFile = (): { tmsData: Tms[]; opData: Operation[] } => {
  const tmsData = readExcelData<Tms>("tms.xlsx", TMS_JSON_FIELDS);

  const opData = readOperationExcelData("Operation-sample.xlsx");

  return { tmsData, opData };
};

const readDb = (): { cbmDb: Cbm[]; appendixDb: AppendixPrice[] } => {
  const cbmDb = getCbmDb() as Cbm[];
  const appendixDb = getAppendixPriceDb() as AppendixPrice[];

  return { cbmDb, appendixDb };
};

const groupAndSortOpData = (opData: Operation[]): Map<string, Operation[]> => {
  const wbMap = new Map<string, Operation[]>();
  opData.forEach((row) => {
    const waybillNo = row.waybillNo;
    if (!wbMap.has(waybillNo)) {
      wbMap.set(waybillNo, []);
    }

    const rows = wbMap.get(waybillNo);
    rows?.push(row);
  });

  wbMap.forEach((rows) => {
    rows.sort((a, b) => {
      const distA = searchAddressDistance(a.shipToAddress) || 0;
      const distB = searchAddressDistance(b.shipToAddress) || 0;
      return distB - distA;
    });
  });

  return wbMap;
};

const checkIsAddingPoint = (
  rows: Operation[],
  addressMock: Record<string, { oldWard: string }>,
): boolean => {
  const firstRowWard = addressMock[rows[0].shipToAddress]?.oldWard || "";

  return rows.slice(1).some((r) => {
    const currentWard = addressMock[r.shipToAddress]?.oldWard || "";
    return currentWard !== firstRowWard;
  });
};

const calculateTotalTrip = (
  rows: Operation[],
  cbmDb: Cbm[],
  addressMock: Record<string, { oldWard: string }>,
  appendixDb: AppendixPrice[],
  tmsMap: Map<string, Tms>,
  isHaveAddingPoint: boolean,
  additionalMap?: Map<string, any>,
): number => {
  return rows.reduce((acc, r, i) => {
    const rCbmEntry = cbmDb.find((c) => c.materialNo === r.materialNo);
    const rCbmPerUnit = rCbmEntry ? Number(rCbmEntry.cbm) : 0;
    const rTotalCbm = rCbmPerUnit * Number(r.deliveryQuantity);
    const rLoadingFee = 25000 * rTotalCbm;

    let rUnitPrice = 0;
    let rExtraFees = 0;

    if (i === 0) {
      // Main point logic
      const rAddressInfo = addressMock[r.shipToAddress] || {};
      const rOldWard = rAddressInfo.oldWard || "";
      const rAppendix = appendixDb.find((a) => a.ward === rOldWard);
      const rTmsRow =
        tmsMap.get(r.waybillNo ? String(r.waybillNo).trim() : "") ||
        ({} as Tms);
      const rPrice = rAppendix?.trip_price[rTmsRow.truckType];

      if (rPrice) {
        rUnitPrice = Number(rPrice);
      }

      // Add extra fees from additionalMap for the trip (waybill)
      if (additionalMap) {
        const waybillNo = r.waybillNo ? String(r.waybillNo).trim() : "";
        const additionalInfo = additionalMap.get(`${waybillNo}_${r.materialNo}`) || {};
        // Sum known numerical fields
        rExtraFees += Number(additionalInfo.overnightFee || 0);
        rExtraFees += Number(additionalInfo.transferFee || 0);
        rExtraFees += Number(additionalInfo.returnFee || 0);
        rExtraFees += Number(additionalInfo.abnormalFee || 0);
        rExtraFees += Number(additionalInfo.waitingFee || 0);
        rExtraFees += Number(additionalInfo.othersFee || 0);
      }
    }

    const rAddingPoint = isHaveAddingPoint && i > 0 ? 90000 : 0;

    return acc + rUnitPrice + rLoadingFee + rAddingPoint + rExtraFees;
  }, 0);
};

const formatCurrency = (element: number | string): string => {
  if (element === "" || element === null || element === undefined) return "";
  const num = Number(element);
  if (isNaN(num)) return "";
  return num.toLocaleString("de-DE", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
};

const processData = (
  tmsData: Tms[],
  opData: Operation[],
  cbmDb: Cbm[],
  appendixDb: AppendixPrice[],
  additionalMap?: Map<string, any>,
): any[][] => {
  const addressMock = getAddressMockDb();

  // Create a map of TMS data by Waybill for faster lookup
  const tmsMap = new Map<string, Tms>();
  tmsData.forEach((row) => {
    if (row.waybillNo) {
      tmsMap.set(String(row.waybillNo).trim(), row);
    }
  });

  const outputRows: any[][] = [];
  let stt = 0;

  const wbMap = groupAndSortOpData(opData);

  wbMap.forEach((rows) => {
    stt++;
    const isHaveAddingPoint = checkIsAddingPoint(rows, addressMock);
    rows.forEach((opRow, index) => {
      const waybillNo = opRow.waybillNo ? String(opRow.waybillNo).trim() : "";
      const tmsRow = tmsMap.get(waybillNo) || ({} as Tms);

      // Lookup CBM
      const cbmEntry = cbmDb.find((c) => c.materialNo === opRow.materialNo);
      const cbmPerUnit = cbmEntry ? Number(cbmEntry.cbm) : 0;
      const totalCbm = cbmPerUnit * (parseFloat(opRow.deliveryQuantity) || 0);
      const category = cbmEntry ? cbmEntry.category : "";

      // Lookup Address Info from Mock
      const addressInfo = addressMock[opRow.shipToAddress] || {};
      const oldWard = addressInfo.oldWard || "";
      const oldProvince = addressInfo.oldProvince || "";
      const newWard = addressInfo.newWard || "";
      const newProvince = addressInfo.newProvince || "";
      const appendix = appendixDb.find((a) => a.ward === oldWard);
      const routeCode = appendix?.area || "";

      let longestPoint = {
        distance: "",
        code: "",
        ward: "",
        unitPrice: "",
      };
      let totalTripAmount = 0;

      if (index === 0) {
        longestPoint = {
          distance:
            searchAddressDistance(opRow.shipToAddress)?.toString() || "",
          code: appendix?.code || "",
          ward: addressInfo.oldWard,
          unitPrice: appendix?.trip_price[tmsRow.truckType].toString() || "",
        };

        totalTripAmount = calculateTotalTrip(
          rows,
          cbmDb,
          addressMock,
          appendixDb,
          tmsMap,
          isHaveAddingPoint,
          additionalMap,
        );
      }
      // Placeholder logic for price
      const loadingFee = Number(25000 * totalCbm);
      const unitPrice = longestPoint.unitPrice
        ? Number(longestPoint.unitPrice)
        : 0;
      const addingPointCost = isHaveAddingPoint && index > 0 ? 90000 : 0;

      const additionalInfo =
        additionalMap?.get(`${waybillNo}_${opRow.materialNo}`) || {};
      const transferFee =
        index === 0 ? Number(additionalInfo.transferFee || 0) : 0;
      const overnightFee =
        index === 0 ? Number(additionalInfo.overnightFee || 0) : 0;
      const returnFee = index === 0 ? Number(additionalInfo.returnFee || 0) : 0;
      const abnormalFee =
        index === 0 ? Number(additionalInfo.abnormalFee || 0) : 0;
      const waitingFee =
        index === 0 ? Number(additionalInfo.waitingFee || 0) : 0;
      const othersFee = index === 0 ? Number(additionalInfo.othersFee || 0) : 0;

      const totalAmount =
        (unitPrice ?? 0) +
        loadingFee +
        addingPointCost +
        transferFee +
        overnightFee +
        returnFee +
        abnormalFee +
        waitingFee +
        othersFee;

      const rowData = [
        stt, // No/STT
        opRow.pickupDate || tmsRow.estimatePickUpTime,
        opRow.dnNo,
        opRow.poNo,
        waybillNo,
        opRow.storageLocation,
        opRow.vehicleNo,
        tmsRow.truckType,
        category,
        opRow.materialNo,
        opRow.materialDesp,
        opRow.deliveryQuantity,
        cbmPerUnit.toFixed(2),
        totalCbm.toFixed(2),
        opRow.soldToCode,
        opRow.soldToName,
        opRow.shipToName,
        "",
        oldWard,
        oldProvince,
        opRow.shipToAddress,
        newWard,
        newProvince,
        routeCode,
        longestPoint.ward ||
          (isHaveAddingPoint ? addressMock[opRow.shipToAddress]?.oldWard : ""),
        "",
        longestPoint.distance,
        longestPoint.code,
        25000,
        longestPoint.unitPrice ? unitPrice : "",
        addingPointCost ? addingPointCost : "",
        loadingFee,
        transferFee ? transferFee : "",
        overnightFee ? overnightFee : "",
        returnFee ? returnFee : "",
        abnormalFee ? abnormalFee : "",
        waitingFee ? waitingFee : "",
        othersFee ? othersFee : "",
        totalAmount, // Total Amount
        index === 0 ? totalTripAmount : "", // Total Amount by trip
        "", // Note
      ];

      outputRows.push(rowData);
    });
  });

  return outputRows;
};

async function main() {
  console.log("Starting processing...");

  const { tmsData, opData } = readImportFile();
  const { cbmDb, appendixDb } = readDb();

  const outputRows = processData(tmsData, opData, cbmDb, appendixDb);

  // 4. Write Output
  const ws = XLSX.utils.aoa_to_sheet([OUTPUT_HEADERS, ...outputRows]);

  // Style the header row (Row 0)
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const address = XLSX.utils.encode_cell({ r: 0, c: C });
    if (!ws[address]) continue;
    ws[address].s = {
      alignment: { wrapText: true, horizontal: "center", vertical: "center" },
      font: { bold: true },
    };
  }

  // Format currency columns
  const currencyCols = [28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39];
  for (let R = range.s.r + 1; R <= range.e.r; ++R) {
    currencyCols.forEach((C) => {
      const address = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[address]) return;
      if (typeof ws[address].v === "number") {
        ws[address].z = "#,##0";
      }
    });
  }

  // Format border
  const borderStyle = {
    top: { style: "thin" },
    bottom: { style: "thin" },
    left: { style: "thin" },
    right: { style: "thin" },
  };

  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const address = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[address]) continue;
      if (!ws[address].s) ws[address].s = {};
      ws[address].s.border = borderStyle;
    }
  }
  // Set column widths for better visibility
  ws["!cols"] = OUTPUT_HEADERS.map(() => ({ wch: 20 }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

  // Saving to a new file
  const outputPath = path.join(SAMPLE_DIR, "Output-sample-generated.xlsx");
  XLSX.writeFile(wb, outputPath);
  console.log(`Generated result at ${outputPath}`);
}

main().catch(console.error);
