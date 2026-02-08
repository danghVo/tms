// Interface Definitions
export interface Tms {
  status: string;
  waybillNo: string;
  TMSDNs: string;
  SAPDN: string;
  addingPoint: string;
  loadingRate: string;
  budgetNumber: string;
  issuingWarehouse: string;
  receivingWarehouse: string;
  carrierName: string;
  truckType: string;
  estimatePickUpTime: string;
  refuseToOrder: string;
  rejectReasonDetail: string;
  createBy: string;
  creationDate: string;
}

export interface Operation {
  vehicleNo: string;
  waybillNo: string;
  pickupDate: string;
  soNo: string;
  soldToCode: string;
  dnNo: string;
  materialNo: string;
  materialDesp: string;
  deliveryQuantity: string;
  storageLocation: string;
  poNo: string;
  shipToName: string;
  soldToName: string;
  shipToAddress: string;
  rdd: string;
}

export interface Cbm {
  materialNo: string;
  materialDesp: string;
  length: number;
  width: number;
  height: number;
  cbm: string; // Often string from Excel unless parsed
  category: string;
}

export interface AppendixPrice {
  area: string;
  city: string;
  ward: string;
  code: string;
  trip_price: Record<string, string | number>;
  cbm_price: Record<string, string | number>;
}
