import * as fs from "fs";
import * as path from "path";

const DATA_DIR = path.join(__dirname, "../data");

export function readJsonDb(filename: string): any {
  const filePath = path.join(DATA_DIR, filename);
  if (!fs.existsSync(filePath)) {
    throw new Error(`Database file not found: ${filePath}`);
  }
  const content = fs.readFileSync(filePath, "utf-8");
  return JSON.parse(content);
}

export function getCbmDb() {
  return readJsonDb("cbm.json");
}

export function getAppendixPriceDb() {
  return readJsonDb("appendix-price.json");
}

export function getAddressMockDb() {
  // Read address mock data
  const addressMockPath = path.join(__dirname, "address-mock.json");
  let addressMock: Record<
    string,
    {
      oldWard: string;
      oldProvince: string;
      newWard: string;
      newProvince: string;
    }
  > = {};

  if (fs.existsSync(addressMockPath)) {
    try {
      addressMock = JSON.parse(fs.readFileSync(addressMockPath, "utf-8"));
    } catch (e) {
      console.warn("Failed to read address-mock.json", e);
    }
  }

  return addressMock;
}

export function getAddressInfo(address: string) {
  return readJsonDb("address-info.json")[address];
}

export function searchAddressDistance(address: string) {
  const addressDistancePath = path.join(__dirname, "address-distance.json");
  let addressDistance: Record<string, string> = {};

  if (fs.existsSync(addressDistancePath)) {
    try {
      addressDistance = JSON.parse(
        fs.readFileSync(addressDistancePath, "utf-8"),
      );
    } catch (e) {
      console.warn("Failed to read address-distance.json", e);
    }
  }

  return addressDistance[address] ? parseFloat(addressDistance[address]) : null;
}
