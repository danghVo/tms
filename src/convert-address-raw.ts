import * as fs from "fs";
import * as path from "path";

const INPUT_FILE = path.join(__dirname, "address_raw.txt");
const OUTPUT_FILE = path.join(__dirname, "address-mock.json");

function processFile() {
  if (!fs.existsSync(INPUT_FILE)) {
    console.error(`Input file not found: ${INPUT_FILE}`);
    return;
  }

  const content = fs.readFileSync(INPUT_FILE, "utf-8");
  const lines = content.split("\n").filter((line) => line.trim() !== "");
  const result: Record<
    string,
    {
      oldWard: string;
      oldProvince: string;
      newWard: string;
      newProvince: string;
    }
  > = {};

  lines.forEach((line) => {
    const parts = line.split("|").map((p) => p.trim());
    if (parts.length >= 5) {
      // Column ordering as per user request:
      // Col 1: Old Ward (index 0)
      // Col 2: Old Province (index 1)
      // Col 3: Address (KEY) (index 2)
      // Col 4: New Ward (index 3)
      // Col 5: New Province (index 4)

      const oldWard = parts[0];
      const oldProvince = parts[1];
      const addressKey = parts[2];
      const newWard = parts[3];
      const newProvince = parts[4];

      result[addressKey] = {
        oldWard,
        oldProvince,
        newWard,
        newProvince,
      };
    }
  });

  fs.writeFileSync(OUTPUT_FILE, JSON.stringify(result, null, 2));
  console.log(`Successfully generated ${OUTPUT_FILE}`);
}

processFile();
