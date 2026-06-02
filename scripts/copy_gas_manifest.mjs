import fs from "node:fs/promises";
import path from "node:path";

const rootDir = process.cwd();
const gasDir = path.join(rootDir, "gas");
const buildDir = path.join(gasDir, "build");

try {
  await fs.mkdir(buildDir, { recursive: true });
  await fs.copyFile(path.join(gasDir, "appsscript.json"), path.join(buildDir, "appsscript.json"));
  console.log("Successfully copied appsscript.json to gas/build/");
} catch (err) {
  console.error("Error copying appsscript.json:", err);
  process.exit(1);
}
