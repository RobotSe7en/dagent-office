import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, rmSync } from "node:fs";
import { fileURLToPath } from "node:url";

if (process.platform !== "linux" || process.arch !== "x64") throw new Error("The pinned legacy Chromium fixture requires Linux x64.");
const directory = fileURLToPath(new URL("../.cache/legacy-browser/chromium-1041/", import.meta.url));
const archive = fileURLToPath(new URL("../.cache/legacy-browser/chromium-1041.zip", import.meta.url));
const executable = directory + "chrome-linux/chrome";
const digest = "1ec3528cecff35bd20184dc7fd07deb794624023607388b0ea8fda696730805d";
if (!existsSync(executable)) {
  mkdirSync(directory, { recursive: true });
  execFileSync("curl", ["--fail", "--location", "--retry", "2", "--max-time", "240", "https://cdn.playwright.dev/builds/chromium/1041/chromium-linux.zip", "--output", archive], { stdio: "inherit", timeout: 250_000 });
  if (createHash("sha256").update(readFileSync(archive)).digest("hex") !== digest) { rmSync(archive); throw new Error("Legacy browser archive digest mismatch."); }
  execFileSync("unzip", ["-q", "-o", archive, "-d", directory], { stdio: "inherit", timeout: 30_000 });
  rmSync(archive);
}
if (createHash("sha256").update(readFileSync(executable)).digest("hex") !== "7d9e070017f4a8bbc92a3c53fcd97dcc84b8d74322675c6e7094279e2ce98f98") throw new Error("Legacy browser executable digest mismatch.");
console.log(executable);
