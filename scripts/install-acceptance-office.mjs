/** Pinned external acceptance tool; never imported by an SDK package. */
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { appendFileSync, existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";

if (process.platform !== "linux" || process.arch !== "x64") throw new Error("The pinned Office acceptance tool requires Linux x64.");
const root = resolve(import.meta.dirname, ".."), cache = join(root, ".cache");
const version = "25.8.1.1", name = `LibreOffice_${version}_Linux_x86-64_deb.tar.gz`;
const url = `https://downloadarchive.documentfoundation.org/libreoffice/old/${version}/deb/x86_64/${name}`;
// Published by the same official archive at `${url}.sha256`.
const sha256 = "f367c434aee8352e9742e192bedcdaef194ffcbcfeaf760c6300fb4fecbba73c";
const archive = join(cache, name), directory = join(cache, "acceptance-office"), unpacked = join(directory, "archive"), installed = join(directory, "installed"), bin = join(directory, "bin");
mkdirSync(cache, { recursive: true });
if (!existsSync(archive)) execFileSync("curl", ["--fail", "--location", "--retry", "2", "--max-time", "300", url, "--output", archive], { stdio: "inherit", timeout: 310_000 });
if (createHash("sha256").update(readFileSync(archive)).digest("hex") !== sha256) {
  rmSync(archive);
  throw new Error("Office acceptance archive digest mismatch.");
}
rmSync(directory, { recursive: true, force: true });
for (const path of [unpacked, installed, bin]) mkdirSync(path, { recursive: true });
execFileSync("tar", ["-xzf", archive, "-C", unpacked], { stdio: "inherit", timeout: 60_000 });
const distribution = readdirSync(unpacked).filter(name => name.startsWith("LibreOffice_"));
if (distribution.length !== 1) throw new Error("The Office acceptance archive has an unexpected root.");
const debs = join(unpacked, distribution[0], "DEBS");
for (const file of readdirSync(debs).filter(name => name.endsWith(".deb")).sort())
  execFileSync("dpkg-deb", ["--extract", join(debs, file), installed], { stdio: "inherit", timeout: 60_000 });
const executable = join(installed, "opt/libreoffice25.8/program/soffice");
const actual = execFileSync(executable, ["--version"], { encoding: "utf8", timeout: 30_000 }).trim();
if (!actual.startsWith("LibreOffice " + version + " ")) throw new Error("The Office acceptance version differs from its pinned archive.");
const quote = value => "'" + value.replaceAll("'", "'\\''") + "'";
writeFileSync(join(bin, "libreoffice"), "#!/bin/sh\nexec " + quote(executable) + ' "$@"\n', { mode: 0o755 });
writeFileSync(join(directory, "environment.json"), JSON.stringify({ version: actual, url, sha256, scope: "external development acceptance only" }, null, 2) + "\n");
if (process.env.GITHUB_PATH) appendFileSync(process.env.GITHUB_PATH, bin + "\n");
console.log(actual + "\nAcceptance PATH: " + bin);
