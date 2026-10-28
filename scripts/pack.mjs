import { execFileSync } from "node:child_process";
import { mkdirSync, readdirSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repositoryRoot = fileURLToPath(new URL("../", import.meta.url));

/** @param {string} destination @returns {string[]} */
export function packPackages(destination) {
  const pnpm = process.env.npm_execpath;
  if (!pnpm) throw new Error("Run this command through pnpm.");
  const directory = resolve(destination);
  mkdirSync(directory, { recursive: true });
  execFileSync(process.execPath, [
    pnpm, "--recursive", "--filter", "./packages/*",
    "pack", "--pack-destination", directory,
  ], { cwd: repositoryRoot, stdio: "pipe", timeout: 60_000 });
  return readdirSync(directory).filter((name) => name.endsWith(".tgz"))
    .map((name) => resolve(directory, name));
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const destination = resolve(repositoryRoot, process.argv[2] ?? ".cache/packages");
  console.log(packPackages(destination).join("\n"));
}
