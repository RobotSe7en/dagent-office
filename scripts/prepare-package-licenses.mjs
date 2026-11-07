import { copyFileSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repositoryRoot = fileURLToPath(new URL("../", import.meta.url));

// Each package's prepack hook copies the canonical notices into its tarball.
// These generated copies are ignored by Git; the root documents are authoritative.
for (const filename of ["LICENSE", "NOTICE", "COMMERCIAL-LICENSE.md"]) {
  copyFileSync(resolve(repositoryRoot, filename), resolve(process.cwd(), filename));
}
