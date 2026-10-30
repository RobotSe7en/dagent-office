import { readFileSync, readdirSync } from "node:fs";
import { isBuiltin } from "node:module";
import { dirname, isAbsolute, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import ts from "typescript";

const repositoryRoot = fileURLToPath(new URL("../", import.meta.url));
const layers = {
  "packages/core": [],
  "packages/render": ["packages/core"],
  "packages/editor": ["packages/core", "packages/render"],
  "packages/office": ["packages/core"],
  "packages/node": ["packages/core", "packages/office", "packages/render"],
  "apps/playground": ["packages/core", "packages/editor", "packages/office", "packages/node"],
};

function sourceFiles(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = resolve(directory, entry.name);
    return entry.isDirectory()
      ? sourceFiles(path)
      : /\.(?:[cm]?[jt]s|[jt]sx)$/.test(entry.name) ? [path] : [];
  });
}

/** Collect static module references, including type imports and dynamic imports. */
function moduleReferences(path) {
  const source = ts.createSourceFile(path, readFileSync(path, "utf8"), ts.ScriptTarget.Latest, true);
  const references = [];
  function visit(node) {
    if (ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) {
      if (node.moduleSpecifier) references.push(node.moduleSpecifier);
    } else if (ts.isImportTypeNode(node) && ts.isLiteralTypeNode(node.argument)) {
      references.push(node.argument.literal);
    } else if (ts.isImportEqualsDeclaration(node) && ts.isExternalModuleReference(node.moduleReference)) {
      if (node.moduleReference.expression) references.push(node.moduleReference.expression);
    } else if (
      ts.isCallExpression(node)
      && (node.expression.kind === ts.SyntaxKind.ImportKeyword
        || (ts.isIdentifier(node.expression) && node.expression.text === "require"))
    ) {
      references.push(node.arguments[0]);
    }
    ts.forEachChild(node, visit);
  }
  visit(source);
  return references.map((node) => node && ts.isStringLiteralLike(node) ? node.text : null);
}

function contains(directory, path) {
  const suffix = relative(directory, path);
  return suffix !== ".." && !suffix.startsWith(".." + sep) && !isAbsolute(suffix);
}

/** @param {string} root @returns {string[]} */
export function checkBoundaries(root = repositoryRoot) {
  const projects = Object.entries(layers).map(([path, allowed]) => ({
    path,
    allowed,
    directory: resolve(root, path),
    manifest: JSON.parse(readFileSync(resolve(root, path, "package.json"), "utf8")),
  }));
  const violations = [];
  for (const project of projects) {
    const runtime = {
      ...project.manifest.dependencies,
      ...project.manifest.optionalDependencies,
      ...project.manifest.peerDependencies,
    };
    const allDependencies = { ...runtime, ...project.manifest.devDependencies };
    const report = (file, message) => violations.push(relative(root, file) + ": " + message);
    for (const [name, version] of Object.entries(allDependencies)) {
      const target = projects.find((item) => item.manifest.name === name);
      if (!target) continue;
      if (!project.allowed.includes(target.path)) {
        report(resolve(project.directory, "package.json"), "Forbidden dependency " + name);
      }
      if (version !== "workspace:*") {
        report(resolve(project.directory, "package.json"), "Use workspace:* for " + name);
      }
    }

    for (const file of sourceFiles(resolve(project.directory, "src"))) {
      for (const specifier of moduleReferences(file)) {
        if (specifier === null) {
          report(file, "Module references must be string literals so package boundaries can be checked");
          continue;
        }
        if (isBuiltin(specifier)) {
          if (project.path !== "packages/node") report(file, "Node builtin is forbidden: " + specifier);
          continue;
        }
        if (specifier.startsWith(".")) {
          if (!contains(project.directory, resolve(dirname(file), specifier))) {
            report(file, "Relative import leaves its package: " + specifier);
          }
          continue;
        }
        const dependency = specifier.startsWith("@")
          ? specifier.split("/").slice(0, 2).join("/")
          : specifier.split("/")[0];
        if (!(dependency in runtime)) {
          report(file, "Undeclared runtime dependency: " + dependency);
        }
        if (
          ["dagent", "dagent-ai", "dagent-enterprise", "openai", "@anthropic-ai/sdk"].includes(dependency)
          || (project.path !== "apps/playground" && project.path !== "packages/editor"
            && ["react", "react-dom"].includes(dependency))
        ) {
          report(file, "Host or UI dependency is forbidden in this layer: " + dependency);
        }
        const target = projects.find((item) => item.manifest.name === dependency);
        if (target) {
          if (!project.allowed.includes(target.path)) report(file, "Forbidden layer: " + dependency);
          const subpath = specifier === dependency ? "." : "." + specifier.slice(dependency.length);
          if (!Object.hasOwn(target.manifest.exports ?? {}, subpath)) {
            report(file, "Import must use a public export: " + specifier);
          }
        }
      }
    }
  }
  return violations;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const violations = checkBoundaries();
  if (violations.length) {
    console.error(violations.join("\n"));
    process.exitCode = 1;
  } else {
    console.log("Package dependencies and source imports respect the architecture.");
  }
}
