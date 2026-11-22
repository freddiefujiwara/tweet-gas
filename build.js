const fs = require("fs");
const path = require("path");

const build = ({
  rootDir = process.cwd(),
  fsModule = fs,
  pathModule = path,
} = {}) => {
  const srcPath = pathModule.join(rootDir, "src", "Code.js");
  const appsscriptPath = pathModule.join(rootDir, "appsscript.json");
  const distDir = pathModule.join(rootDir, "dist");
  const distPath = pathModule.join(distDir, "Code.gs");
  const distManifestPath = pathModule.join(distDir, "appsscript.json");

  fsModule.mkdirSync(distDir, { recursive: true });
  const source = fsModule.readFileSync(srcPath, "utf8");
  const compiled = source
    .replace(/^export\s+/gm, "")
    .replace(/^export\s*\{[^}]*\};?\s*$/gm, "");
  fsModule.writeFileSync(distPath, compiled);
  fsModule.copyFileSync(appsscriptPath, distManifestPath);
};

const runIfMain = ({ main = require.main, current = module, buildFn = build } = {}) => {
  if (main === current) {
    buildFn();
    return true;
  }
  return false;
};

runIfMain();

module.exports = { build, runIfMain };
