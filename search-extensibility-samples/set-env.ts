//#!/usr/bin/env node
 
/**
 * set-env.ts
 * ---------------------------------------------------------------------------
 * Configures SPFx / PnP Modern Search extensibility project files for a
 * specific deployment environment (BASE | DEV | UAT | PROD).
 *
 * Usage:
 *   node set-env.js --env UAT
 *
 * Add to package.json scripts:
 *   "set-env:base": "tsc -p set-env.tsconfig.json && node set-env.js --env BASE",
 *   "set-env:dev":  "tsc -p set-env.tsconfig.json && node set-env.js --env DEV",
 *   "set-env:uat":  "tsc -p set-env.tsconfig.json && node set-env.js --env UAT",
 *   "set-env:prod": "tsc -p set-env.tsconfig.json && node set-env.js --env PROD"
 * ---------------------------------------------------------------------------
 */
 
import * as fs from 'fs';
import * as path from 'path';
 
// ============================================================================
// TYPES
// ============================================================================
 
type Environment = 'BASE' | 'DEV' | 'UAT' | 'PROD';
 
interface EnvConfig {
  suffix: string;
  solutionId: string;
  manifestId: string;
}
 
interface BundleMap {
  [key: string]: unknown;
}
 
interface ConfigJson {
  bundles?: BundleMap;
  [key: string]: unknown;
}
 
interface PackageSolutionPaths {
  zippedPackage?: string;
  [key: string]: unknown;
}
 
interface PackageSolutionJson {
  solution?: {
    name?: string;
    id?: string;
    paths?: PackageSolutionPaths;
    [key: string]: unknown;
  };
  paths?: PackageSolutionPaths;
  [key: string]: unknown;
}
 
interface ManifestJson {
  alias?: string;
  id?: string;
  [key: string]: unknown;
}
 
interface PackageJson {
  name?: string;
  [key: string]: unknown;
}
 
// ============================================================================
// CONFIG — Fill in your real GUIDs and base values here
// ============================================================================
 
const ENV_CONFIG: Record<Environment, EnvConfig> = {
  BASE: {
    suffix:     '',
    solutionId: '4896309f-54bc-48f2-a7ee-0785e019d195',
    manifestId: 'dc4f961b-dbe0-44b4-982d-5776bf99d015',
  },
  DEV: {
    suffix:     'dev',
    solutionId: '4896309f-54bc-48f2-a7ee-0785e019d196',
    manifestId: 'dc4f961b-dbe0-44b4-982d-5776bf99d016',
  },
  UAT: {
    suffix:     'uat',
    solutionId: '4896309f-54bc-48f2-a7ee-0785e019d197',
    manifestId: 'dc4f961b-dbe0-44b4-982d-5776bf99d017',
  },
  PROD: {
    suffix:     'prod',
    solutionId: '4896309f-54bc-48f2-a7ee-0785e019d198',
    manifestId: 'dc4f961b-dbe0-44b4-982d-5776bf99d018',
  },
};
 
// ============================================================================
// FILE PATHS — Relative to the project root (where this script lives)
// ============================================================================
 
const ROOT:        string = path.resolve(__dirname, '..');
const CONFIG_JSON: string = path.join(ROOT, 'config', 'config.json');
const PKG_SOL:     string = path.join(ROOT, 'config', 'package-solution.json');
const PKG_JSON:    string = path.join(ROOT, 'package.json');
const SRC_DIR:     string = path.join(ROOT, 'src');
 
// ============================================================================
// KNOWN ENVIRONMENT SUFFIXES — used for stripping before re-applying
// ============================================================================
 
const KNOWN_SUFFIXES: string[] = [
  ENV_CONFIG.BASE.suffix, 
  ENV_CONFIG.DEV.suffix, 
  ENV_CONFIG.UAT.suffix, 
  ENV_CONFIG.PROD.suffix
];
 
// ============================================================================
// LOGGING
// ============================================================================
 
let warnings: number = 0;
 
function ok(msg: string):     void { console.log(`  ✅  ${msg}`); }
function warn(msg: string):   void { console.warn(`  ⚠️   ${msg}`); warnings++; }
function info(msg: string):   void { console.log(`\n📋  ${msg}`); }
function header(msg: string): void {
  console.log('\n' + '═'.repeat(60));
  console.log(`  ${msg}`);
  console.log('═'.repeat(60));
}
 
// ============================================================================
// HELPERS
// ============================================================================
 
/** Append a suffix to a base string with a separator. Returns base unchanged if no suffix. */
function applyEnvSuffix(base: string, suffix: string, separator: string = '-', firstCharCapitalized: boolean = false): string {
  const suff = !firstCharCapitalized ? suffix : `${suffix.charAt(0).toUpperCase()}${suffix.slice(1, suffix.length)}`;
  return suffix ? `${base}${separator}${suff}` : base;
}
 
/**
 * Strip a known environment suffix from the end of a string.
 * @param value     The full value that may have a suffix applied
 * @param separator The separator used before the suffix (e.g. '-', ' - ', '')
 */
function stripEnvSuffix(value: string, separator: string, firstCharCapitalized: boolean = false): string {
  let result = value;
  let changed = true;

  while (changed) {
    changed = false;

    for (const s of KNOWN_SUFFIXES) {
      if (!s) continue; // skip BASE (empty string)

      const suff = firstCharCapitalized
        ? `${s.charAt(0).toUpperCase()}${s.slice(1)}`
        : s;

      const tail = `${separator}${suff}`;

      if (result.endsWith(tail)) {
        result = result.slice(0, result.length - tail.length);
        changed = true;
      }
    }
  }

  return result;
}
 
/** Return a path relative to the project root (for display purposes). */
function relPath(p: string): string {
  return path.relative(ROOT, p);
}
 
/** Read and parse a JSON file. Returns null and warns on failure. */
function readJson<T>(filePath: string): T | null {
  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf8')) as T;
  } catch (e) {
    warn(`Could not read/parse JSON at ${relPath(filePath)}: ${(e as Error).message}`);
    return null;
  }
}
 
/** Write a JSON file with 2-space indentation. Returns false and warns on failure. */
function writeJson(filePath: string, data: unknown): boolean {
  try {
    fs.writeFileSync(filePath, JSON.stringify(data, null, 2) + '\n', 'utf8');
    return true;
  } catch (e) {
    warn(`Could not write JSON to ${relPath(filePath)}: ${(e as Error).message}`);
    return false;
  }
}
 
/** Read a text file. Returns null and warns on failure. */
function readText(filePath: string): string | null {
  try {
    return fs.readFileSync(filePath, 'utf8');
  } catch (e) {
    warn(`Could not read file at ${relPath(filePath)}: ${(e as Error).message}`);
    return null;
  }
}
 
/** Write a text file. Returns false and warns on failure. */
function writeText(filePath: string, content: string): boolean {
  try {
    fs.writeFileSync(filePath, content, 'utf8');
    return true;
  } catch (e) {
    warn(`Could not write file to ${relPath(filePath)}: ${(e as Error).message}`);
    return false;
  }
}
 
/** Recursively find files under a directory matching a predicate. */
function findFiles(
  dir: string,
  predicate: (name: string, fullPath: string) => boolean,
  results: string[] = []
): string[] {
  if (!fs.existsSync(dir)) return results;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      findFiles(fullPath, predicate, results);
    } else if (entry.isFile() && predicate(entry.name, fullPath)) {
      results.push(fullPath);
    }
  }
  return results;
}
 
// ============================================================================
// STEP 1 — config/config.json
// ============================================================================
 
function updateConfigJson(suffix: string): void {
  info('Step 1 — config/config.json (bundle name)');
 
  const data = readJson<ConfigJson>(CONFIG_JSON);
  if (!data) return;
 
  const bundles = data.bundles;
  if (!bundles || typeof bundles !== 'object') {
    warn(`No "bundles" object found in ${relPath(CONFIG_JSON)}`);
    return;
  }
 
  const keys = Object.keys(bundles);
  if (keys.length === 0) {
    warn(`"bundles" object is empty in ${relPath(CONFIG_JSON)}`);
    return;
  }
 
  let changed = false;
  const updatedBundles: BundleMap = {};
 
  for (const key of keys) {
    const stripped = stripEnvSuffix(key, '-');
    const newKey   = applyEnvSuffix(stripped, suffix, '-');
 
    if (newKey !== key) {
      updatedBundles[newKey] = bundles[key];
      ok(`Bundle key: "${key}" → "${newKey}"`);
      changed = true;
    } else {
      updatedBundles[key] = bundles[key];
      ok(`Bundle key: "${key}" (no change needed)`);
    }
  }
 
  if (changed) {
    data.bundles = updatedBundles;
    if (writeJson(CONFIG_JSON, data)) {
      ok(`Saved ${relPath(CONFIG_JSON)}`);
    }
  }
}
 
// ============================================================================
// STEP 2 — config/package-solution.json
// ============================================================================
 
function updatePackageSolution(suffix: string, solutionId: string): void {
  info('Step 2 — config/package-solution.json (solution name, id, zippedPackage)');
 
  const data = readJson<PackageSolutionJson>(PKG_SOL);
  if (!data) return;
 
  const sol = data.solution;
  if (!sol) {
    warn(`No "solution" object found in ${relPath(PKG_SOL)}`);
    return;
  }
 
  // — name
  if (typeof sol.name === 'string') {
    const strippedName = stripEnvSuffix(sol.name, ' - ');
    const newName      = suffix ? `${strippedName} - ${suffix}` : strippedName;
    ok(`solution.name: "${sol.name}" → "${newName}"`);
    sol.name = newName;
  } else {
    warn(`solution.name not found or not a string in ${relPath(PKG_SOL)}`);
  }
 
  // — id
  ok(`solution.id: "${sol.id}" → "${solutionId}"`);
  sol.id = solutionId;
 
  // — paths.zippedPackage (may live under solution.paths or top-level paths)
  const paths = sol.paths ?? data.paths;
  if (paths && typeof paths.zippedPackage === 'string') {
    const withoutExt  = paths.zippedPackage.replace(/\.sppkg$/, '');
    const strippedPkg = stripEnvSuffix(withoutExt, '-');
    const newZipped   = applyEnvSuffix(strippedPkg, suffix, '-') + '.sppkg';
    ok(`paths.zippedPackage: "${paths.zippedPackage}" → "${newZipped}"`);
    paths.zippedPackage = newZipped;
  } else {
    warn(`paths.zippedPackage not found in ${relPath(PKG_SOL)}`);
  }
 
  if (writeJson(PKG_SOL, data)) {
    ok(`Saved ${relPath(PKG_SOL)}`);
  }
}
 
// ============================================================================
// STEP 3 — {projectName}.manifest.json (auto-discovered under src/)
// ============================================================================
 
function updateManifestFiles(suffix: string, manifestId: string): void {
  info('Step 3 — *.manifest.json files (alias, id)');
 
  const manifests = findFiles(SRC_DIR, (name) => name.endsWith('.manifest.json'));
 
  if (manifests.length === 0) {
    warn(`No *.manifest.json files found under ${relPath(SRC_DIR)}`);
    return;
  }
 
  for (const filePath of manifests) {
    const data = readJson<ManifestJson>(filePath);
    if (!data) continue;
 
    // — alias
    if (typeof data.alias === 'string') {
      const stripped = stripEnvSuffix(data.alias, '', true);
      const newAlias = applyEnvSuffix(stripped, suffix, '', true);
      ok(`${relPath(filePath)} alias: "${data.alias}" → "${newAlias}"`);
      data.alias = newAlias;
    } else {
      warn(`"alias" not found in ${relPath(filePath)}`);
    }
 
    // — id
    ok(`${relPath(filePath)} id: "${data.id}" → "${manifestId}"`);
    data.id = manifestId;
 
    if (writeJson(filePath, data)) {
      ok(`Saved ${relPath(filePath)}`);
    }
  }
}
 
// ============================================================================
// STEP 4 — package.json (name)
// ============================================================================
 
function updatePackageJson(suffix: string): void {
  info('Step 4 — package.json (name)');
 
  const data = readJson<PackageJson>(PKG_JSON);
  if (!data) return;
 
  if (typeof data.name === 'string') {
    const stripped = stripEnvSuffix(data.name, '-');
    const newName  = applyEnvSuffix(stripped, suffix, '-');
    ok(`name: "${data.name}" → "${newName}"`);
    data.name = newName;
    if (writeJson(PKG_JSON, data)) {
      ok(`Saved ${relPath(PKG_JSON)}`);
    }
  } else {
    warn(`"name" not found in ${relPath(PKG_JSON)}`);
  }
}
 
// ============================================================================
// STEP 5 — IExtensibilityLibrary files (ServiceKey, definition names)
// ============================================================================
 
function updateExtensibilityFiles(suffix: string): void {
  info('Step 5 — IExtensibilityLibrary files (ServiceKey, definition names)');
 
  const files = findFiles(SRC_DIR, (name: string, fullPath: string) => {
    if (!name.endsWith('.ts') && !name.endsWith('.tsx')) return false;
    try {
      return fs.readFileSync(fullPath, 'utf8').includes('implements IExtensibilityLibrary');
    } catch {
      return false;
    }
  });
 
  if (files.length === 0) {
    warn(`No files containing "implements IExtensibilityLibrary" found under ${relPath(SRC_DIR)}`);
    return;
  }
 
  for (const filePath of files) {
    let content = readText(filePath);
    if (content === null) continue;
 
    let fileChanged = false;
 
    // ── ServiceKey.create(...) ──────────────────────────────────────────────
    // Strips any existing env suffix then re-applies the new one.
    const serviceKeyRegex = /ServiceKey\.create<[^>]+>\('([^']+?)(?:dev|uat|prod|base)?'/g;
 
    content = content.replace(serviceKeyRegex, (match: string, baseKey: string) => {
      const strippedKey = stripEnvSuffix(baseKey, '', true);
      const newKey      = applyEnvSuffix(strippedKey, suffix, '', true);
      const updated     = match.replace(`'${baseKey}`, `'${newKey}`);
      if (updated !== match) {
        ok(`${relPath(filePath)} ServiceKey: "${baseKey}" → "${newKey}"`);
        fileChanged = true;
      }
      return updated;
    });
 
    // ── ILayoutDefinition / IQueryModifierDefinition / IDataSourceDefinition ──
    // Track brace depth to identify lines inside these definition blocks,
    // then update any name: '...' properties found within them.
 
    const lines: string[] = content.split('\n');
 
    const definitionTypes: string[] = [
      'ILayoutDefinition',
      'IQueryModifierDefinition',
      'IDataSourceDefinition',
    ];
 
    const inDefinitionBlock = new Set<number>();
 
    for (let i = 0; i < lines.length; i++) {
      if (definitionTypes.some((t) => lines[i].includes(t))) {
        inDefinitionBlock.add(i);
        let depth = 0;
        for (let j = i; j < lines.length; j++) {
          for (const ch of lines[j]) {
            if (ch === '{') depth++;
            if (ch === '}') depth--;
          }
          inDefinitionBlock.add(j);
          if (j > i && depth <= 0) break;
        }
      }
    }
 
    // Matches name: 'Some Name (UAT)' or name: 'Some Name' — strips any
    // existing env suffix in parens before re-applying.
    const nameRegex = /(\bname:\s*)(['"])(.+?)(?:\s*\((?:dev|uat|prod|base)\))?\2/g;
 
    const newLines: string[] = lines.map((line: string, idx: number) => {
      if (!inDefinitionBlock.has(idx)) return line;
 
      return line.replace(nameRegex, (match: string, prefix: string, quote: string, baseName: string) => {
        const newName  = suffix ? `${baseName} (${suffix})` : baseName;
        const updated  = `${prefix}${quote}${newName}${quote}`;
        if (updated !== match) {
          ok(`${relPath(filePath)} definition name: "${baseName}" → "${newName}"`);
          fileChanged = true;
        }
        return updated;
      });
    });
 
    if (fileChanged) {
      if (writeText(filePath, newLines.join('\n'))) {
        ok(`Saved ${relPath(filePath)}`);
      }
    } else {
      ok(`${relPath(filePath)} — no changes needed`);
    }
  }
}
 
// ============================================================================
// CLI ENTRY POINT
// ============================================================================
 
function parseArgs(): Environment {
  const args   = process.argv.slice(2);
  const envIdx = args.indexOf('--env');
 
  if (envIdx === -1 || !args[envIdx + 1]) {
    console.error('\n❌  Missing required argument: --env <BASE|DEV|UAT|PROD>\n');
    process.exit(1);
  }
 
  return args[envIdx + 1].toUpperCase() as Environment;
}
 
function main(): void {
  const env = parseArgs();
  console.log(env);
 
  if (!ENV_CONFIG[env]) {
    console.error(
      `\n❌  Unknown environment "${env}". Valid options: ${Object.keys(ENV_CONFIG).join(', ')}\n`
    );
    process.exit(1);
  }
 
  const { suffix, solutionId, manifestId } = ENV_CONFIG[env];
 
  header(`Setting environment: ${env}${!suffix ? ' (reset to base)' : ''}`);
  console.log(`  Suffix   : ${suffix || '(none — restoring base values)'}`);
  console.log(`  Solution : ${solutionId}`);
  console.log(`  Manifest : ${manifestId}`);
 
  updateConfigJson(suffix);
  updatePackageSolution(suffix, solutionId);
  updateManifestFiles(suffix, manifestId);
  updatePackageJson(suffix);
  updateExtensibilityFiles(suffix);
 
  console.log('\n' + '═'.repeat(60));
  if (warnings === 0) {
    console.log('  ✅  All steps completed successfully — 0 warnings.');
  } else {
    console.log(`  ⚠️   Completed with ${warnings} warning(s) — review output above.`);
  }
  console.log('═'.repeat(60) + '\n');
}
 
main();