const assert = require("node:assert/strict");
const { spawnSync } = require("node:child_process");
const { createHash } = require("node:crypto");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { test } = require("node:test");

const project = require("../package.json");
const npmCli = process.env.npm_execpath;
assert.ok(npmCli, "Run these tests with npm test so the same npm CLI is used.");

function fixture(t) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "werkhaus-repair-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  for (const name of ["user.npmrc", "global.npmrc"]) {
    fs.writeFileSync(path.join(directory, name), "");
  }
  const env = {
    ...process.env,
    npm_config_cache: path.join(directory, "cache"),
    npm_config_userconfig: path.join(directory, "user.npmrc"),
    npm_config_globalconfig: path.join(directory, "global.npmrc"),
    npm_config_offline: "true",
    npm_config_audit: "false",
    npm_config_fund: "false",
    npm_config_update_notifier: "false",
    npm_config_ignore_scripts: "false",
    npm_config_bin_links: "true",
    npm_config_package_lock: "true",
    npm_config_dry_run: "false",
    NODE_ENV: "development",
  };
  function npm(args) {
    const result = spawnSync(process.execPath, [npmCli, ...args], {
      cwd: directory,
      env,
      encoding: "utf8",
      timeout: 60000,
    });
    assert.ifError(result.error);
    return result;
  }
  function writeJson(name, value) {
    fs.writeFileSync(path.join(directory, name), JSON.stringify(value, null, 2) + "\n");
  }
  const manifest = {
    name: "dependency-repair-test",
    version: "1.0.0",
    private: true,
    // Execute the real project command, including any future helper scripts.
    scripts: project.scripts,
  };
  const lock = {
    name: manifest.name,
    version: manifest.version,
    lockfileVersion: 3,
    requires: true,
    packages: { "": { name: manifest.name, version: manifest.version } },
  };
  function save() {
    writeJson("package.json", manifest);
    writeJson("package-lock.json", lock);
  }
  function addLockedTool() {
    const source = path.join(directory, "tool-source");
    fs.mkdirSync(source);
    fs.writeFileSync(
      path.join(source, "package.json"),
      JSON.stringify({
        name: "repair-fixture",
        version: "1.0.0",
        bin: { "repair-fixture": "bin.cjs" },
      })
    );
    fs.writeFileSync(
      path.join(source, "bin.cjs"),
      '#!/usr/bin/env node\nconsole.log(require("./package.json").version);\n',
      { mode: 0o755 }
    );
    const packed = npm(["pack", "./tool-source", "--ignore-scripts", "--json"]);
    assert.equal(packed.status, 0, packed.stdout + packed.stderr);
    const archive = JSON.parse(packed.stdout)[0].filename;
    const reference = `file:./${archive}`;
    manifest.dependencies = { "repair-fixture": reference };
    lock.packages[""].dependencies = { ...manifest.dependencies };
    lock.packages["node_modules/repair-fixture"] = {
      version: "1.0.0",
      resolved: reference,
      integrity:
        "sha512-" +
        createHash("sha512")
          .update(fs.readFileSync(path.join(directory, archive)))
          .digest("base64"),
      bin: { "repair-fixture": "bin.cjs" },
    };
    save();
    return reference;
  }
  save();
  return { directory, manifest, npm, writeJson, addLockedTool };
}

function readFiles(directory) {
  return ["package.json", "package-lock.json"].map((name) =>
    fs.readFileSync(path.join(directory, name))
  );
}

test("repair replaces stale dependencies, restores executable bins and preserves locked bytes on repeated runs", (t) => {
  const f = fixture(t);
  f.addLockedTool();
  // Noncanonical whitespace makes lockfile regeneration observable as well.
  fs.appendFileSync(path.join(f.directory, "package-lock.json"), "\n  \n");
  const before = readFiles(f.directory);
  const stale = path.join(f.directory, "node_modules", "obsolete-package");
  fs.mkdirSync(stale, { recursive: true });
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const result = f.npm(["run", "repair:deps"]);
    assert.equal(result.status, 0, result.stdout + result.stderr);
    assert.deepEqual(readFiles(f.directory), before);
    assert.equal(fs.existsSync(stale), false);
    const installed = path.join(f.directory, "node_modules", "repair-fixture", "bin.cjs");
    const bin = path.join(f.directory, "node_modules", ".bin", "repair-fixture");
    fs.accessSync(bin, fs.constants.X_OK);
    const command = spawnSync(process.execPath, [installed], { encoding: "utf8" });
    assert.equal(command.status, 0, command.stderr);
    assert.equal(command.stdout.trim(), "1.0.0");
    // Simulate damaged Mac executable permissions before the next repair.
    fs.chmodSync(installed, 0o644);
  }
});

test("missing lockfile fails without generating a replacement", (t) => {
  const f = fixture(t);
  f.addLockedTool();
  fs.unlinkSync(path.join(f.directory, "package-lock.json"));
  const before = fs.readFileSync(path.join(f.directory, "package.json"));
  const result = f.npm(["run", "repair:deps"]);
  assert.notEqual(result.status, 0, result.stdout + result.stderr);
  assert.match(result.stderr, /npm ci|clean install/i);
  assert.equal(fs.existsSync(path.join(f.directory, "package-lock.json")), false);
  assert.deepEqual(fs.readFileSync(path.join(f.directory, "package.json")), before);
});

test("a package/lock mismatch fails without resolving a new lockfile", (t) => {
  const f = fixture(t);
  const reference = f.addLockedTool();
  f.manifest.dependencies["missing-from-lock"] = reference;
  f.writeJson("package.json", f.manifest);
  const before = readFiles(f.directory);
  const result = f.npm(["run", "repair:deps"]);
  assert.notEqual(result.status, 0, result.stdout + result.stderr);
  assert.match(result.stderr, /in sync|Missing:.*from lock file/i);
  assert.deepEqual(readFiles(f.directory), before);
});

test("installation errors are propagated without replacing the lockfile", (t) => {
  const f = fixture(t);
  f.manifest.scripts = { ...f.manifest.scripts, preinstall: 'node -e "process.exit(17)"' };
  f.writeJson("package.json", f.manifest);
  const before = readFiles(f.directory);
  const result = f.npm(["run", "repair:deps"]);
  assert.equal(result.status, 17, result.stdout + result.stderr);
  assert.deepEqual(readFiles(f.directory), before);
});
