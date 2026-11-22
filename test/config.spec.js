import { describe, expect, it } from "vitest";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Guards for CI/configuration regressions that unit tests cannot see:
// a masked "<redacted>" placeholder left in a workflow file, a job-level
// "if" that references the secrets context (which invalidates the whole
// workflow), and the web app access level in appsscript.json being reset
// to "ANYONE" by a redeploy (which silently breaks anonymous access).
const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const workflowsDir = path.join(root, ".github", "workflows");
const workflowFiles = fs
  .readdirSync(workflowsDir)
  .filter((f) => /\.(yml|yaml)$/.test(f));
const read = (file) => fs.readFileSync(path.join(workflowsDir, file), "utf8");

describe("CI configuration", () => {
  it("finds workflow files", () => {
    expect(workflowFiles.length).toBeGreaterThan(0);
  });

  it.each(workflowFiles)(
    "%s does not contain a masked <redacted> placeholder",
    (file) => {
      expect(read(file)).not.toContain("<redacted>");
    },
  );

  it.each(workflowFiles)(
    "%s does not reference secrets in an if condition",
    (file) => {
      const bad = read(file)
        .split("\n")
        .filter((line) => /^\s*if:.*secrets\./.test(line));
      expect(bad).toEqual([]);
    },
  );

  it("keeps anonymous web app access in appsscript.json", () => {
    const manifest = JSON.parse(
      fs.readFileSync(path.join(root, "appsscript.json"), "utf8"),
    );
    if (manifest.webapp) {
      expect(manifest.webapp.access).toBe("ANYONE_ANONYMOUS");
    }
  });
});
