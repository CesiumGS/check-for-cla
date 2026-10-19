// Exercises main() end-to-end with a stubbed fetch, asserting when the label is applied.
import assert from "node:assert";

const BASE_ENV = {
  GITHUB_REPOSITORY: "CesiumGS/cesium-native",
  GITHUB_ACTOR: "octocat",
  GITHUB_TOKEN: "fake-token",
  PULL_REQUEST_ID: "123",
  MICROSOFT_GRAPH_INFO_JSON: JSON.stringify({
    tenantId: "tenant",
    clientId: "client",
    clientSecret: "secret",
    siteId: "site",
    driveId: "drive",
    individualWorkbookItemId: "individual",
    corporateWorkbookItemId: "corporate",
  }),
  CONTRIBUTORS_URL: "",
  CLA_URL: "https://github.com/CesiumGS/community/tree/main/CLAs",
  PULL_REQUEST_GUIDELINES_URL:
    "https://cesium.com/learn/cesium-native/ref-doc/contributing.html#opening-a-pull-request",
  CLA_LABEL: "",
};

const json = (body, status = 200) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });

function install(scenario, calls) {
  globalThis.fetch = async (input, options) => {
    const url = typeof input === "string" ? input : input.url;
    calls.push({ url, body: options?.body });

    if (url.includes("login.microsoftonline.com")) {
      return json({ access_token: "fake-access-token" });
    }

    if (url.includes("graph.microsoft.com")) {
      if (scenario === "error") {
        return json({ error: "service unavailable" }, 503);
      }
      if (scenario === "signed" && url.includes("/items/individual/")) {
        return json({ values: [["octocat"]] });
      }
      return json({ values: [["someone-else"]] });
    }

    if (url.includes("api.github.com")) {
      if (scenario === "label-missing" && url.endsWith("/labels")) {
        return json({ message: "Not Found" }, 404);
      }
      return json({ id: 1 }, 201);
    }

    throw new Error(`Unexpected request: ${url}`);
  };
}

const rejections = [];
process.on("unhandledRejection", (reason) => rejections.push(reason));

async function run(scenario, env = {}) {
  const calls = [];
  Object.assign(process.env, BASE_ENV, env);
  install(scenario, calls);

  await import(`./index.js?scenario=${scenario}`);
  await new Promise((resolve) => setTimeout(resolve, 50));

  return {
    commented: calls.filter((c) => c.url.endsWith("/comments")),
    labeled: calls.filter((c) => c.url.endsWith("/labels")),
  };
}

const error = await run("error");
assert.strictEqual(error.commented.length, 1, "error: should post one comment");
assert.match(
  error.commented[0].body,
  /error checking the CLA/,
  "error: comment should report the failure",
);
assert.strictEqual(
  error.labeled.length,
  0,
  "error: should NOT label on lookup failure",
);
console.log("PASS  lookup error -> comments, does not label");

const unsigned = await run("unsigned");
assert.strictEqual(unsigned.commented.length, 1, "unsigned: one comment");
assert.strictEqual(unsigned.labeled.length, 1, "unsigned: should label");
assert.match(
  unsigned.labeled[0].body,
  /PR - Needs Signed CLA/,
  "unsigned: correct label name",
);
assert.doesNotMatch(
  unsigned.commented[0].body,
  /contributors/,
  "unsigned: contributors line omitted when CONTRIBUTORS_URL unset",
);
console.log("PASS  unsigned -> comments and labels");

const signed = await run("signed");
assert.strictEqual(signed.commented.length, 1, "signed: one comment");
assert.match(
  signed.commented[0].body,
  /we have a CLA on file/,
  "signed: confirmation comment",
);
assert.strictEqual(signed.labeled.length, 0, "signed: should not label");
console.log("PASS  signed -> comments, does not label");

const labelMissing = await run("label-missing");
assert.strictEqual(
  labelMissing.commented.length,
  1,
  "label missing: comment is still posted",
);
assert.strictEqual(
  labelMissing.labeled.length,
  1,
  "label missing: labeling was attempted",
);
assert.deepStrictEqual(
  rejections,
  [],
  "label missing: must not reject, which would fail the job",
);
console.log("PASS  missing label -> comments, warns, does not fail");

const configuredLinks = await run("configured-links", {
  CLA_URL: "https://github.com/CesiumGS/community/tree/main/CLAs",
  PULL_REQUEST_GUIDELINES_URL:
    "https://cesium.com/learn/cesium-native/ref-doc/contributing.html#opening-a-pull-request",
});
assert.match(
  configuredLinks.commented[0].body,
  /community\/tree\/main\/CLAs/,
  "configured: uses the supplied CLA link",
);
assert.match(
  configuredLinks.commented[0].body,
  /cesium\.com\/learn\/cesium-native\/ref-doc\/contributing\.html#opening-a-pull-request/,
  "configured: uses the supplied guidelines link",
);
console.log("PASS  links configured -> uses the supplied docs URLs");

const customLabel = await run("custom-label", { CLA_LABEL: "needs-cla" });
assert.strictEqual(customLabel.labeled.length, 1, "custom label: should label");
assert.match(
  customLabel.labeled[0].body,
  /needs-cla/,
  "custom label: uses the supplied label name",
);
console.log("PASS  label configured -> applies the supplied label");

// Runs last: it is the only scenario expected to reject.
const missingLinks = await run("missing-links", { CLA_URL: "" });
assert.strictEqual(
  missingLinks.commented.length,
  0,
  "missing links: must not comment with a dead link",
);
assert.strictEqual(rejections.length, 1, "missing links: should fail the job");
assert.match(
  String(rejections[0]),
  /cla-url and pull-request-guidelines-url/,
  "missing links: error names the missing inputs",
);
console.log("PASS  links missing -> fails loudly, posts nothing");

console.log("\nAll checks passed.");
