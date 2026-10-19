import { Octokit } from "@octokit/core";
import Handlebars from "handlebars";
import fs from "node:fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const PULL_REQUST_INFO = {
  id: process.env.PULL_REQUEST_ID,
  owner: process.env.GITHUB_REPOSITORY.split("/")[0],
  repoName: process.env.GITHUB_REPOSITORY.split("/")[1],
  username: process.env.GITHUB_ACTOR,
  gitHubToken: process.env.GITHUB_TOKEN,
};

const parseMicrosoftGraphInfo = () => {
  const configJson = process.env.MICROSOFT_GRAPH_INFO_JSON;
  if (!configJson) {
    throw new Error("MICROSOFT_GRAPH_INFO_JSON not found.");
  }

  let parsedConfig;
  try {
    parsedConfig = JSON.parse(configJson);
  } catch {
    throw new Error("MICROSOFT_GRAPH_INFO_JSON is not valid JSON.");
  }

  return {
    tenantId: parsedConfig.tenantId,
    clientId: parsedConfig.clientId,
    clientSecret: parsedConfig.clientSecret,
    siteId: parsedConfig.siteId,
    driveId: parsedConfig.driveId,
    individualWorkbookItemId: parsedConfig.individualWorkbookItemId,
    individualTableName: parsedConfig.individualTableName ?? "CLA_Individual",
    individualColumnName:
      parsedConfig.individualColumnName ?? "GitHub Username",
    corporateWorkbookItemId: parsedConfig.corporateWorkbookItemId,
    corporateTableName: parsedConfig.corporateTableName ?? "CLA_Corporate",
    corporateColumnName: parsedConfig.corporateColumnName ?? "Schedule A",
  };
};

const MICROSOFT_GRAPH_INFO = parseMicrosoftGraphInfo();

const CLA_URL = process.env.CLA_URL;
const PULL_REQUEST_GUIDELINES_URL = process.env.PULL_REQUEST_GUIDELINES_URL;

// Not every repository has a CONTRIBUTORS file; when unset the checklist item is omitted.
const CONTRIBUTORS_URL = process.env.CONTRIBUTORS_URL;

const CLA_LABEL = process.env.CLA_LABEL || "PR - Needs Signed CLA";

const getGraphAccessToken = async () => {
  const tokenUrl = `https://login.microsoftonline.com/${MICROSOFT_GRAPH_INFO.tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: MICROSOFT_GRAPH_INFO.clientId,
    client_secret: MICROSOFT_GRAPH_INFO.clientSecret,
    scope: "https://graph.microsoft.com/.default",
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  if (!response.ok) {
    throw new Error(
      `Failed to obtain Microsoft Graph access token (${response.status}).`,
    );
  }

  const tokenResponse = await response.json();
  return tokenResponse.access_token;
};

const getValuesFromTableColumnValues = async (
  workbookItemId,
  tableName,
  columnName,
) => {
  const accessToken = await getGraphAccessToken();

  const table = encodeURIComponent(tableName);
  const column = encodeURIComponent(columnName);

  const url =
    `https://graph.microsoft.com/v1.0/sites/${MICROSOFT_GRAPH_INFO.siteId}` +
    `/drives/${MICROSOFT_GRAPH_INFO.driveId}` +
    `/items/${workbookItemId}` +
    `/workbook/tables/${table}/columns/${column}/range`;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to read Excel workbook range (${response.status})`);
  }

  const workbookResponse = await response.json();
  return workbookResponse.values ?? [];
};

const checkIfIndividualCLAFound = async () => {
  const rows = await getValuesFromTableColumnValues(
    MICROSOFT_GRAPH_INFO.individualWorkbookItemId,
    MICROSOFT_GRAPH_INFO.individualTableName,
    MICROSOFT_GRAPH_INFO.individualColumnName,
  );

  for (let i = 0; i < rows.length; i++) {
    if (rows[i].length === 0) {
      continue;
    }

    let rowUsername;
    if (rows[i][0] && rows[i][0].length > 0) {
      rowUsername = rows[i][0].toLowerCase();
    }
    if (PULL_REQUST_INFO.username.toLowerCase() === rowUsername) {
      return true;
    }
  }

  return false;
};

const checkIfCorporateCLAFound = async () => {
  const rows = await getValuesFromTableColumnValues(
    MICROSOFT_GRAPH_INFO.corporateWorkbookItemId,
    MICROSOFT_GRAPH_INFO.corporateTableName,
    MICROSOFT_GRAPH_INFO.corporateColumnName,
  );

  for (let i = 0; i < rows.length; i++) {
    if (rows[i].length === 0) {
      continue;
    }

    // We're more lenient with the ScheduleA username check since it's an unformatted text field.
    let rowScheduleA = rows[i][0].toLowerCase();
    rowScheduleA = rowScheduleA.replace(/\n/g, " ");
    const words = rowScheduleA.split(" ");

    for (let j = 0; j < words.length; j++) {
      // Checking for substrings because many GitHub usernames added as "github.com/username".
      if (words[j].includes(PULL_REQUST_INFO.username.toLowerCase())) {
        return true;
      }
    }
  }

  return false;
};

const checkIfUserHasSignedAnyCLA = async () => {
  const foundIndividualCLA = await checkIfIndividualCLAFound();
  if (foundIndividualCLA) {
    return true;
  }

  const foundCorporateCLA = await checkIfCorporateCLAFound();
  return foundCorporateCLA;
};

const getCommentBody = (hasSignedCLA, errorFoundOnCLACheck) => {
  const commentTemplate = fs.readFileSync(
    join(
      dirname(fileURLToPath(import.meta.url)),
      "templates/pullRequestComment.hbs",
    ),
    "utf-8",
  );

  const getCommentFromTemplate = Handlebars.compile(commentTemplate);
  const commentBody = getCommentFromTemplate({
    errorCla: errorFoundOnCLACheck,
    hasCla: hasSignedCLA,
    username: PULL_REQUST_INFO.username,
    claUrl: CLA_URL,
    pullRequestGuidelinesUrl: PULL_REQUEST_GUIDELINES_URL,
    contributorsUrl: CONTRIBUTORS_URL,
  });

  return commentBody;
};

const postCommentOnPullRequest = async (hasSignedCLA, errorFoundOnCLACheck) => {
  const octokit = new Octokit();

  return octokit.request(
    `POST /repos/${PULL_REQUST_INFO.owner}/${PULL_REQUST_INFO.repoName}/issues/${PULL_REQUST_INFO.id}/comments`,
    {
      owner: PULL_REQUST_INFO.username,
      repo: PULL_REQUST_INFO.repoName,
      issue_number: PULL_REQUST_INFO.id,
      body: getCommentBody(hasSignedCLA, errorFoundOnCLACheck),
      headers: {
        authorization: `bearer ${PULL_REQUST_INFO.gitHubToken}`,
        accept: "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
      },
    },
  );
};

const addLabelToPullRequest = async () => {
  const octokit = new Octokit();

  return octokit.request(
    `POST /repos/${PULL_REQUST_INFO.owner}/${PULL_REQUST_INFO.repoName}/issues/${PULL_REQUST_INFO.id}/labels`,
    {
      labels: [CLA_LABEL],
      headers: {
        authorization: `bearer ${PULL_REQUST_INFO.gitHubToken}`,
        accept: "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
      },
    },
  );
};

const main = async () => {
  let hasSignedCLA;
  let errorFoundOnCLACheck;

  if (!CLA_URL || !PULL_REQUEST_GUIDELINES_URL) {
    throw new Error(
      "Both cla-url and pull-request-guidelines-url must be set; the comment is not useful without them.",
    );
  }

  if (
    !MICROSOFT_GRAPH_INFO.tenantId ||
    !MICROSOFT_GRAPH_INFO.clientId ||
    !MICROSOFT_GRAPH_INFO.clientSecret ||
    !MICROSOFT_GRAPH_INFO.siteId ||
    !MICROSOFT_GRAPH_INFO.driveId ||
    !MICROSOFT_GRAPH_INFO.individualWorkbookItemId ||
    !MICROSOFT_GRAPH_INFO.corporateWorkbookItemId
  ) {
    throw new Error(
      "Missing required Microsoft Graph environment variables for CLA lookup.",
    );
  }

  try {
    hasSignedCLA = await checkIfUserHasSignedAnyCLA();
  } catch (error) {
    errorFoundOnCLACheck = error.toString();
  }

  await postCommentOnPullRequest(hasSignedCLA, errorFoundOnCLACheck);

  // A failed lookup is not evidence of a missing CLA, so only label on a confirmed negative.
  if (!hasSignedCLA && !errorFoundOnCLACheck) {
    try {
      await addLabelToPullRequest();
    } catch (error) {
      // The comment carries the actionable information, so labeling is best effort.
      console.log(
        `::warning::Could not apply the "${CLA_LABEL}" label. Create it in this repository to enable labeling. ${error}`,
      );
    }
  }
};

main();
