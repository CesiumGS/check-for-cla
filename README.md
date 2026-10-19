# check-for-cla

A GitHub Action that checks whether a pull request author has a Contributor License Agreement (CLA) on file, comments on the pull request with the result, and applies a `PR - Needs Signed CLA` label when no CLA is found.

CLA records are stored in Excel workbooks in SharePoint and read through the Microsoft Graph API.

## Setting this up in a repository

### 1. Create the PR label

Executing the action puts a label on pull requests with no CLA on file. The label defaults to `PR - Needs Signed CLA`, which can be changed with the `cla-label` input. If the label does not exist the action logs a warning and carries on, so this step is optional — but without it the comment is the only signal.

Create the label through the GitHub CLI or manually:

```sh
gh label create "PR - Needs Signed CLA" --repo CesiumGS/<your-repo>
```

Or add it in the  under **Issues → Labels** in the repository. 

The label name has to match whatever the workflow passes as `cla-label`, since the action applies an existing label rather than creating one.

### 2. Make the credentials available

The action reads CLA records from SharePoint, which requires the `MICROSOFT_GRAPH_INFO_JSON` secret, stored in the GitHub repository's secrets.

### 3. Decide where the comment should link

The comment added to the PR by the GitHub Action points contributors at two places: where to sign a CLA, and the pull request guidelines. Both are required, and the action fails if either is missing — a comment telling someone to sign a CLA without saying where is worse than no comment at all.

Point them at the repository's published documentation in `main`.

To change these without editing the workflow code, store them as repository variables in GitHub under **Settings → Secrets and variables → Actions → Variables** and reference them with `vars`, as shown in the next step.

### 4. Add the workflow

In your repository, create `.github/workflows/cla.yml`:

```yaml
name: check-cla
on:
  pull_request_target:
    types: [opened]

jobs:
  CheckCLA:
    runs-on: ubuntu-latest
    permissions:
      issues: write
      pull-requests: write
    steps:
      - uses: CesiumGS/check-for-cla@v1.0.0
        with:
          github-token: ${{ secrets.GITHUB_TOKEN }}
          pull-request-id: ${{ github.event.number }}
          microsoft-graph-info-json: ${{ secrets.MICROSOFT_GRAPH_INFO_JSON }}
          cla-url: https://github.com/CesiumGS/community/tree/main/CLAs
          pull-request-guidelines-url: https://cesium.com/learn/cesium-native/ref-doc/contributing.html#opening-a-pull-request
```

To drive the links from repository variables instead of literals:

```yaml
          cla-url: ${{ vars.CLA_URL }}
          pull-request-guidelines-url: ${{ vars.PULL_REQUEST_GUIDELINES_URL }}
          cla-label: ${{ vars.CLA_LABEL }}
```

An unset variable resolves to an empty string. For `cla-label` that falls back to the default name, but for the two URLs it fails the run, so make sure the variables exist before switching to this form.

If the repository has a `CONTRIBUTORS` file, add it so contributors are asked to list themselves.

```yaml
          contributors-url: https://github.com/CesiumGS/<your-repo>/blob/main/CONTRIBUTORS.md
```

Note: in the yaml  **`pull_request_target` is required.** The check needs repository secrets, and this trigger supplies them without executing code from the pull request. Switching to `pull_request` breaks the check for forks, which is where it matters most.

### 5. Verify

Open a test pull request from an account known to have signed. Expect a confirmation comment and no label.

Then open one from an account with no CLA on file. Expect a checklist comment and the `PR - Needs Signed CLA` label.

In both cases the job should report success. This action never fails a run for a missing CLA, so a red check means something else went wrong. The common causes are a missing or malformed `MICROSOFT_GRAPH_INFO_JSON`, or missing `issues: write` / `pull-requests: write` permissions. A missing label only produces a warning.

## Inputs

| Input | Required | Description |
| --- | --- | --- |
| `github-token` | yes | Token used to comment and label. Needs `issues: write` and `pull-requests: write`. |
| `pull-request-id` | yes | Number of the pull request to check. |
| `microsoft-graph-info-json` | yes | Microsoft Graph credentials and workbook identifiers. See below. |
| `cla-url` | yes | Where contributors go to sign a CLA. |
| `pull-request-guidelines-url` | yes | Link to the pull request guidelines. |
| `contributors-url` | no | Link to the repository's `CONTRIBUTORS` file. Omit in repositories without one and the checklist item is left out. |
| `cla-label` | no | Label applied when no CLA is found. Defaults to `PR - Needs Signed CLA`. The label must already exist in the repository. |

### `microsoft-graph-info-json`

A JSON object. Table and column names fall back to the defaults shown.

| Key | Required | Default |
| --- | --- | --- |
| `tenantId` | yes | |
| `clientId` | yes | |
| `clientSecret` | yes | |
| `siteId` | yes | |
| `driveId` | yes | |
| `individualWorkbookItemId` | yes | |
| `corporateWorkbookItemId` | yes | |
| `individualTableName` | no | `CLA_Individual` |
| `individualColumnName` | no | `GitHub Username` |
| `corporateTableName` | no | `CLA_Corporate` |
| `corporateColumnName` | no | `Schedule A` |

## Behavior

The individual workbook is matched on exact username. The corporate workbook is matched on substring, because usernames are often recorded there as `github.com/username` inside a free-text field.

| Outcome | Comment | Label |
| --- | --- | --- |
| CLA found | confirmation | no |
| No CLA found | instructions and checklist | yes |
| Lookup failed | error details | no |

A failed lookup is not treated as a missing CLA, so an outage in SharePoint or Microsoft Graph will not label incoming pull requests.

Labeling is best effort. The comment is posted first, and if the label cannot be applied the action logs a warning rather than failing.

The action never fails the job, including when no CLA is found.

## Development

```sh
npm install
npm test
```

`test.mjs` stubs `fetch` to cover the signed, unsigned, and lookup-failure paths without network access.
