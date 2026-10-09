# Task: fix the Recipient Guard GitHub Pages 404

## Context

This folder contains an Outlook add-in ("Recipient Guard") of the Smart Alerts type:

- `manifest.xml`: XML manifest, already installed in Outlook via *Add from File*. All its URLs
  point to `https://amwebexpert.github.io/outlook-recipient-guard/`.
- `commands.html` + `launchevent.js`: runtime for the `OnMessageSend` handler.
- `taskpane.html`: rule configuration task pane (roamingSettings).
- `assets/icon-{16,32,64,80,128}.png`: icons.
- `set-base-url.sh`: rewrites the manifest URLs if the host changes.

These files must be served over HTTPS at the **root** of the Pages site.

## Symptom

`https://amwebexpert.github.io/outlook-recipient-guard/taskpane.html` → **404**.
In Outlook, the add-in icon is broken and the task pane does not open.

## What you need to do

1. **Local Git state**: `git status`, `git remote -v`, `git log --oneline -3`.
   The folder must be a repo with a GitHub remote and the files above committed
   **at the root** (not in a subfolder).
2. **Remote repo**: `gh repo view --json nameWithOwner,visibility,defaultBranchRef,url`.
   - Check the owner. If the repo is under an organization (e.g. `Lichens-Innovation`)
     rather than `amwebexpert`, the Pages URL will be `https://<org>.github.io/outlook-recipient-guard/`.
   - Check that the files are pushed: `gh api repos/{owner}/{repo}/contents --jq '.[].name'`.
3. **GitHub Pages**: `gh api repos/{owner}/{repo}/pages`.
   - If 404 → Pages is not enabled:
     `gh api -X POST repos/{owner}/{repo}/pages -f "source[branch]=main" -f "source[path]=/"`
   - Check that `source.branch` = default branch and `source.path` = `/`.
   - If the repo is **private** on a free plan, Pages is not available: make it public
     (`gh repo edit --visibility public --accept-visibility-change-consequences`) after asking
     me for confirmation. The code contains no sensitive addresses (rules are stored in the
     mailbox, not in the repo).
4. **Build**: `gh api repos/{owner}/{repo}/pages/builds/latest --jq '.status,.error'`.
   Add an empty `.nojekyll` file at the root to avoid any Jekyll processing,
   commit + push, and wait for the build to be `built`.
5. **Verification**: every URL must return 200:
   ```bash
   BASE=https://<owner>.github.io/outlook-recipient-guard
   for f in commands.html taskpane.html launchevent.js assets/icon-16.png assets/icon-32.png \
            assets/icon-64.png assets/icon-80.png assets/icon-128.png; do
     printf "%s %s\n" "$(curl -s -o /dev/null -w '%{http_code}' "$BASE/$f")" "$f"
   done
   ```
6. **If the final URL differs** from `https://amwebexpert.github.io/outlook-recipient-guard`
   (different owner, custom domain): run `./set-base-url.sh <new URL>`, increment
   `<Version>` in `manifest.xml` (e.g. `1.0.0.1`), commit + push, and tell me to remove and
   reinstall the add-in in Outlook with the new `manifest.xml`.

## Constraints

- Do not change the logic of `launchevent.js` or `taskpane.html`.
- Do not rename or move files (the paths are referenced in the manifest).
- Ask for confirmation before any change to the repo's visibility.

## Expected result

A short report: cause of the 404, fix applied, final base URL, output of the `curl` loop
(all 200), and whether the manifest needs to be reinstalled in Outlook.
