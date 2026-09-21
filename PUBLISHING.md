# Publishing Checklist

Use this checklist before pushing the skill to GitHub.

## Files to Publish

- `README.md`
- `.gitignore`
- `.gitattributes`
- `PUBLISHING.md`
- `skill/visio-template-drawing/SKILL.md`
- `skill/visio-template-drawing/agents/openai.yaml`
- `skill/visio-template-drawing/scripts/`
- `skill/visio-template-drawing/references/`
- `skill/visio-template-drawing/examples/`

## Files Kept Local

- `outputs/`
- root `examples/`
- root `scripts/`
- `示例/`
- conversation summaries and draft blog notes
- generated `.vsdx` files and image previews

## Before Pushing

1. Confirm no tracked file contains private absolute paths.
2. Run the Python syntax check.
3. Run the PowerShell syntax check.
4. Run layout validation on bundled examples.
5. Run `git status` and verify only public files are staged.
6. Choose a license if the repository will be public and reusable.

## Updating an Existing Repository

1. Read the live remote branch (`git ls-remote origin refs/heads/main`) and compare with the local commit. Do not rely on a cached `origin/main`.
2. Preserve unrelated working-tree changes and untracked files. Review `git diff` and stage named public files only; never use blanket `git add .` for maintenance releases.
3. Include updated lifecycle documentation and generic test scripts/examples, but exclude run registries, logs, diagnostics, screenshots, private inputs, local reports and generated atlases.
4. Run skill validation, both language syntax checks, all bundled layout examples, and the relevant live Visio acceptance suites. Record human concurrency checks separately. Inspect the staged content for private paths and unexpected assets.
5. Commit only the reviewed diff. Fetch/recheck the remote immediately before pushing to the existing correct branch. If it moved, integrate safely; do not overwrite existing work or force-push.
6. Verify the actual remote commit after the push. A transport failure is not a successful release.

Connection overrides for a blocked transport should be scoped to the current command and preserve the repository's configured origin. Keep local installation backup/diff records outside the public package. Synchronize only reviewed files, compare hashes, and preserve installation-only files.
