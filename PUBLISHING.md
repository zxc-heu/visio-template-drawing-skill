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

## GitHub Push

After creating an empty GitHub repository:

```powershell
git init
git add .
git commit -m "Initial Visio template drawing skill"
git branch -M main
git remote add origin https://github.com/<github-user>/<repo-name>.git
git push -u origin main
```
