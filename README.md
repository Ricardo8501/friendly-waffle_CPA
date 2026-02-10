# friendly-waffle_CPA

## Automated RQ1 analysis artifacts

A GitHub Actions workflow runs `analysis/rq1_analysis.py` when a pull request is merged.

- Workflow file: `.github/workflows/rq1-analysis-on-pr-merge.yml`
- Trigger: `pull_request` with `types: [closed]` and merged condition
- Output: analysis CSV/Markdown files and generated PNG visualizations uploaded as workflow artifacts
- Generated files are **not committed** to the repo; download them from the workflow run artifacts.
