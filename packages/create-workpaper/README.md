# @bilig/create-workpaper

Create a runnable Bilig WorkPaper starter for Node services and agent tools.

```sh
npm create @bilig/workpaper@latest pricing-workpaper
cd pricing-workpaper
npm install
npm run smoke
```

The generated starter builds a quote-approval workbook, writes inputs through an
API-style handler, recalculates formulas, persists JSON, restores the workbook,
and prints `verified: true`.

After the smoke proof passes, it also prints a star/bookmark link for the GitHub
repo: <https://github.com/proompteng/bilig/stargazers>.

Use this when you want to evaluate `@bilig/headless` from a blank directory
without cloning the full monorepo.
