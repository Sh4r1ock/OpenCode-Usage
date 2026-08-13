<p align="center">
  <img src="assets/opencode-icon.svg" alt="OpenCode Usage" width="120"/>
</p>

<h1 align="center">OpenCode Usage</h1>

<p align="center">
  See exactly where your OpenCode tokens and money go.
</p>

<p align="center">
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.11%2B-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python 3.11+"/></a>
  <a href="https://www.chartjs.org/"><img src="https://img.shields.io/badge/Chart.js-4.4-FF6384?style=flat-square&logo=chartdotjs&logoColor=white" alt="Chart.js 4.4"/></a>
  <a href="https://playwright.dev/"><img src="https://img.shields.io/badge/Playwright-2EAD33?style=flat-square&logo=playwright&logoColor=white" alt="Playwright"/></a>
</p>

---

OpenCode Usage turns your OpenCode spending into a clean, visual dashboard: token volume, cache hit rate, and cost at a glance, with drill-downs by model and by any time range — from months down to individual minutes.

![OpenCode Usage](assets/OpenCode-Usage.png)

## Features

- **Clear overview** — record count, total tokens, and total cost in cards right at the top
- **Cache insights** — cache hit rate on every bar, plus an overlay line
- **Full time control** — view by month, week, day, hour, or minute, and filter by model
- **Always current** — incremental updates fetch only what's new, and auto-update keeps the dashboard fresh even while you're away
- **Two ways to view** — tokens or cost, light or dark theme, English or Chinese

## Quick Start

```bash
cd opencode_usage
pip install -r requirements.txt
python opencode_usage.py
```

Your browser opens automatically at `http://127.0.0.1:9901`.

## Usage

From the dashboard you can:

- Click **Update** to refresh your usage data
- Open **Settings** to choose your login (GitHub or Google), time range, and update interval
- Enable **Auto Update** to keep data fresh on a schedule, even when the page is closed

To scrape data from the command line instead, run `python scrape_usage.py --help`.

## License

Released under the [MIT License](LICENSE).
