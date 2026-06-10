# Geo Experiment & Incrementality — A Worked Example

I built this workbook to work through Difference-in-Differences end-to-end on a realistic marketing scenario. Reading about DiD in textbooks never quite clicked for me — what made it stick was sitting with data, designing each step of the analysis myself, and watching the gap between *what platforms claim* and *what's actually incremental* appear on a spreadsheet.

## The scenario

A fictional Online Travel Agency (OTA) spends $10,000 per week on Google Display ads in each of several Vietnamese cities. The Country Manager wants to know whether increasing that spend by 50% (to $15,000/week) will generate enough incremental bookings to justify the cost.

The platform-reported ROAS is famously over-optimistic for display — it credits itself for bookings that would have happened anyway. So the right tool isn't a dashboard; it's a geo experiment. This workbook walks through that experiment end-to-end, from city selection to final recommendation.

## What's inside

[`ota_geo_experiment.xlsx`](https://github.com/TheDataAnalyst-RITU/geo-experiment-incrementality/blob/main/OTA_geo_experiment.xlsx) has six tabs:

**README** — the scenario, design choices, and a list of exercises to work through. Start here when you open the file.

**01_Setup** — the eight Vietnamese cities used in the test, paired and assigned to treatment or control. Also has the global parameters: weekly spend per city, average booking value, commission rate.

**02_Bookings** — the raw weekly bookings data. 80 rows: 8 cities × 10 weeks (4 pre-period + 6 test period). Realistic week-to-week noise so the data doesn't behave like a textbook example.

**03_Spend** — matching weekly ad spend per city, showing exactly when and where the +50% intervention starts.

**04_Analysis** — the worked Difference-in-Differences calculation in six steps:
1. Total bookings per week, by group
2. Average weekly bookings, by group and period
3. Change for each group (test average − pre average)
4. DiD = (treatment change) − (control change) = incremental bookings per week
5. Translation to money: incremental CPA, incremental commission revenue, incremental ROAS
6. Recommendation framing

Every cell uses live formulas pointing back to the data tabs. Change a number in `02_Bookings` and the whole analysis recalculates. The formulas are deliberately inspectable — you can see exactly how SUMIFS, AVERAGEIFS, and the DiD subtraction flow together.

**05_Sanity_Check** — the parallel-trends check on the pre-period. DiD is only valid if treatment and control were moving together before the intervention. This tab confirms they were, by comparing week-over-week growth rates.

## The result (mild spoiler)

When you run through the analysis, the incremental ROAS comes out to roughly **0.55x** — meaning the 50% spend increase does NOT pay for itself on commission revenue. I designed it that way on purpose.

This is the most realistic outcome of a real geo experiment: the platform dashboard would have claimed display was killing it, but the incrementality test reveals that at this particular scale of increase, the extra spend isn't actually working.

The right recommendation isn't "kill display." It's: *don't roll out the +50% — but test a smaller increase (say +20%) to find where incremental ROAS clears the hurdle.* That gap between "what the platform claims" and "what's actually incremental" is the entire reason this kind of testing exists in marketing.

## Why this method matters

Last-click attribution and platform-reported ROAS both *assume* that the spend caused the booking. A geo experiment is one of the few ways to actually *measure* whether it did. For a business spending meaningfully on paid acquisition, the difference between attributed ROAS and incremental ROAS can be huge — sometimes the difference between a profitable channel and a money-losing one.

The Difference-in-Differences method works by comparing the change in the treatment group to the change in the control group, which strips out everything that affected both equally (seasonality, market growth, competitor activity). What's left is the part of the change that's actually attributable to the intervention.

## How to use this workbook

Work through the exercises in the README tab in order — they walk from parallel-trends checking through the final money translation. Try to explain each step out loud, as if presenting to a stakeholder.

If you want to extend it, the most natural additions are: a statistical significance check (the DiD point estimate has uncertainty around it that I didn't quantify), a regression-based version with city and time fixed effects, or a synthetic-control version that blends multiple control regions into one custom counterfactual.

## How I built it

I wanted realistic-looking data to work with — proper week-to-week noise, a believable underlying market lift, a clear treatment effect to recover — without using any real business numbers. So I wrote a small Python script ([`generate_data.py`](https://github.com/TheDataAnalyst-RITU/geo-experiment-incrementality/blob/main/generate_data.py)) to produce synthetic weekly bookings and spend data. AI helped me think through the data generation logic so the numbers would be plausible rather than too clean. The script outputs [`bookings.csv`](https://github.com/TheDataAnalyst-RITU/geo-experiment-incrementality/blob/main/bookings.csv) and [`spend.csv`](https://github.com/TheDataAnalyst-RITU/geo-experiment-incrementality/blob/main/spend.csv), which I imported into Excel.

Everything else — the workbook structure, the sheet design, every formula in the analysis tab, the sanity checks, the money translation, the interpretation, the recommendation — I built directly in Excel. The whole point of the exercise was to do the analysis myself; that's where the actual learning lives.

The Python script uses a fixed random seed, so the data is fully reproducible. Run python [`generate_data.py`](https://github.com/TheDataAnalyst-RITU/geo-experiment-incrementality/blob/main/generate_data.py) and you'll get the same numbers I worked with.

## Caveats

This is a teaching artifact, not a production analysis framework. A few things I deliberately kept simple:

- I used a basic DiD calculation rather than a fixed-effects regression with clustered standard errors. The regression version is what you'd actually run in practice, but it would obscure the underlying logic, which defeats the point.
- I didn't formally test the DiD estimate for statistical significance. With only 4 cities per arm, real power analysis matters a lot — that's a known limitation of small-sample geo tests and the natural next thing to add.
- The "incrementality reveal" treats the DiD answer as the ground truth. In a real setting you'd cross-validate with a second method or a follow-up test before acting on it.

## A note on the scenario

This is a generic OTA scenario. Any resemblance to a specific company's actual operations, spend, or market structure is illustrative only — the numbers are synthetic and chosen for clarity, not pulled from any real business.

## License

Free to use, fork, adapt, or share.
