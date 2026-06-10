"""
Generate synthetic weekly bookings and spend data for an OTA geo experiment.

Scenario: a fictional OTA tests a +50% display ad budget increase in 4 Vietnamese
cities (treatment) versus 4 matched cities (control), over a 10-week window
(4 pre-period weeks + 6 test-period weeks).

Outputs:
    bookings.csv — weekly bookings per city
    spend.csv    — weekly ad spend per city

Both files are imported into ota_geo_experiment.xlsx for the full DiD analysis.

The script generates plausible numbers with realistic week-to-week noise so the
data doesn't behave like a textbook example. It is deterministic — same seed
produces the same numbers every run.
"""

import csv
import random

# ----- Cities and their pre-period baselines -----
# Each city is assigned to a treatment/control group and paired with a similar
# city in the opposite group based on historical booking volume.
CITIES = {
    'Da Nang':       ('Treatment', 3200),
    'Hai Phong':     ('Control',   3400),
    'Can Tho':       ('Treatment', 3000),
    'Nha Trang':     ('Control',   3100),
    'Hue':           ('Treatment', 2800),
    'Vung Tau':      ('Control',   3000),
    'Quy Nhon':      ('Treatment', 2800),
    'Buon Ma Thuot': ('Control',   2900),
}

# Pre-period: weeks -4 through -1 (no intervention yet)
# Test period: weeks 1 through 6 (treatment cities get the +50% spend)
WEEKS = [-4, -3, -2, -1, 1, 2, 3, 4, 5, 6]

CONTROL_SPEND_PER_WEEK   = 10_000
TREATMENT_SPEND_PER_WEEK = 15_000   # +50% during test period

# Underlying market factor each week — affects both groups equally.
# This is what DiD subtracts out by comparing changes across groups.
MARKET_FACTOR = {
    -4: 1.000, -3: 1.008, -2: 0.992, -1: 1.017,
    1:  1.045,  2: 1.060,  3: 1.070,  4: 1.085,  5: 1.095,  6: 1.115,
}

# Additional lift applied ONLY to treatment cities during the test period —
# this is the "true" causal effect that the DiD analysis should recover.
TREATMENT_LIFT = {
    -4: 0.000, -3: 0.000, -2: 0.000, -1: 0.000,
    1:  0.045,  2: 0.050,  3: 0.052,  4: 0.058,  5: 0.062,  6: 0.068,
}

NOISE_RANGE = 0.015   # ±1.5% random week-to-week noise per city
SEED = 7


def period_of(week):
    return 'Pre' if week < 0 else 'Test'


def generate_bookings():
    """Return weekly bookings per city as a list of dicts."""
    random.seed(SEED)
    rows = []
    for city, (group, baseline) in CITIES.items():
        for week in WEEKS:
            base = MARKET_FACTOR[week]
            lift = TREATMENT_LIFT[week] if group == 'Treatment' else 0
            noise = random.uniform(-NOISE_RANGE, NOISE_RANGE)
            bookings = round(baseline * (base + lift + noise))
            rows.append({
                'city':     city,
                'group':    group,
                'week':     week,
                'period':   period_of(week),
                'bookings': bookings,
            })
    return rows


def generate_spend():
    """Return weekly ad spend per city as a list of dicts."""
    rows = []
    for city, (group, _) in CITIES.items():
        for week in WEEKS:
            if period_of(week) == 'Pre':
                spend = CONTROL_SPEND_PER_WEEK
            elif group == 'Treatment':
                spend = TREATMENT_SPEND_PER_WEEK
            else:
                spend = CONTROL_SPEND_PER_WEEK
            rows.append({
                'city':   city,
                'group':  group,
                'week':   week,
                'period': period_of(week),
                'spend':  spend,
            })
    return rows


def write_csv(rows, path):
    with open(path, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)


if __name__ == '__main__':
    write_csv(generate_bookings(), 'bookings.csv')
    write_csv(generate_spend(),    'spend.csv')
    print('Generated bookings.csv and spend.csv')
