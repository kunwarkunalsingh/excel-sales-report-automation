
from pathlib import Path

import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parent
INPUT_DIR = PROJECT_DIR / "input"

COLUMNS = [
    "order_id",
    "order_date",
    "product",
    "category",
    "region",
    "quantity",
    "unit_price",
]


def main():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)

    january_sales = [
        [1001, "2026-01-05", "Laptop", "Electronics", "North", 2, 650],
        [1002, "2026-01-08", "Mouse", "Accessories", "South", 5, 15],
        [1003, "2026-01-12", "Keyboard", "Accessories", "East", 3, 35],
        [1004, "2026-01-18", "Monitor", "Electronics", "West", 2, 180],
        [1005, "2026-01-22", "Laptop", "Electronics", "North", 1, 650],
    ]

    february_sales = [
        # This repeats an existing order line to test duplicate detection.
        [1002, "2026-01-08", "Mouse", "Accessories", "South", 5, 15],
        [1006, "2026-02-03", "Desk", "Furniture", "West", 2, 120],
        [1007, "2026-02-10", "Chair", "Furniture", "North", 4, 75],
        [1008, "2026-02-14", "Webcam", "Accessories", "East", 0, 45],
        [1009, "2026-02-19", "Headphones", "Accessories", "", 2, 25],
        [1010, "2026-02-25", "Tablet", "Electronics", "South", 1, -200],
    ]

    pd.DataFrame(
        january_sales, columns=COLUMNS
    ).to_excel(INPUT_DIR / "sales_january.xlsx", index=False)

    pd.DataFrame(
        february_sales, columns=COLUMNS
    ).to_excel(INPUT_DIR / "sales_february.xlsx", index=False)

    print(f"Sample Excel files created in: {INPUT_DIR}")


if __name__ == "__main__":
    main()
