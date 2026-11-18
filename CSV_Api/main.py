from pathlib import Path
import json
from typing import Any, Optional

import pandas as pd
from fastapi import FastAPI, HTTPException, Query
from pydantic import BaseModel, Field

app = FastAPI(
    title="Bike Purchase Data API",
    description="API for retrieving and analyzing Bike Buyers data",
    version="1.0",
)

# --------------------------------------------------
# 1. DATA SOURCE
# Keep main.py and Excel_Data.xlsx in the same folder.
# --------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "Excel_Data.xlsx"
SHEET_NAME = "bike_buyers raw data"


def load_data() -> pd.DataFrame:
    if not EXCEL_FILE.exists():
        raise HTTPException(
            status_code=404,
            detail=f"Excel file not found: {EXCEL_FILE}",
        )

    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    except ValueError as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Could not find worksheet '{SHEET_NAME}'. Check the Excel sheet name.",
        ) from exc
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Could not read Excel file: {str(exc)}",
        ) from exc


# Convert a Pandas DataFrame into JSON-compatible Python records.
def records(df):
    return json.loads(df.to_json(orient="records"))
# --------------------------------------------------
# 2. HEALTH CHECK
# --------------------------------------------------

@app.get("/")
def home():
    return {"message": "Bike Purchase API is running"}


# --------------------------------------------------
# 3. GET AVAILABLE COLUMNS
# --------------------------------------------------

@app.get("/bikes/columns")
def get_columns():
    df = load_data()
    return {
        "columns": df.columns.tolist(),
        "total_columns": len(df.columns),
    }


# --------------------------------------------------
# 4. GET ALL RECORDS
# --------------------------------------------------

@app.get("/bikes")
def get_bikes(
    limit: int = Query(default=100, ge=1, le=1026),
    skip: int = Query(default=0, ge=0),
):
    df = load_data()
    result = df.iloc[skip : skip + limit]

    return {
        "total_records": len(df),
        "returned_records": len(result),
        "data": records(result),
    }


# --------------------------------------------------
# 5. FILTER RECORDS
# Keep this static route before the /bikes/id route.
# --------------------------------------------------

@app.get("/bikes/filter")
def filter_bikes(
    gender: Optional[str] = None,
    region: Optional[str] = None,
    occupation: Optional[str] = None,
    education: Optional[str] = None,
    purchased_bike: Optional[str] = None,
    home_owner: Optional[str] = None,
    marital_status: Optional[str] = None,
    commute_distance: Optional[str] = None,
    min_income: Optional[float] = None,
    max_income: Optional[float] = None,
    min_age: Optional[int] = None,
    max_age: Optional[int] = None,
):
    df = load_data()

    filters = [
        (gender, "Gender"),
        (region, "Region"),
        (occupation, "Occupation"),
        (education, "Education"),
        (purchased_bike, "Purchased Bike"),
        (home_owner, "Home Owner"),
        (marital_status, "Marital Status"),
        (commute_distance, "Commute Distance"),
    ]

    for value, column in filters:
        if value is not None:
            df = df[
                df[column].astype(str).str.casefold()
                == value.casefold()
            ]

    if min_income is not None:
        df = df[df["Income"] >= min_income]

    if max_income is not None:
        df = df[df["Income"] <= max_income]

    if min_age is not None:
        df = df[df["Age"] >= min_age]

    if max_age is not None:
        df = df[df["Age"] <= max_age]

    return {
        "total_records": len(df),
        "data": records(df),
    }


# --------------------------------------------------
# 6. COUNT BIKE PURCHASES
# --------------------------------------------------

@app.get("/bikes/count")
def count_bikes():
    df = load_data()
    counts = df["Purchased Bike"].value_counts().to_dict()

    return {
        "total_customers": len(df),
        "purchased": int(counts.get("Yes", 0)),
        "not_purchased": int(counts.get("No", 0)),
    }


# --------------------------------------------------
# 7. SUMMARY STATISTICS
# --------------------------------------------------

@app.get("/bikes/summary")
def bike_summary():
    df = load_data()

    return {
        "total_records": len(df),
        "average_income": float(df["Income"].mean()),
        "average_age": float(df["Age"].mean()),
        "average_children": float(df["Children"].mean()),
        "average_cars": float(df["Cars"].mean()),
        "regions": df["Region"].value_counts().to_dict(),
        "bike_purchase": df["Purchased Bike"].value_counts().to_dict(),
    }


# --------------------------------------------------
# 8. GROUP BY
# --------------------------------------------------

@app.get("/bikes/group-by")
def group_bikes(column: str):
    df = load_data()

    if column not in df.columns:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid column name. Available columns: {df.columns.tolist()}",
        )

    result = (
        df.groupby(column)
        .size()
        .reset_index(name="count")
    )

    return {
        "group_by": column,
        "results": records(result),
    }


# --------------------------------------------------
# 9. GET ONE RECORD BY ID
# Use /bikes/id/{bike_id} to avoid conflicts with
# static routes such as /bikes/filter and /bikes/count.
# --------------------------------------------------

@app.get("/bikes/id/{bike_id}")
def get_bike(bike_id: int):
    df = load_data()
    result = df[df["ID"] == bike_id]

    if result.empty:
        raise HTTPException(status_code=404, detail="Bike buyer not found")

    return records(result)[0]


# --------------------------------------------------
# 10. DYNAMIC QUERY API
# --------------------------------------------------

class Filter(BaseModel):
    column: str
    operator: str
    value: Any


class BikeQuery(BaseModel):
    filters: list[Filter] = Field(default_factory=list)
    filter_logic: str = "AND"
    group_by: Optional[str] = None
    aggregations: dict[str, str] = Field(default_factory=dict)
    sort_by: Optional[str] = None
    sort_order: str = "asc"
    limit: int = Field(default=100, ge=1, le=1000)


@app.post("/bikes/query")
def dynamic_query(request: BikeQuery):
    df = load_data()

    allowed_operators = [
        "equals",
        "not_equals",
        "greater_than",
        "less_than",
        "greater_than_or_equal",
        "less_than_or_equal",
        "contains",
        "in",
    ]

    conditions = []

    # Validate and construct each filter condition.
    for item in request.filters:
        column = item.column
        operator = item.operator
        value = item.value

        if column not in df.columns:
            raise HTTPException(status_code=400, detail=f"Invalid column: {column}")

        if operator not in allowed_operators:
            raise HTTPException(status_code=400, detail=f"Invalid operator: {operator}")

        series = df[column]

        try:
            if operator == "equals":
                condition = (
                    series.astype(str).str.casefold()
                    == str(value).casefold()
                )

            elif operator == "not_equals":
                condition = (
                    series.astype(str).str.casefold()
                    != str(value).casefold()
                )

            elif operator == "greater_than":
                condition = series > value

            elif operator == "less_than":
                condition = series < value

            elif operator == "greater_than_or_equal":
                condition = series >= value

            elif operator == "less_than_or_equal":
                condition = series <= value

            elif operator == "contains":
                condition = series.astype(str).str.contains(
                    str(value), case=False, regex=False
                )

            elif operator == "in":
                if not isinstance(value, list):
                    raise HTTPException(
                        status_code=400,
                        detail="'in' operator requires value to be a list",
                    )
                condition = series.isin(value)

            conditions.append(condition)

        except TypeError as exc:
            raise HTTPException(
                status_code=400,
                detail=f"Invalid value or comparison for column '{column}': {value}",
            ) from exc

    # Apply AND / OR logic.
    if conditions:
        logic = request.filter_logic.upper()

        if logic not in ["AND", "OR"]:
            raise HTTPException(status_code=400, detail="Use AND or OR for filter_logic")

        combined = conditions[0]

        for condition in conditions[1:]:
            if logic == "AND":
                combined = combined & condition
            else:
                combined = combined | condition

        df = df[combined]

    # Grouping and aggregations.
    if request.aggregations:
        allowed_aggregations = ["count", "sum", "mean", "min", "max"]

        for column, operation in request.aggregations.items():
            if column not in df.columns:
                raise HTTPException(
                    status_code=400,
                    detail=f"Invalid aggregation column: {column}",
                )

            if operation not in allowed_aggregations:
                raise HTTPException(
                    status_code=400,
                    detail=f"Invalid aggregation: {operation}",
                )

            if operation in ["sum", "mean"] and not pd.api.types.is_numeric_dtype(df[column]):
                raise HTTPException(
                    status_code=400,
                    detail=f"{operation} requires a numeric column: {column}",
                )

        if request.group_by:
            if request.group_by not in df.columns:
                raise HTTPException(
                    status_code=400,
                    detail=f"Invalid group_by column: {request.group_by}",
                )

            grouped = df.groupby(request.group_by).agg(
                request.aggregations
            ).reset_index()

            df = grouped

        else:
            output = {}

            for column, operation in request.aggregations.items():
                if operation == "count":
                    output[f"{column}_count"] = int(df[column].count())
                elif operation == "sum":
                    output[f"{column}_sum"] = float(df[column].sum())
                elif operation == "mean":
                    output[f"{column}_mean"] = float(df[column].mean())
                elif operation == "min":
                    output[f"{column}_min"] = df[column].min()
                elif operation == "max":
                    output[f"{column}_max"] = df[column].max()

            return {
                "total_matching_records": len(df),
                "aggregations": output,
            }

    elif request.group_by:
        if request.group_by not in df.columns:
            raise HTTPException(
                status_code=400,
                detail=f"Invalid group_by column: {request.group_by}",
            )

        df = (
            df.groupby(request.group_by)
            .size()
            .reset_index(name="count")
        )

    # Sorting.
    if request.sort_by:
        if request.sort_by not in df.columns:
            raise HTTPException(
                status_code=400,
                detail=f"Invalid sort column: {request.sort_by}",
            )

        sort_order = request.sort_order.lower()

        if sort_order not in ["asc", "desc"]:
            raise HTTPException(status_code=400, detail="Use asc or desc for sort_order")

        df = df.sort_values(
            by=request.sort_by,
            ascending=(sort_order == "asc"),
        )

    total = len(df)
    df = df.head(request.limit)

    return {
        "total_matching_records": total,
        "returned_records": len(df),
        "data": records(df),
    }
