import json
from pathlib import Path
from contextlib import asynccontextmanager
 
import pandas as pd
import sqlglot
from sqlglot import exp
 
from fastapi import FastAPI, HTTPException, Query, Body
from pydantic import BaseModel, Field
 
 
# ==========================================
# CONFIGURATION
# ==========================================
 
EXCEL_FILE = Path(__file__).parent / "Excel_Data.xlsx"
 
# This is the SQL table name, not the Excel filename.
TABLE_NAME = "Bike_data"
 
MAX_ROWS = 5000
 
 
# ==========================================
# LOAD EXCEL AT STARTUP
# ==========================================
 
@asynccontextmanager
async def lifespan(app: FastAPI):
 
    if not EXCEL_FILE.exists():
        raise RuntimeError(f"Excel file not found: {EXCEL_FILE}")
 
    df = pd.read_excel(EXCEL_FILE)
 
    df.columns = df.columns.astype(str).str.strip()
 
    if not df.columns.is_unique:
        raise RuntimeError("Excel has duplicate column names")
 
    app.state.df = df
 
    print("Excel loaded successfully")
    print("Total rows:", len(df))
    print("Available columns:", list(df.columns))
 
    yield
 
 
app = FastAPI(
    title="Excel Pandas SQL API",
    version="1.0",
    lifespan=lifespan
)
 
 
# ==========================================
# REQUEST MODEL
# ==========================================
 
class SQLRequest(BaseModel):
    query: str = Field(min_length=1, max_length=10000)
 
 
# ==========================================
# COLUMN VALIDATION
# ==========================================
 
def get_column(name, df):
 
    matches = [
        col for col in df.columns
        if col.lower() == name.lower()
    ]
 
    if not matches:
        raise ValueError(f"Column not found: {name}")
 
    return matches[0]
 
 
# ==========================================
# GET SQL VALUES
# ==========================================
 
def get_value(node, df):
 
    if isinstance(node, exp.Column):
 
        if node.table and node.table.lower() != TABLE_NAME.lower():
            raise ValueError("Invalid table reference")
 
        column = get_column(node.name, df)
 
        return df[column]
 
    if isinstance(node, exp.Literal):
 
        if node.is_string:
            return node.this
 
        try:
            return float(node.this) if "." in node.this else int(node.this)
        except ValueError:
            raise ValueError("Invalid numeric value")
 
    if isinstance(node, exp.Null):
        return None
 
    raise ValueError("Unsupported SQL value")
 
 
# ==========================================
# WHERE CONDITION PROCESSING
# ==========================================
 
def evaluate_where(node, df):
 
    if isinstance(node, exp.Paren):
        return evaluate_where(node.this, df)
 
    if isinstance(node, exp.And):
 
        return (
            evaluate_where(node.this, df)
            & evaluate_where(node.expression, df)
        )
 
    if isinstance(node, exp.Or):
 
        return (
            evaluate_where(node.this, df)
            | evaluate_where(node.expression, df)
        )
 
    if isinstance(node, exp.Not):
 
        return ~evaluate_where(node.this, df)
 
    left = node.this
    right = node.expression
 
    left_value = get_value(left, df)
    right_value = get_value(right, df)
 
    if isinstance(node, exp.EQ):
        return left_value == right_value
 
    if isinstance(node, exp.NEQ):
        return left_value != right_value
 
    if isinstance(node, exp.GT):
        return left_value > right_value
 
    if isinstance(node, exp.GTE):
        return left_value >= right_value
 
    if isinstance(node, exp.LT):
        return left_value < right_value
 
    if isinstance(node, exp.LTE):
        return left_value <= right_value
 
    if isinstance(node, exp.Is):
 
        if isinstance(right, exp.Null):
            return left_value.isna()
 
    if isinstance(node, exp.In):
 
        values = [
            get_value(value, df)
            for value in node.expressions
        ]
 
        return left_value.isin(values)
 
    if isinstance(node, exp.Like):
 
        pattern = str(right_value)
 
        regex = (
            "^"
            + "".join(
                ".*" if char == "%"
                else "." if char == "_"
                else __import__("re").escape(char)
                for char in pattern
            )
            + "$"
        )
 
        return left_value.astype("string").str.match(
            regex,
            na=False
        )
 
    raise ValueError("Unsupported WHERE condition")
 
 
# ==========================================
# AGGREGATE FUNCTIONS
# ==========================================
 
def execute_aggregate(expression, df):
 
    if isinstance(expression, exp.Count):
 
        if isinstance(expression.this, exp.Star):
            return len(df)
 
        return get_value(expression.this, df).count()
 
    if isinstance(expression, exp.Sum):
        return get_value(expression.this, df).sum()
 
    if isinstance(expression, exp.Avg):
        return get_value(expression.this, df).mean()
 
    if isinstance(expression, exp.Min):
        return get_value(expression.this, df).min()
 
    if isinstance(expression, exp.Max):
        return get_value(expression.this, df).max()
 
    raise ValueError("Unsupported aggregate function")
 
 
# ==========================================
# SQL TO PANDAS TRANSLATOR
# ==========================================
 
def execute_pandas_query(query, df):
 
    # Parse SQL
    statements = sqlglot.parse(query, read="duckdb")
 
    if len(statements) != 1 or statements[0] is None:
        raise ValueError("Only one SQL statement is allowed")
 
    parsed = statements[0]
 
    # SELECT only
    if not isinstance(parsed, exp.Select):
        raise ValueError("Only SELECT queries are allowed")
 
    # Reject unsupported SQL features
    unsupported = [
        "joins", "group", "having", "with",
        "distinct", "qualify", "windows"
    ]
 
    for feature in unsupported:
 
        if parsed.args.get(feature):
            raise ValueError(
                f"{feature.upper()} is not supported"
            )
 
    # Validate table
    tables = list(parsed.find_all(exp.Table))
 
    if len(tables) != 1:
        raise ValueError("Only one table is allowed")
 
    table = tables[0]
 
    if (
        table.name.lower() != TABLE_NAME.lower()
        or not isinstance(table.this, exp.Identifier)
        or table.db
        or table.catalog
    ):
        raise ValueError("Invalid table name")
 
    working_df = df.copy()
 
    # --------------------------------------
    # WHERE
    # --------------------------------------
 
    where = parsed.args.get("where")
 
    if where:
 
        mask = evaluate_where(where.this, working_df)
 
        working_df = working_df.loc[
            mask.fillna(False)
        ]
 
    # --------------------------------------
    # SELECT
    # --------------------------------------
 
    select_items = parsed.expressions
 
    aggregate_types = (
        exp.Count,
        exp.Sum,
        exp.Avg,
        exp.Min,
        exp.Max
    )
 
    has_aggregate = any(
        isinstance(item.this if isinstance(item, exp.Alias) else item,
                   aggregate_types)
        for item in select_items
    )
 
    if has_aggregate:
 
        if len(select_items) != 1:
            raise ValueError(
                "Use one aggregate expression per query"
            )
 
        item = select_items[0]
 
        expression = (
            item.this if isinstance(item, exp.Alias)
            else item
        )
 
        alias = (
            item.alias if isinstance(item, exp.Alias)
            else expression.sql()
        )
 
        value = execute_aggregate(expression, working_df)
 
        result = pd.DataFrame([{alias: value}])
 
    else:
 
        columns = []
 
        for item in select_items:
 
            if isinstance(item, exp.Star):
 
                columns.extend(working_df.columns.tolist())
 
            elif isinstance(item, exp.Column):
 
                columns.append(
                    get_column(item.name, working_df)
                )
 
            elif (
                isinstance(item, exp.Alias)
                and isinstance(item.this, exp.Column)
            ):
 
                original = get_column(
                    item.this.name,
                    working_df
                )
 
                columns.append(original)
 
                working_df = working_df.rename(
                    columns={original: item.alias}
                )
 
                columns[-1] = item.alias
 
            else:
                raise ValueError(
                    "Unsupported SELECT expression"
                )
 
        result = working_df[columns].copy()
 
    # --------------------------------------
    # ORDER BY
    # --------------------------------------
 
    order_node = parsed.args.get("order")
 
    if order_node:
 
        sort_columns = []
        ascending = []
 
        for item in order_node.expressions:
 
            if not isinstance(item.this, exp.Column):
                raise ValueError("ORDER BY supports columns only")
 
            column = get_column(
                item.this.name,
                result
            )
 
            sort_columns.append(column)
 
            ascending.append(
                not item.args.get("desc", False)
            )
 
        result = result.sort_values(
            by=sort_columns,
            ascending=ascending
        )
 
    # --------------------------------------
    # LIMIT
    # --------------------------------------
 
    limit_node = parsed.args.get("limit")
 
    if limit_node:
 
        limit_value = limit_node.expression
 
        if not isinstance(limit_value, exp.Literal):
            raise ValueError("LIMIT must be a number")
 
        requested_limit = int(limit_value.this)
 
        if requested_limit < 0:
            raise ValueError("LIMIT cannot be negative")
 
        result = result.head(
            min(requested_limit, MAX_ROWS)
        )
 
    else:
 
        result = result.head(MAX_ROWS)
 
    return result.reset_index(drop=True)
 
 
# ==========================================
# HOME ENDPOINT
# ==========================================
 
@app.get("/")
def home():
 
    return {
        "message": "Excel Pandas SQL API is running"
    }
 
 
# ==========================================
# HEALTH ENDPOINT
# ==========================================
 
@app.get("/health")
def health():
 
    return {
        "status": "healthy",
        "records_loaded": len(app.state.df)
    }
 
 
# ==========================================
# SQL EXECUTION API
# SUPPORTS BOTH GET AND POST
# ==========================================
 
@app.api_route(
    "/execute-sql",
    methods=["GET", "POST"]
)
def execute_sql(
    query: str | None = Query(default=None),
    body: SQLRequest | None = Body(default=None)
):
 
    try:
 
        # Get SQL from GET parameter or POST JSON body
        if query is not None:
            sql_query = query
 
        elif body is not None:
            sql_query = body.query
 
        else:
            raise HTTPException(
                status_code=422,
                detail="SQL query is required"
            )
 
        sql_query = sql_query.strip()
 
        # Execute translated Pandas query
        result = execute_pandas_query(
            sql_query,
