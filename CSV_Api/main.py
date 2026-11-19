import json
from pathlib import Path
from contextlib import asynccontextmanager
 
import duckdb
import pandas as pd
import sqlglot
from sqlglot import exp
 
from fastapi import FastAPI, HTTPException, Query, Body
from pydantic import BaseModel, Field
 
 
# ==================================
# CONFIGURATION
# ==================================
 
EXCEL_FILE = "Excel_Data.xlsx"
TABLE_NAME = "Bike_data"
MAX_ROWS = 5000
 
 
# ==================================
# LOAD EXCEL AT STARTUP
# ==================================
 
@asynccontextmanager
async def lifespan(app: FastAPI):
 
    if not EXCEL_FILE.exists():
        raise RuntimeError(f"Excel file not found: {EXCEL_FILE}")
 
    # Read Excel using Pandas
    df = pd.read_excel(EXCEL_FILE)
 
    df.columns = df.columns.astype(str).str.strip()
 
    if not df.columns.is_unique:
        raise RuntimeError("Duplicate Excel column names found")
 
    # Create DuckDB in-memory connection
    conn = duckdb.connect(
        database=":memory:",
        config={"enable_external_access": "false"}
    )
 
    # Register Pandas DataFrame
    conn.register("excel_data", df)
 
    # Create SQL table
    conn.execute(f'''
        CREATE TABLE "{TABLE_NAME}" AS
        SELECT * FROM excel_data
    ''')
 
    app.state.db = conn
 
    print("Excel loaded successfully")
    print("Total records:", len(df))
    print("Available columns:", list(df.columns))
 
    yield
 
    conn.close()
 
 
# ==================================
# FASTAPI APPLICATION
# ==================================
 
app = FastAPI(
    title="Excel SQL Execution API",
    version="2.0",
    lifespan=lifespan
)
 
 
# ==================================
# REQUEST MODEL
# ==================================
 
class SQLRequest(BaseModel):
    query: str = Field(min_length=1, max_length=10000)
 
 
# ==================================
# SQL VALIDATION
# ==================================
 
def validate_query(query):
 
    statements = sqlglot.parse(query, read="duckdb")
 
    if len(statements) != 1 or statements[0] is None:
        raise ValueError("Only one SQL statement is allowed")
 
    parsed = statements[0]
 
    # Allow read-only SELECT statements
    if not isinstance(parsed, exp.Select):
        raise ValueError("Only SELECT queries are allowed")
 
    # Reject commands and unsupported external data access
    forbidden = (
        exp.Command,
        exp.Insert,
        exp.Update,
        exp.Delete,
        exp.Drop,
        exp.Create,
        exp.Alter,
        exp.Copy,
        exp.Merge,
    )
 
    for node in parsed.walk():
        if isinstance(node, forbidden):
            raise ValueError("Only read-only SQL is allowed")
 
    # Permit the actual Excel table and CTE references
    cte_names = {
        cte.alias_or_name.lower()
        for cte in parsed.find_all(exp.CTE)
    }
 
    tables = list(parsed.find_all(exp.Table))
 
    if not tables:
        raise ValueError(
            f"Query must use the {TABLE_NAME} table"
        )
 
    for table in tables:
 
        if table.db or table.catalog:
            raise ValueError("Qualified table access is not allowed")
 
        name = table.name.lower()
 
        if name != TABLE_NAME.lower() and name not in cte_names:
            raise ValueError(
                f"Table not allowed: {table.name}"
            )
 
    return parsed
 
 
# ==================================
# SQL EXECUTION
# ==================================
 
def execute_query(query):
 
    # Parse and validate
    parsed = validate_query(query)
 
    # Convert parsed SQL into DuckDB SQL
    safe_query = parsed.sql(dialect="duckdb")
 
    # Execute query and retrieve at most 5001 records
    cursor = app.state.db.execute(safe_query)
 
    columns = [item[0] for item in cursor.description]
 
    rows = cursor.fetchmany(MAX_ROWS + 1)
 
    truncated = len(rows) > MAX_ROWS
 
    rows = rows[:MAX_ROWS]
 
    # Convert result into Pandas DataFrame
    result_df = pd.DataFrame(rows, columns=columns)
 
    # Convert DataFrame to JSON-compatible records
    records = json.loads(
        result_df.to_json(
            orient="records",
            date_format="iso"
        )
    )
 
    return {
        "status": "success",
        "count": len(records),
        "max_records": MAX_ROWS,
        "truncated": truncated,
        "data": records
    }
 
 
# ==================================
# HOME ENDPOINT
# ==================================
 
@app.get("/")
def home():
 
    return {
        "message": "Excel SQL API is running"
    }
 
 
# ==================================
# HEALTH CHECK
# ==================================
 
@app.get("/health")
def health():
 
    return {
        "status": "healthy",
        "records_loaded": len(app.state.db.execute(
            f'SELECT * FROM "{TABLE_NAME}"'
        ).fetchmany(MAX_ROWS + 1))
    }
 
 
# ==================================
# SQL EXECUTION ENDPOINT
# GET + POST
# ==================================
 
@app.api_route(
    "/execute-sql",
    methods=["GET", "POST"]
)
def execute_sql(
    query: str | None = Query(default=None),
    body: SQLRequest | None = Body(default=None)
):
 
    try:
 
        # Get query from GET parameter
        # or POST JSON body
 
        if query is not None:
            sql_query = query
 
        elif body is not None:
            sql_query = body.query
 
        else:
            raise HTTPException(
                status_code=422,
                detail="SQL query is required"
            )
 
        # Execute SQL
        result = execute_query(sql_query)
 
        return result
 
    except HTTPException:
        raise
 
    except Exception as error:
 
        raise HTTPException(
            status_code=400,
            detail=str(error)
        )
