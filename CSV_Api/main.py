import json
from contextlib import asynccontextmanager
 
import duckdb
import pandas as pd
import sqlglot
from sqlglot import exp
 
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
 
 
# Configuration
EXCEL_FILE = "Excel_Data.xlsx"
TABLE_NAME = "Bike_data"
MAX_ROWS = 5000
 
 
# Load Excel file at startup
@asynccontextmanager
async def lifespan(app: FastAPI):
 
    df = pd.read_excel(EXCEL_FILE)
 
    conn = duckdb.connect(
        database=":memory:",
        config={"enable_external_access": "false"}
    )
 
    conn.register("excel_data", df)
 
    conn.execute("""
        CREATE TABLE suppliers AS
        SELECT * FROM excel_data
    """)
 
    app.state.db = conn
 
    print("Excel data loaded successfully!")
    print("Total records:", len(df))
 
    yield
 
    conn.close()
 
 
# FastAPI application
app = FastAPI(
    title="Excel SQL Execution API",
    version="1.0",
    lifespan=lifespan
)
 
 
# Request model
class SQLRequest(BaseModel):
    query: str = Field(min_length=1, max_length=10000)
 
 
# Home endpoint
@app.get("/")
def home():
    return {
        "message": "Excel SQL API is running"
    }
 
 
# SQL execution endpoint
@app.post("/execute-sql")
def execute_sql(request: SQLRequest):
 
    query = request.query.strip()
 
    try:
 
        # 1. Parse SQL query
        statements = sqlglot.parse(
            query,
            read="duckdb"
        )
 
        if len(statements) != 1 or statements[0] is None:
            raise HTTPException(
                status_code=400,
                detail="Only one SQL statement is allowed"
            )
 
        parsed = statements[0]
 
        # 2. Allow SELECT queries only
        if not isinstance(parsed, exp.Select):
            raise HTTPException(
                status_code=400,
                detail="Only SELECT queries are allowed"
            )
 
        # 3. Validate table access
        tables = list(parsed.find_all(exp.Table))
 
        if not tables:
            raise HTTPException(
                status_code=400,
                detail="Query must reference suppliers table"
            )
 
        for table in tables:
 
            if (
                table.name.lower() != TABLE_NAME
                or not isinstance(table.this, exp.Identifier)
                or table.db
                or table.catalog
            ):
                raise HTTPException(
                    status_code=400,
                    detail="Access to this table is not allowed"
                )
 
        # 4. Enforce maximum 5000 records
        limit_node = parsed.args.get("limit")
 
        if limit_node is None:
 
            parsed = parsed.limit(MAX_ROWS)
 
        else:
 
            requested_limit = int(
                limit_node.expression.this
            )
 
            if requested_limit < 0:
                raise HTTPException(
                    status_code=400,
                    detail="LIMIT cannot be negative"
                )
 
            parsed = parsed.limit(
                min(requested_limit, MAX_ROWS)
            )
 
        # 5. Convert parsed SQL to DuckDB SQL
        safe_query = parsed.sql(
            dialect="duckdb"
        )
 
        # 6. Execute SQL query
        result = app.state.db.execute(
            safe_query
        ).fetchdf()
 
        # 7. Convert result into JSON
        records = json.loads(
            result.to_json(
                orient="records",
                date_format="iso"
            )
        )
 
        # 8. Return API response
        return {
            "status": "success",
            "count": len(records),
            "data": records
        }
 
    except HTTPException:
        raise
 
    except Exception as e:
 
        raise HTTPException(
            status_code=400,
            detail=f"SQL execution failed: {str(e)}"
        )
