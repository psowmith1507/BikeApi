from fastapi import FastAPI
from pydantic import BaseModel, Field
import pandas as pd
import duckdb
import sqlglot
from sqlglot import exp
from pathlib import Path
 
app = FastAPI(title="NL2SQL Execution API")
 
# Excel file location
EXCEL_FILE = Path("Excel_Data.xlsx")
 
# Load Excel data
df = pd.read_excel(EXCEL_FILE)
 
# Create DuckDB connection
connection = duckdb.connect(database=":memory:")
 
# Register Pandas DataFrame as SQL table
connection.register("Excel_Data", df)
 
 
# Request model
class SQLRequest(BaseModel):
    query: str = Field(min_length=1)
 
 
# Health endpoint
@app.get("/health")
def health():
    return {
        "status": "healthy",
        "service": "NL2SQL API"
    }
 
 
# Schema endpoint
@app.get("/api/v1/schema")
def get_schema():
    return {
        "table": "Bike_data",
        "columns": [
            {"name": col, "type": str(dtype)}
            for col, dtype in df.dtypes.items()
        ]
    }
 
 
# SQL validation
def validate_sql(query):
 
    try:
        statements = sqlglot.parse(query, read="duckdb")
 
        if len(statements) != 1 or statements[0] is None:
            return False, "Only one SQL statement is allowed"
 
        statement = statements[0]
 
        # Only SELECT queries are allowed
        if not isinstance(statement, exp.Select):
            return False, "Only SELECT queries are allowed"
 
        # Restrict queries to the employees table
        tables = statement.find_all(exp.Table)
 
        for table in tables:
            if table.name.lower() != "excel_data":
                return False, f"Unknown table: {table.name}"
 
        return True, None
 
    except Exception:
        return False, "Invalid SQL syntax"
 
 
# Main SQL execution endpoint
@app.post("/api/v1/sql/execute")
def execute_sql(request: SQLRequest):
 
    query = request.query.strip()
 
    # Validate SQL
    valid, validation_error = validate_sql(query)
 
    if not valid:
        return {
            "success": False,
            "status": "VALIDATION_ERROR",
            "data": None,
            "error": {
                "code": "INVALID_SQL",
                "message": validation_error
            },
            "retryable": False
        }
 
    try:
        # Execute SQL
        result = connection.execute(query).fetchdf()
 
        # Empty result
        if result.empty:
            return {
                "success": True,
                "status": "EMPTY_RESULT",
                "data": [],
                "error": None,
                "retryable": False
            }
 
        # Convert DataFrame to JSON-compatible records
        records = result.astype(object).where(
            pd.notna(result), None
        ).to_dict(orient="records")
 
        return {
            "success": True,
            "status": "SUCCESS",
            "data": records,
            "error": None,
            "retryable": False
        }
 
    except duckdb.Error as e:
 
        return {
            "success": False,
            "status": "SQL_ERROR",
            "data": None,
            "error": {
                "code": "QUERY_EXECUTION_ERROR",
                "message": str(e)
            },
            "retryable": True
        }
 
    except Exception:
 
        return {
            "success": False,
            "status": "INTERNAL_ERROR",
            "data": None,
            "error": {
                "code": "INTERNAL_SERVER_ERROR",
                "message": "An unexpected server error occurred"
            },
            "retryable": False
        }
