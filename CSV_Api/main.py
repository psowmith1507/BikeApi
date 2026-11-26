
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from contextlib import asynccontextmanager
from pathlib import Path
import pandas as pd
import duckdb
import sqlglot
from sqlglot import exp
import json


# --------------------------------
# 1. DATA SOURCE CONFIGURATION
# --------------------------------

BASE_DIR = Path(__file__).resolve().parent

EXCEL_FILE = BASE_DIR / "Excel_Data.xlsx"

TABLE_NAME = "Bike_data"


# --------------------------------
# 2. LOAD EXCEL DATA
# --------------------------------

@asynccontextmanager
async def lifespan(app: FastAPI):

    if not EXCEL_FILE.exists():
        raise RuntimeError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    df = pd.read_excel(EXCEL_FILE)

    connection = duckdb.connect(
        database=":memory:"
    )

    try:
        connection.register(
            "temp_excel",
            df
        )

        connection.execute(
            f"""
            CREATE TABLE {TABLE_NAME}
            AS SELECT * FROM temp_excel
            """
        )

        connection.unregister("temp_excel")

        app.state.conn = connection

        print("Excel data loaded successfully")
        print("Total records:", len(df))
        print("Table name:", TABLE_NAME)

        yield

    finally:
        connection.close()


# --------------------------------
# 3. FASTAPI APPLICATION
# --------------------------------

app = FastAPI(
    title="Excel SQL Execution API",
    description="Execute SELECT queries on Excel data",
    lifespan=lifespan
)


# --------------------------------
# 4. REQUEST MODEL
# --------------------------------

class SQLRequest(BaseModel):
    query: str


# --------------------------------
# 5. COMMON RESPONSE FORMAT
# --------------------------------

def create_response(
    status,
    message,
    row_count,
    data,
    error
):
    return {
        "status": status,
        "message": message,
        "row_count": row_count,
        "data": data,
        "error": error
    }


# --------------------------------
# 6. SQL EXECUTION API
# --------------------------------

@app.post("/execute-sql")
def execute_sql(
    request: SQLRequest,
    http_request: Request
):

    query = request.query.strip()

    # Empty query validation
    if not query:

        return JSONResponse(
            status_code=400,
            content=create_response(
                "failure",
                "SQL query cannot be empty",
                0,
                None,
                "Empty SQL query"
            )
        )

    try:

        # Parse SQL query
        parsed_queries = sqlglot.parse(
            query,
            read="duckdb"
        )

        # Allow only one SQL statement
        if len(parsed_queries) != 1:

            raise ValueError(
                "Only one SQL statement is allowed"
            )

        parsed_query = parsed_queries[0]

        # Allow only SELECT statements
        if not isinstance(parsed_query, exp.Select):

            raise ValueError(
                "Only SELECT queries are allowed"
            )

        # Execute query
        conn = http_request.app.state.conn

        result = conn.execute(query)

        columns = [
            description[0]
            for description in result.description
        ]

        rows = result.fetchall()

        records = [
            dict(zip(columns, row))
            for row in rows
        ]

        # Convert values into JSON-compatible format
        data = json.loads(
            json.dumps(
                records,
                default=str,
                allow_nan=False
            )
        )

        return create_response(
            "success",
            "Query executed successfully",
            len(data),
            data,
            None
        )

    except Exception as e:

        return JSONResponse(
            status_code=400,
            content=create_response(
                "failure",
                "SQL execution failed",
                0,
                None,
                str(e)
            )
        )


# --------------------------------
# 7. HOME API
# --------------------------------

@app.get("/")
def home():

    return {
        "status": "success",
        "message": "Excel SQL API is running",
        "data_source": EXCEL_FILE.name,
        "table_name": TABLE_NAME
    }
