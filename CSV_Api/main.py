
import sqlite3
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

# Initialize FastAPI
app = FastAPI(title="Pandas SQL Query API")

# Excel file location
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "Excel_Data.xlsx"

# Load Excel data into Pandas DataFrame
df = pd.read_excel(EXCEL_FILE)

# API request model
class SQLRequest(BaseModel):
    query: str = Field(
        ...,
        min_length=1,
        description="SQL SELECT query to execute against bikes table"
    )


@app.get("/")
def home():
    return {
        "message": "Pandas SQL API is running"
    }


@app.get("/data")
def get_data():
    """Return first 5 records from DataFrame."""

    return {
        "columns": df.columns.tolist(),
        "data": df.head().to_dict(orient="records")
    }


@app.post("/execute-sql")
def execute_sql(request: SQLRequest):

    query = request.query.strip()

    # Allow only SELECT and WITH queries
    if not query.upper().startswith(("SELECT ", "WITH ")):
        raise HTTPException(
            status_code=400,
            detail="Only SELECT queries are allowed"
        )

    try:
        # Create temporary in-memory SQLite database
        connection = sqlite3.connect(":memory:")

        # Copy Pandas DataFrame into SQLite table
        df.to_sql(
            "bikes",
            connection,
            index=False,
            if_exists="replace"
        )

        # Make database read-only for query execution
        connection.execute("PRAGMA query_only = ON")

        # Execute SQL query
        cursor = connection.execute(query)

        # Get column names
        columns = [
            description[0]
            for description in cursor.description
        ]

        # Fetch matching records
        rows = cursor.fetchall()

        # Convert rows into JSON-friendly dictionaries
        results = [
            dict(zip(columns, row))
            for row in rows
        ]

        connection.close()

        return {
            "status": "success",
            "row_count": len(results),
            "columns": columns,
            "data": results
        }

    except sqlite3.Error as e:
        raise HTTPException(
            status_code=400,
            detail=f"SQL execution error: {str(e)}"
        )

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=str(e)
        )
