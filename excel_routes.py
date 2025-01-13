from fastapi import Query, HTTPException, APIRouter
from typing import Optional, Dict, Any, List, Union
import os
from openpyxl import load_workbook  # type: ignore

file_path = "D:/Spreadsheet_Projects/Excel_FastAPI/Employee_Details.xlsx"

excel_router = APIRouter(prefix="/workbook", tags=["Excel_Data"])

@excel_router.get("/fetch-data/", response_model=Union[Dict[str, Any], Dict[str, List[Dict[str, Any]]]])
async def fetch_data(
    sheet_name: str = Query(..., description="Name of the sheet to fetch data from"),
    location: Optional[str] = Query(None, description="Filter by location"),
    department: Optional[str] = Query(None, description="Filter by department"),
    employee_name: Optional[str] = Query(None, description="Filter by employee name"),
) -> Union[Dict[str, Any], Dict[str, List[Dict[str, Any]]]]:
    """
    Fetch details from an Excel spreadsheet.

    This endpoint fetches all rows from the specified sheet in the Excel file.
    Users can optionally filter the results by providing `location`, `department`,
    and/or `employee_name` as query parameters.

    Args:
        sheet_name (str): Name of the sheet to fetch data from.
        location (Optional[str]): Optional filter for the location column.
        department (Optional[str]): Optional filter for the department column.
        employee_name (Optional[str]): Optional filter for the employee name column.

    Returns:
        Union[Dict[str, Any], Dict[str, List[Dict[str, Any]]]]: A dictionary containing
        either an error message or a list of filtered records.
    """
    # Check if the file exists
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Excel file not found. Please check the file path.")

    try:
        # Load the workbook and the specified sheet
        workbook = load_workbook(filename=file_path, data_only=True)
        if sheet_name not in workbook.sheetnames:
            raise HTTPException(status_code=404, detail=f"Sheet '{sheet_name}' not found in the Excel file.")
        worksheet = workbook[sheet_name]
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error reading the Excel file: {str(e)}")

    # Extract headers
    header_row = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True))
    headers = [header.strip() if isinstance(header, str) else header for header in header_row]

    # Extract records
    records = [
        {headers[idx]: value for idx, value in enumerate(row)}
        for row in worksheet.iter_rows(min_row=2, values_only=True)
    ]

    # Filter dynamically based on input
    filters: Dict[str, Optional[str]] = {
        "Location": location,
        "Department": department,
        "Employee Details": employee_name,
    }

    # Apply filters dynamically with partial match support
    filtered_records: List[Dict[str, Any]] = [
        record
        for record in records
        if all(
            value is None
            or (record_value := record.get(key)) is not None
            and str(value).strip().lower() in str(record_value).strip().lower()
            for key, value in filters.items()
        )
    ]

    return {"records": filtered_records}
