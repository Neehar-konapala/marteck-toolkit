import os
import pandas as pd
def sync_pto_data(downloads_path=r"C:/Users/neeha/Downloads"):
    """Read PTO data from data.xlsx and return its contents."""

    try:
        # Default to user's Downloads folder if not provided
        if not downloads_path:
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")

        # Build full file path
        file_path = os.path.join(downloads_path, "data.xlsx")

        # Check if file exists
        if not os.path.exists(file_path):
            return {
                "status": "error",
                "message": f"File not found: {file_path}"
            }

        # Read first sheet automatically
        df = pd.read_excel(file_path)

        # Prepare response
        return {
            "status": "success",
            "message": "Successfully read data.xlsx",
            "columns": list(df.columns),
            "row_count": len(df),
            "sample_rows": df.head(5).to_dict(orient="records")
        }

    except Exception as e:
        return {
            "status": "error",
            "message": str(e)
        }
