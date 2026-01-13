#!/usr/bin/env python3

import json
import sys
from pathlib import Path
import pandas as pd


def read_excel_file(filename: str, sheet_name: str = None) -> dict:
    try:
        # Resolve Downloads path cross-platform
        downloads_path = Path.home() / "Downloads"
        file_path = downloads_path / filename

        if not file_path.exists():
            return {
                "error": f"File not found: {file_path}",
                "capability": "read_excel_file"
            }

        # Read Excel
        df = pd.read_excel(file_path, sheet_name=sheet_name or 0)

        # Convert dataframe to list-of-dicts for JSON output
        data_records = df.to_dict(orient="records")

        return {
            "result": {
                "file": str(file_path),
                "rows": len(data_records),
                "data": data_records
            },
            "capability": "read_excel_file"
        }

    except Exception as e:
        return {
            "error": str(e),
            "capability": "read_excel_file"
        }


def main():
    try:
        input_data = json.load(sys.stdin)

        capability = input_data.get("capability")
        args = input_data.get("args", {})

        if capability == "read_excel_file":
            result = read_excel_file(
                filename=args.get("filename"),
                sheet_name=args.get("sheet_name")
            )
            print(json.dumps(result, indent=2))
        else:
            print(json.dumps({
                "error": f"Unknown capability: {capability}",
                "capability": capability
            }, indent=2))

    except Exception as e:
        print(json.dumps({
            "error": f"Error: {str(e)}",
            "capability": "unknown"
        }, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
