import asyncio
import json
import os
from typing import Optional

from fpdf import FPDF
from google.adk.tools.function_tool import FunctionTool
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


async def create_pdf(
    title: str,
    sections_json: str,
    output_path: str,
    author: Optional[str] = None,
) -> dict:
    """
    Create a PDF document with text, tables, and multiple sections.

    Args:
        title: The document title shown on the first page.
        sections_json: A JSON string representing a list of sections. Each section is an object with:
            - "heading" (string): Section heading text.
            - "paragraphs" (list of strings): Body text paragraphs.
            - "table" (object): Optional table with "headers" (list of strings) and "rows" (list of lists of strings).
            Example: '[{"heading": "Summary", "paragraphs": ["This is a test.", "Another paragraph."], "table": {"headers": ["Name", "Value"], "rows": [["A", "1"], ["B", "2"]]}}]'
        output_path: File path to save the PDF (e.g., 'report.pdf').
        author: Optional author name for the PDF metadata.

    Returns:
        Dictionary with 'status', 'file_path', and 'message'.
    """
    try:
        sections = json.loads(sections_json)
    except json.JSONDecodeError as e:
        return {"status": "error", "message": f"Invalid JSON in sections_json: {e}"}

    try:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)

        if author:
            pdf.set_author(author)
        pdf.set_title(title)

        pdf.add_page()
        pdf.set_font("Helvetica", "B", 24)
        pdf.cell(0, 20, title, ln=True, align="C")
        pdf.ln(10)

        for section in sections:
            heading = section.get("heading", "")
            paragraphs = section.get("paragraphs", [])
            table = section.get("table")

            if heading:
                pdf.set_font("Helvetica", "B", 16)
                pdf.cell(0, 10, heading, ln=True)
                pdf.ln(2)

            if paragraphs:
                pdf.set_font("Helvetica", "", 11)
                for p in paragraphs:
                    pdf.multi_cell(0, 6, str(p))
                    pdf.ln(3)

            if table:
                headers = table.get("headers", [])
                rows = table.get("rows", [])

                if headers:
                    col_count = len(headers)
                    available_width = pdf.w - 20
                    col_width = available_width / col_count

                    pdf.set_font("Helvetica", "B", 10)
                    pdf.set_fill_color(66, 133, 244)
                    pdf.set_text_color(255, 255, 255)
                    for h in headers:
                        pdf.cell(col_width, 8, str(h), border=1, fill=True, align="C")
                    pdf.ln()

                    pdf.set_font("Helvetica", "", 10)
                    pdf.set_text_color(0, 0, 0)
                    for i, row in enumerate(rows):
                        if i % 2 == 0:
                            pdf.set_fill_color(240, 240, 240)
                        else:
                            pdf.set_fill_color(255, 255, 255)
                        for cell in row:
                            pdf.cell(col_width, 7, str(cell), border=1, fill=True)
                    pdf.ln()

                pdf.ln(5)

        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        pdf.output(output_path)
        abs_path = os.path.abspath(output_path)
        return {
            "status": "success",
            "file_path": abs_path,
            "message": f"PDF created successfully at {abs_path}",
        }
    except Exception as e:
        return {"status": "error", "message": f"Failed to create PDF: {str(e)}"}


async def create_excel(
    sheets_json: str,
    output_path: str,
) -> dict:
    """
    Create an Excel spreadsheet with one or more sheets, each containing data with optional formatting.

    Args:
        sheets_json: A JSON string representing a list of sheets. Each sheet is an object with:
            - "name" (string): Sheet tab name.
            - "headers" (list of strings): Column headers.
            - "rows" (list of lists): Row data. Each cell can be a string, number, or formula string starting with '='.
            - "column_widths" (list of integers): Optional column widths.
            - "header_style" (object): Optional header styling with "bold" (bool), "bg_color" (hex string), "font_color" (hex string), "font_size" (int).
            - "auto_fit" (bool): Optional, auto-adjust column widths. Defaults to true.
            Example: '[{"name": "Sales", "headers": ["Product", "Qty", "Price"], "rows": [["Widget", 10, 9.99], ["Gadget", 5, 24.99]]}]'
        output_path: File path to save the Excel file (e.g., 'data.xlsx').

    Returns:
        Dictionary with 'status', 'file_path', and 'message'.
    """
    try:
        sheets = json.loads(sheets_json)
    except json.JSONDecodeError as e:
        return {"status": "error", "message": f"Invalid JSON in sheets_json: {e}"}

    try:
        wb = Workbook()
        first = True

        for sheet_data in sheets:
            name = sheet_data.get("name", "Sheet1")
            headers = sheet_data.get("headers", [])
            rows = sheet_data.get("rows", [])
            column_widths = sheet_data.get("column_widths", [])
            header_style = sheet_data.get("header_style", {})
            auto_fit = sheet_data.get("auto_fit", True)

            if first:
                ws = wb.active
                ws.title = name
                first = False
            else:
                ws = wb.create_sheet(title=name)

            bold = header_style.get("bold", True)
            bg_color = header_style.get("bg_color", "4285F4")
            font_color = header_style.get("font_color", "FFFFFF")
            font_size = header_style.get("font_size", 11)

            header_font = Font(bold=bold, color=font_color, size=font_size)
            header_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            header_alignment = Alignment(horizontal="center", vertical="center")

            if headers:
                for col_idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col_idx, value=header)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_alignment

            for row_idx, row in enumerate(rows, 2):
                for col_idx, value in enumerate(row, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)

            if column_widths:
                for col_idx, width in enumerate(column_widths, 1):
                    ws.column_dimensions[get_column_letter(col_idx)].width = width
            elif auto_fit:
                for col_idx in range(1, max(len(headers), 1) + 1):
                    max_len = 0
                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
                        for cell in row:
                            if cell.value:
                                max_len = max(max_len, len(str(cell.value)))
                    ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 50)

        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        wb.save(output_path)
        abs_path = os.path.abspath(output_path)
        return {
            "status": "success",
            "file_path": abs_path,
            "message": f"Excel file created successfully at {abs_path}",
        }
    except Exception as e:
        return {"status": "error", "message": f"Failed to create Excel file: {str(e)}"}


async def run_cli_command(
    command: str,
    working_dir: Optional[str] = None,
    timeout: int = 60,
) -> dict:
    """
    Execute a CLI command and return the output.

    Args:
        command: The shell command to execute (e.g., 'dir', 'python script.py', 'npm install')
        working_dir: Optional working directory to run the command in. Defaults to current directory.
        timeout: Maximum execution time in seconds. Defaults to 60.

    Returns:
        Dictionary with 'stdout', 'stderr', and 'return_code' keys
    """
    try:
        process = await asyncio.create_subprocess_shell(
            command,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            cwd=working_dir,
        )

        try:
            stdout_bytes, stderr_bytes = await asyncio.wait_for(
                process.communicate(),
                timeout=timeout,
            )
        except asyncio.TimeoutError:
            process.kill()
            await process.communicate()
            return {
                "stdout": "",
                "stderr": f"Command timed out after {timeout} seconds",
                "return_code": "-1",
            }

        stdout = stdout_bytes.decode("utf-8", errors="replace")
        stderr = stderr_bytes.decode("utf-8", errors="replace")

        return {
            "stdout": stdout[:10000],
            "stderr": stderr[:10000],
            "return_code": str(process.returncode),
        }

    except Exception as e:
        return {
            "stdout": "",
            "stderr": f"Failed to execute command: {str(e)}",
            "return_code": "-1",
        }


create_pdf_tool = FunctionTool(create_pdf)
create_excel_tool = FunctionTool(create_excel)
run_cli_command_tool = FunctionTool(run_cli_command)
