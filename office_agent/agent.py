from google.adk.agents.llm_agent import Agent
from .tools import create_pdf_tool, create_excel_tool, run_cli_command_tool

root_agent = Agent(
    model="gemma-4-31b-it",
    name="office_agent",
    description=(
        "An office assistant that creates PDF and Excel files, and runs CLI commands."
    ),
    instruction=(
        "You are an office productivity assistant. You can:\n"
        "\n"
        "1. **create_pdf** - Generate PDF documents with titles, headings, paragraphs, and tables.\n"
        "2. **create_excel** - Generate Excel spreadsheets with multiple sheets, headers, rows, and formatting.\n"
        "3. **run_cli_command** - Execute shell/CLI commands for any file operations.\n"
        "\n"
        "When creating files:\n"
        "- Always save to the current working directory.\n"
        "- Confirm the file was created and share the path.\n"
        "- For PDFs, use sections with headings, paragraphs, and optional tables.\n"
        "- For Excel, provide sheet names, headers, and row data.\n"
    ),
    tools=[create_pdf_tool, create_excel_tool, run_cli_command_tool],
)
