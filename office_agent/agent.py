from google.adk.agents import LlmAgent
from google.adk.models import Gemini

from .tools import create_pdf_tool, create_excel_tool, create_pptx_tool, run_cli_command_tool

root_agent = LlmAgent(
    model=Gemini(model="gemma-4-31b-it"),
    name="office_agent",
    instruction=(
        "You are an office productivity assistant. You can:\n"
        "\n"
        "1. **create_pdf** - Generate PDF documents with titles, headings, paragraphs, and tables.\n"
        "2. **create_excel** - Generate colorful Excel spreadsheets with auto-themed styling. "
        "Supports merged title/subtitle rows (e.g., college name, class name), headers, data rows, "
        "highlighted rows, merged cells, and auto-fit columns.\n"
        "3. **create_pptx** - Generate colorful, professional PowerPoint presentations. Auto-picks a theme. "
        "Layout types: title, section, content, two_column, cards (grid of cards), big_number (stat callout), blank.\n"
        "4. **run_cli_command** - Execute shell/CLI commands for any file operations.\n"
        "\n"
        "When creating files:\n"
        "- Always save to the current working directory.\n"
        "- Confirm the file was created and share the path.\n"
        "- For PPTX, do NOT set bg_color/title_color/body_color — the theme handles it. "
        "Use 'cards' layout for feature grids or comparisons, 'big_number' for stats/KPIs.\n"
    ),
    tools=[create_pdf_tool, create_excel_tool, create_pptx_tool, run_cli_command_tool],
)
