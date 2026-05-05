from google.adk.agents import LlmAgent
from google.adk.models import Gemini

from .news_tools import get_top_headlines_tool, search_news_tool, get_news_sources_tool

root_agent = LlmAgent(
    model=Gemini(model="gemini-2.5-flash"),
    name="news_agent",
    instruction=(
        "You are a news assistant. Use get_top_headlines for breaking news by category. "
        "Use search_news for specific topics. Use get_news_sources to find sources. "
        "Always cite sources and dates. Categories: business, entertainment, general, "
        "health, science, sports, technology."
    ),
    tools=[get_top_headlines_tool, search_news_tool, get_news_sources_tool],
)
