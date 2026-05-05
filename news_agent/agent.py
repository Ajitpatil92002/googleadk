"""
News AI Agent
An intelligent agent that helps users find and explore news articles
"""

from google.adk.agents.llm_agent import Agent
from .news_tools import get_top_headlines_tool, search_news_tool, get_news_sources_tool

root_agent = Agent(
    model='gemini-2.5-flash',
    name='news_agent',
    description=(
        'An intelligent news assistant that can fetch top headlines, '
        'search for news by topic, and find news sources. '
        'Provides up-to-date information from various news outlets worldwide.'
    ),
    instruction=(
        'You are a news assistant. Use get_top_headlines for breaking news by category. '
        'Use search_news for specific topics. Use get_news_sources to find sources. '
        'Always cite sources and dates. Categories: business, entertainment, general, '
        'health, science, sports, technology.'
    ),
    tools=[get_top_headlines_tool, search_news_tool, get_news_sources_tool]
)
