"""
News API Tools for AI Agent
Provides tools to fetch news from various sources using NewsAPI.org
"""

import os
import httpx
from typing import Optional, Dict, Any
from google.adk.tools.function_tool import FunctionTool


async def get_top_headlines(
    category: Optional[str] = None,
    country: str = "us",
    page_size: int = 10
) -> Dict[str, Any]:
    """
    Fech top headlines from various news sources.

    Args:
        category: Category to filter by (business, entertainment, general, health, science, sports, technology)
        country: Country code for local headlines (default: 'us')
        page_size: Number of results to return (1-100, default: 10)

    Returns:
        Dictionary with articles containing title, description, url, source, publishedAt
    """
    api_key = os.getenv("NEWS_API_KEY")
    if not api_key:
        return {"error": "NEWS_API_KEY not found in environment variables"}

    base_url = "https://newsapi.org/v2/top-headlines"

    params = {
        "apiKey": api_key,
        "country": country,
        "pageSize": min(page_size, 100)
    }

    if category:
        params["category"] = category

    async with httpx.AsyncClient() as client:
        try:
            response = await client.get(base_url, params=params, timeout=10.0)
            response.raise_for_status()
            data = response.json()

            if data.get("status") == "ok":
                return {
                    "status": "success",
                    "total_results": data.get("totalResults", 0),
                    "articles": [
                        {
                            "title": article.get("title"),
                            "description": article.get("description"),
                            "url": article.get("url"),
                            "source": article.get("source", {}).get("name"),
                            "published_at": article.get("publishedAt"),
                            "url_to_image": article.get("urlToImage")
                        }
                        for article in data.get("articles", [])
                    ]
                }
            else:
                return {"error": data.get("message", "Unknown error")}
        except httpx.HTTPError as e:
            return {"error": f"HTTP error occurred: {str(e)}"}
        except Exception as e:
            return {"error": f"Unexpected error: {str(e)}"}


async def search_news(
    query: str,
    from_date: Optional[str] = None,
    to_date: Optional[str] = None,
    language: str = "en",
    sort_by: str = "publishedAt",
    page_size: int = 10
) -> Dict[str, Any]:
    """
    Search for news articles by keyword/topic.

    Args:
        query: Search query (keywords or phrases)
        from_date: Start date (YYYY-MM-DD format)
        to_date: End date (YYYY-MM-DD format)
        language: Language code (default: 'en')
        sort_by: Order of results (publishedAt, popularity, relevancy)
        page_size: Number of results to return (1-100, default: 10)

    Returns:
        Dictionary with matching articles
    """
    api_key = os.getenv("NEWS_API_KEY")
    if not api_key:
        return {"error": "NEWS_API_KEY not found in environment variables"}

    base_url = "https://newsapi.org/v2/everything"

    params = {
        "apiKey": api_key,
        "q": query,
        "language": language,
        "sortBy": sort_by,
        "pageSize": min(page_size, 100)
    }

    if from_date:
        params["from"] = from_date
    if to_date:
        params["to"] = to_date

    async with httpx.AsyncClient() as client:
        try:
            response = await client.get(base_url, params=params, timeout=10.0)
            response.raise_for_status()
            data = response.json()

            if data.get("status") == "ok":
                return {
                    "status": "success",
                    "total_results": data.get("totalResults", 0),
                    "articles": [
                        {
                            "title": article.get("title"),
                            "description": article.get("description"),
                            "url": article.get("url"),
                            "source": article.get("source", {}).get("name"),
                            "published_at": article.get("publishedAt"),
                            "url_to_image": article.get("urlToImage"),
                            "content": article.get("content")
                        }
                        for article in data.get("articles", [])
                    ]
                }
            else:
                return {"error": data.get("message", "Unknown error")}
        except httpx.HTTPError as e:
            return {"error": f"HTTP error occurred: {str(e)}"}
        except Exception as e:
            return {"error": f"Unexpected error: {str(e)}"}


async def get_news_sources(
    category: Optional[str] = None,
    language: str = "en",
    country: str = "us"
) -> Dict[str, Any]:
    """
    Get available news sources.

    Args:
        category: Category to filter by
        language: Language code (default: 'en')
        country: Country code (default: 'us')

    Returns:
        Dictionary with available news sources
    """
    api_key = os.getenv("NEWS_API_KEY")
    if not api_key:
        return {"error": "NEWS_API_KEY not found in environment variables"}

    base_url = "https://newsapi.org/v2/top-headlines/sources"

    params = {
        "apiKey": api_key,
        "language": language,
        "country": country
    }

    if category:
        params["category"] = category

    async with httpx.AsyncClient() as client:
        try:
            response = await client.get(base_url, params=params, timeout=10.0)
            response.raise_for_status()
            data = response.json()

            if data.get("status") == "ok":
                return {
                    "status": "success",
                    "sources": [
                        {
                            "id": source.get("id"),
                            "name": source.get("name"),
                            "description": source.get("description"),
                            "url": source.get("url"),
                            "category": source.get("category"),
                            "language": source.get("language"),
                            "country": source.get("country")
                        }
                        for source in data.get("sources", [])
                    ]
                }
            else:
                return {"error": data.get("message", "Unknown error")}
        except httpx.HTTPError as e:
            return {"error": f"HTTP error occurred: {str(e)}"}
        except Exception as e:
            return {"error": f"Unexpected error: {str(e)}"}


# Create FunctionTool instances
get_top_headlines_tool = FunctionTool(get_top_headlines)
search_news_tool = FunctionTool(search_news)
get_news_sources_tool = FunctionTool(get_news_sources)
