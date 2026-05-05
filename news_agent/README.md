# News AI Agent

A powerful AI agent built with Google ADK that fetches and analyzes news from various sources using the News API.

## Setup

### 1. Get News API Key

1. Visit [NewsAPI.org](https://newsapi.org/)
2. Sign up for a free account
3. Get your API key from the dashboard

### 2. Configure Environment

Edit `.env` file and add your News API key:

```env
GOOGLE_GENAI_USE_VERTEXAI=0
GOOGLE_API_KEY=your_google_api_key
NEWS_API_KEY=your_news_api_key_here
```

### 3. Test Your Setup

```bash
cd news_agent
python test_news.py
```

## Running the Agent

```bash
adk dev news_agent
```

## Example Queries

**Get Top Headlines:**
- "What are the top headlines today?"
- "Show me the latest technology news"
- "Get breaking sports news"
- "What's happening in business news?"

**Search News:**
- "Search for news about AI and machine learning"
- "Find articles about climate change from last week"
- "Show me news about Tesla stock"

**Categories Available:**
- business
- entertainment
- general
- health
- science
- sports
- technology

## Tools

1. **get_top_headlines** - Get latest headlines by category
2. **search_news** - Search articles by keyword/topic
3. **get_news_sources** - Discover available news sources
