from google.adk.agents import LlmAgent
from google.adk.models import Gemini
from google.adk.tools import google_search

root_agent = LlmAgent(
    model=Gemini(model="gemini-2.5-flash"),
    name="root_agent",
    instruction="Answer user questions to the best of your knowledge",
    tools=[google_search],
)
