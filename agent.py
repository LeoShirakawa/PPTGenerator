import os
import requests
from typing import Dict, Any

from google.adk.agents import LlmAgent

# The URL for the new, consolidated Cloud Run service endpoint
CLOUD_RUN_SERVICE_URL = os.getenv("CLOUD_RUN_SERVICE_URL", "https://your-cloud-run-service.run.app/generate_from_text/")

def create_presentation_from_text(text_input: str) -> Dict[str, Any]:
    """
    Calls the backend service to generate a PowerPoint presentation from a text prompt.

    Args:
        text_input: A string containing the topic and content for the presentation.

    Returns:
        A dictionary containing the result from the backend service.
    """
    print(f"INFO: Forwarding text prompt to Cloud Run service: {text_input[:200]}...")

    try:
        payload = {"text": text_input}
        headers = {"Content-Type": "application/json"}
        
        response = requests.post(CLOUD_RUN_SERVICE_URL, json=payload, headers=headers)
        response.raise_for_status()  # Raise an exception for HTTP errors (4xx or 5xx)

        response_data = response.json()
        print(f"INFO: Received successful response from Cloud Run: {response_data}")
        return response_data

    except requests.exceptions.RequestException as e:
        error_detail = e.response.text if e.response else str(e)
        print(f"ERROR: The backend service returned an error. Status: {e.response.status_code if e.response else 'N/A'}. Detail: {error_detail}")
        return {"status": "error", "message": f"The backend service failed to process the request. Detail: {error_detail}"}
    except Exception as e:
        print(f"ERROR: An unexpected error occurred: {e}")
        return {"status": "error", "message": f"An unexpected error occurred: {e}"}


# Define the root agent that uses the simplified tool
root_agent = LlmAgent(
    name="presentation_generator_agent_v3",
    model="gemini-2.5-pro",
    description="A presentation generation agent that takes a text prompt and creates a PowerPoint file via a backend service.",
    instruction="""
    You are a presentation creation assistant.
    When the user provides a topic or text for a presentation, you MUST call the `create_presentation_from_text` tool.
    Pass the user's entire request as the `text_input` argument to the tool.
    After the tool call is finished, present the final result to the user, including the link to the generated file.
    Do not try to answer directly or generate any content yourself. Your only job is to call the tool.
    """,
    tools=[create_presentation_from_text]
)