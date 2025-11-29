
"""
ai_connect.py - Uses OpenAI API to process macro output JSON.
Refactored for better error handling, logging, and clarity.
"""

import json
import os
import logging
from typing import Any

from dotenv import load_dotenv
from openai import APIError, AuthenticationError, OpenAI

# Load environment variables from .env file
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)



def edit_json_with_openai(
    json_path: str,
    model: str = "gpt-4.1-mini",
    api_key: str | None = None,
) -> Any:
    """
    Send JSON data to OpenAI for processing and return the edited result.
    Handles API key retrieval, error handling, and logging.
    """
    # Get API key from parameter, .env file, environment variable, or raise error
    if api_key:
        api_key_value = api_key
    else:
        api_key_value = os.getenv("OPENAI_API_KEY")
        if not api_key_value:
            logging.error("OpenAI API key not provided. Set it as a parameter, or set OPENAI_API_KEY in your .env file or environment variable.")
            raise ValueError(
                "OpenAI API key not provided. Set it as a parameter, or set OPENAI_API_KEY in your .env file or environment variable."
            )

    # Initialize OpenAI client
    client = OpenAI(api_key=api_key_value)

    # Load data
    try:
        with open(json_path, "r", encoding="utf-8") as file_handle:
            data = json.load(file_handle)
        logging.info(f"Loaded JSON data from {json_path}")
    except FileNotFoundError as exc:
        logging.error(f"JSON file not found: {json_path}")
        raise FileNotFoundError(f"JSON file not found: {json_path}") from exc
    except json.JSONDecodeError as exc:
        logging.error(f"Invalid JSON in file {json_path}: {exc}")
        raise json.JSONDecodeError(
            f"Invalid JSON in file {json_path}: {exc}", exc.doc, exc.pos
        ) from exc

    # Compose system message and user content
    system_msg = "give me native english"
    edited_data = []
    for item in data:
        user_content = item.get("指定文字列を削除/変換")
        try:
            response = client.responses.create(
                model=model,
                instructions=system_msg,
                input=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "input_text",
                                "text": user_content,
                            }
                        ],
                    }
                ],
            )
            edited_data.append(response.output_text)
        except json.JSONDecodeError as exc:
            logging.error(f"Model did not return valid JSON: {exc}")
            raise ValueError(
                f"Model did not return valid JSON: {exc}"
            ) from exc
        except AuthenticationError as exc:
            logging.error("Invalid OpenAI API key. Please check your API key.")
            raise ValueError("Invalid OpenAI API key. Please check your API key.") from exc
        except APIError as exc:
            logging.error(f"OpenAI API error: {exc}")
            raise RuntimeError(f"OpenAI API error: {exc}") from exc

    return edited_data



# Example usage:
if __name__ == "__main__":
    try:
        edited_data = edit_json_with_openai("./runMacro/target_macro_output.json")
        print(edited_data)
    except Exception as e:
        logging.error(f"Error in OpenAI processing: {e}")