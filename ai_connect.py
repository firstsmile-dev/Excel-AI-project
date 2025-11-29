import json
import os
from typing import Any

from dotenv import load_dotenv
from openai import APIError, AuthenticationError, OpenAI

# Load environment variables from .env file
load_dotenv()


def edit_json_with_openai(
    json_path: str,
    model: str = "gpt-4.1-mini",
    api_key: str | None = None,
) -> Any:
    # Get API key from parameter, .env file, environment variable, or raise error
    if api_key:
        api_key_value = api_key
    else:
        # Try to get from environment (loaded from .env file or system environment)
        api_key_value = os.getenv("OPENAI_API_KEY")
        if not api_key_value:
            raise ValueError(
                "OpenAI API key not provided. Set it as a parameter, "
                "or set OPENAI_API_KEY in your .env file or environment variable."
            )

    # Initialize OpenAI client
    client = OpenAI(api_key=api_key_value)

    # Load data
    try:
        with open(json_path, "r", encoding="utf-8") as file_handle:
            data = json.load(file_handle)
    except FileNotFoundError as exc:
        raise FileNotFoundError(f"JSON file not found: {json_path}") from exc
    except json.JSONDecodeError as exc:
        raise json.JSONDecodeError(
            f"Invalid JSON in file {json_path}: {exc}", exc.doc, exc.pos
        ) from exc

    # Compose system message and user content
    system_msg = (
        "give me native english"
    )
    for item in data:
        # Send the JSON itself as input; you can change this string to include instructions
        user_content = item.get("指定文字列を削除/変換")

        try:
            # Call OpenAI Responses API
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

            # The model should return pure JSON text
            edited_data = response.output_text

        except json.JSONDecodeError as exc:
            raise ValueError(
                f"Model did not return valid JSON: {exc}"
            ) from exc
        except AuthenticationError as exc:
            raise ValueError("Invalid OpenAI API key. Please check your API key.") from exc
        except APIError as exc:
            raise RuntimeError(f"OpenAI API error: {exc}") from exc

    return edited_data


# Example usage:
if __name__ == "__main__":
    edited_data = edit_json_with_openai("./runMacro/target_macro_output.json")
    print(edited_data)