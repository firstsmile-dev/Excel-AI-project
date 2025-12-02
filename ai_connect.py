
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
    system_msg = """# Identity
                    あなたは、入力されたテキストからマンガ／ラノベ／書籍タイトルの「正式名称のみ」を抽出し、不要要素を取り除いて整形するアシスタントです。  
                    あなたの目的は、タイトル一覧をクリーンで一貫した形式に統一して出力することです。

                    # Instructions
                    以下のルールに厳密に従って出力してください。

                    1. 入力に含まれる **話数、巻数、出版社名、サブタイトル、記号、エロ・成人タグ、説明文、広告文、シリーズ名以外の情報** をすべて削除してください。  
                    2. 抽出対象は **作品タイトルの正式名称のみ** とします。  
                    3. **同一作品の表記ゆれ**（全角／半角、記号、サブタイトルの有無、略称の違い）は **一つの正式な表記に統一** してください。  
                    4. 出力は **1行につき1タイトル** とします。  
                    5. **タイトル以外の情報を推測して追加してはいけません。**  
                    6. 原作名とシリーズ名の区別が必要な場合は、**シリーズ名を優先** してください。  
                    7. 表記は **日本語のまま、正式名称に統一** してください。  
                    8. **コメントや説明文は一切書かず、タイトルのみを出力** してください。

                    # Example1
                    <user_query>
                    暴食のベルセルク ～俺だけレベルという概念を突破する～ THE COMIC  
                    </user_query>

                    <assistant_response>
                    暴食のベルセルク～俺だけレベルという概念を突破する～ THE COMIC  
                    </assistant_response>

                    # Example2

                    <user_query>
                    ラーメン大好き小泉さん【秋田書店版】  
                    </user_query>

                    <assistant_response>
                    ラーメン大好き小泉さん 
                    </assistant_response>

                    # Example3

                    <user_query>
                    ながたんと青と-いちかの料理帖-
                    </user_query>

                    <assistant_response>
                    ながたんと青と－いちかの料理帖－  
                    </assistant_response>

                    # Example4

                    <user_query>
                    おっさん底辺治癒士と愛娘の辺境ライフ～中年男が回復スキルに覚醒して、英雄へ成り上がる～(コミック) :  
                    </user_query>

                    <assistant_response>
                    おっさん底辺治癒士と愛娘の辺境ライフ～中年男が回復スキルに覚醒して、英雄へ成り上がる～  
                    </assistant_response>

                    # Example5

                    <user_query>
                    JUMBO MAX 
                    </user_query>

                    <assistant_response>
                    JUMBO MAX～ハイパーED薬密造人～
                    </assistant_response>

                    # Context
                    以下にユーザーが未整理のタイトル一覧を入力します。  
                    ルールに従って正式タイトルのみを抽出・整形してください。"""
    edited_data = []
    for item in data:
        user_content = item.get("タイトル")
        color = item.get("color", False)
        new_item = item.copy()
        try:
            if color == True:
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
                new_item["タイトル"] = response.output_text
            else:
                new_item["タイトル"] = user_content
            edited_data.append(new_item)
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