
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
                    9. タイトルに巻数を示す数字（3、ローマ数字、日本語の漢数字など）が含まれている場合はお知らせください。
                       数字が含まれているときは、それが本の巻数を正確に示しているかどうかを判定し、巻数であると判断した場合は数字だけを教えてください（例：3）。

                    # Example1
                    <user_query>
                    ちびっ子転生日記帳～お友達いっは?いつくりましゅ!～ THE COMIC 2 (マッグガーデンコミック Beat'sシリーズ)  
                    </user_query>

                    <assistant_response>
                    ちびっ子転生日記帳～お友達いっぱいつくりましゅ!～ THE COMIC
                    2
                    </assistant_response>

                    # Example2

                    <user_query>
                    ミッドナイトレストラン 7to7  
                    </user_query>

                    <assistant_response>
                    ミッドナイトレストラン 7to7
                    0
                    </assistant_response>

                    # Example3

                    <user_query>
                    ながたんと青と-いちかの料理帖-
                    </user_query>

                    <assistant_response>
                    ながたんと青と－いちかの料理帖－  
                    0
                    </assistant_response>

                    # Example4

                    <user_query>
                    おっさん底辺治癒士と愛娘の辺境ライフ～中年男が回復スキルに覚醒して、英雄へ成り上がる～(コミック) :  
                    </user_query>

                    <assistant_response>
                    おっさん底辺治癒士と愛娘の辺境ライフ～中年男が回復スキルに覚醒して、英雄へ成り上がる～ 
                    0
                    </assistant_response>

                    # Example5

                    <user_query>
                    ハボウの轍 4 ~公安調査庁調査官・土師空也~
                    </user_query>

                    <assistant_response>
                    ハボウの轍～公安調査庁調査官・土師空也～
                    4
                    </assistant_response>

                    # Example6

                    <user_query>
                    バリタチNo.1に負けた俺がネコデビューするまで (DAITO COMICS)
                    </user_query>

                    <assistant_response>
                    バリタチNo.1に負けた俺がネコデビューするまで
                    0
                    </assistant_response>

                    # Example7

                    <user_query>
                    私と結婚した事、後悔していませんか?VI (秋水デジタルコミックス)
                    </user_query>

                    <assistant_response>
                    私と結婚した事、後悔していませんか?
                    4
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
                text = response.output_text
                print("response text", text)
                lines = [line.strip() for line in text.split("\n") if line.strip()]
                title = lines[0]
                explanation = lines[1]
                new_item["タイトル"] = title
                if(explanation != "0" ):
                    new_item["巻数"] = explanation
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

def input_json_convert_csv(json_data, csv_path:str):
    import csv
    """Convert JSON data to CSV and save to the specified path."""
    if not json_data:
        logging.warning("No data provided for CSV conversion.")
        return
    try:
        with open(csv_path, "r", encoding="cp932", errors="ignore") as f:
            reader = csv.reader(f)
            rows = list(reader)
        real_data = rows[0:1]  # header row
        for item in json_data:
            row = [""] * len(rows[0])
            row[2] = item.get("タイトル", "") #C column
            row[6] = item.get("巻数", "") #G column
            row[14] = item.get("ASIN", "") #O column

            real_data.append(row)

        with open(csv_path, "w", encoding="cp932", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(real_data)
        return True
    except Exception as exc:
        logging.error(f"Error converting JSON to CSV: {exc}")
        raise RuntimeError(f"Error converting JSON to CSV: {exc}") from exc

# Example usage:
def __main__():
    try:
        edited_data = edit_json_with_openai("./runMacro/target_macro_output.json")
        convert_info = input_json_convert_csv(edited_data, "./public/最終.csv")
        print(convert_info)
    except Exception as e:
        logging.error(f"Error in OpenAI processing: {e}")