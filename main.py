import os
import io
import json
import re
from datetime import datetime
from typing import List, Dict, Any, Optional, Union, Literal

from fastapi import FastAPI, HTTPException, Request
from pydantic import BaseModel, ValidationError

# Local imports
import ppt_generator

# Google Cloud clients
from google.cloud import storage
import vertexai
from vertexai.generative_models import GenerativeModel
from google.auth import default


# Logging configuration
import logging
logging.basicConfig(level=os.environ.get("LOG_LEVEL", "INFO").upper())

# --- Configuration ---
BUCKET_NAME = os.getenv("GCS_BUCKET_NAME", "your-gcs-bucket-name")

# --- Pydantic Models for API ---
class TextPayload(BaseModel):
    text: str

# Common properties for all slides
class BaseSlide(BaseModel):
    notes: Optional[str] = None

# Specific slide types based on GooglePatternVer. schema
class TitleSlide(BaseSlide):
    type: Literal["title"]
    title: str
    date: str # YYYY.MM.DD format

class SectionSlide(BaseSlide):
    type: Literal["section"]
    title: str
    sectionNo: Optional[int] = None

class ClosingSlide(BaseSlide):
    type: Literal["closing"]

class ContentSlide(BaseSlide):
    type: Literal["content"]
    title: str
    subhead: Optional[str] = None
    points: Optional[List[str]] = None
    twoColumn: Optional[bool] = None
    columns: Optional[List[List[str]]] = None
    images: Optional[List[Union[str, Dict[str, str]]]] = None # string URL or {url: string, caption: string}

class CompareSlide(BaseSlide):
    type: Literal["compare"]
    title: str
    subhead: Optional[str] = None
    leftTitle: str
    rightTitle: str
    leftItems: List[str]
    rightItems: List[str]
    images: Optional[List[str]] = None

class ProcessSlide(BaseSlide):
    type: Literal["process"]
    title: str
    subhead: Optional[str] = None
    steps: List[str]
    images: Optional[List[str]] = None

class TimelineSlide(BaseSlide):
    type: Literal["timeline"]
    title: str
    subhead: Optional[str] = None
    milestones: List[Dict[str, Any]] # { label: str, date: str, state?: 'done'|'next'|'todo' }
    images: Optional[List[str]] = None

class DiagramSlide(BaseSlide):
    type: Literal["diagram"]
    title: str
    subhead: Optional[str] = None
    lanes: List[Dict[str, Any]] # { title: str, items: List[str] }
    images: Optional[List[str]] = None

class CardsSlide(BaseSlide):
    type: Literal["cards"]
    title: str
    subhead: Optional[str] = None
    columns: Optional[Literal[2, 3]] = None
    items: List[Union[str, Dict[str, str]]] # string or {title: string, desc: string}
    images: Optional[List[str]] = None

class TableSlide(BaseSlide):
    type: Literal["table"]
    title: str
    subhead: Optional[str] = None
    headers: List[str]
    rows: List[List[str]]

class ProgressSlide(BaseSlide):
    type: Literal["progress"]
    title: str
    subhead: Optional[str] = None
    items: List[Dict[str, Any]] # { label: str, percent: number }

# Union of all possible slide types
Slide = Union[
    TitleSlide, SectionSlide, ClosingSlide, ContentSlide, CompareSlide,
    ProcessSlide, TimelineSlide, DiagramSlide, CardsSlide, TableSlide, ProgressSlide
]

class PresentationPayload(BaseModel):
    title: str
    author: str
    slides: List[Slide]

# --- FastAPI App ---
app = FastAPI(
    title="PowerPoint Generation Service",
    description="An API to create PowerPoint presentations from text and upload them to GCS.",
    version="3.0.0"
)

# --- Helper Function to call LLM and generate structured data ---
def generate_structured_data_from_text(text_input: str) -> Dict[str, Any]:
    """Calls the Vertex AI Gemini model to convert raw text into a structured JSON payload."""
    project_id = os.getenv("GOOGLE_CLOUD_PROJECT")
    location = os.getenv("GOOGLE_CLOUD_LOCATION")

    logging.info(f"Initializing Vertex AI for project '{project_id}' in '{location}'...")
    try:
        # Explicitly request the cloud-platform scope to call Vertex AI
        credentials, _ = default(
            scopes=['https://www.googleapis.com/auth/cloud-platform']
        )
        
        vertexai.init(project=project_id, location=location, credentials=credentials)
        model = GenerativeModel("gemini-2.5-pro")
    except Exception as e:
        logging.error(f"Failed to initialize Vertex AI or model: {e}", exc_info=True)
        raise

    prompt = f"""
    Based on the following text, generate a JSON object for a presentation.
    The JSON object must strictly adhere to the following rules and schema definitions.

    

    1. **【ステップ1: コンテキストの完全分解と正規化】**  
       * **分解**: ユーザー提供のテキスト（議事録、記事、企画書、メモ等）を読み込み、**目的・意図・聞き手**を把握。内容を「**章（Chapter）→ 節（Section）→ 要点（Point）**」の階層に内部マッピング。  
       * **正規化**: 入力前処理を自動実行。（タブ→スペース、連続スペース→1つ、スマートクォート→ASCIIクォート、改行コード→LF、用語統一）  
    2. **【ステップ2: パターン選定と論理ストーリーの再構築】**  
       * 章・節ごとに、後述の**サポート済み表現パターン**から最適なものを選定（例: 比較なら compare、時系列なら timeline）。  
       * 聞き手に最適な**説得ライン**（問題解決型、PREP法、時系列など）へ再配列。  
    3. **【ステップ3: スライドタイプへのマッピング】**  
       * ストーリー要素を **Googleパターン・スキーマ**に**最適割当**。  
       * 表紙 → title / 章扉 → section（※**大きな章番号**を描画） / 本文 → content, compare, process, timeline, diagram, cards, table, progress / 結び → closing  
    4. **【ステップ4: オブジェクトの厳密な生成】**  
       * **3.0 スキーマ**と**4.0 ルール**に準拠し、文字列をエスケープ（' → \', \ → \\）して1件ずつ生成。  
       * **インライン強調記法**を使用可：  
         * **太字** → 太字  
         * [[重要語]] → **太字＋Googleブルー**（#4285F4）  
       * **画像URLの抽出**: 入力テキスト内の ![](...png|.jpg|.jpeg|.gif|.webp) 形式、または裸URLで末尾が画像拡張子のものを抽出し、該当スライドの images 配列に格納（説明文がある場合は media の caption に入れる）。  
       * **スピーカーノート生成**: 各スライドの内容に基づき、発表者が話すべき内容の**ドラフトを生成**し、notesプロパティに格納する。  
    5. **【ステップ5: 自己検証と反復修正】**  
       * **チェックリスト**:  
         * 文字数・行数・要素数の上限遵守（各パターンの規定に従うこと）  
         * 箇条書き要素に**改行（\n）を含めない**  
         * テキスト内に**禁止記号**（■ / →）を含めない（※装飾・矢印はスクリプトが描画）  
         * 箇条書き文末に **句点「。」を付けない**（体言止め推奨）  
         * notesプロパティが各スライドに適切に設定されているか確認  
         * title.dateはYYYY.MM.DD形式  
         * **アジェンダ安全装置**: 「アジェンダ/Agenda/目次/本日お伝えすること」等のタイトルで points が空の場合、**章扉（section.title）から自動生成**するため、空配列を返さず **ダミー3点**以上を必ず生成  
    6. **【ステップ6: 最終出力】**  
    * 検証済みオブジェクトを論理順に JSON オブジェクトに格納。

    ## **3.0 slideDataスキーマ定義（GooglePatternVer.+SpeakerNotes）**

    **共通プロパティ**

    * **notes?: string**: すべてのスライドオブジェクトに任意で追加可能。スピーカーノートに設定する発表原稿のドラフト（プレーンテキスト）。

    **スライドタイプ別定義**

    * **タイトル**: {{ type: 'title', title: '...', date: 'YYYY.MM.DD', notes?: '...' }}
    * **章扉**: {{ type: 'section', title: '...', sectionNo?: number, notes?: '...' }} ※sectionNo を指定しない場合は自動連番
    * **クロージング**: {{ type: 'closing', notes?: '...' }}

    **本文パターン（必要に応じて選択）**

    * **content（1カラム/2カラム＋画像＋小見出し）**: {{ type: 'content', title: '...', subhead?: string, points?: string[], twoColumn?: boolean, columns?: [string[], string[]], images?: (string | {{ url: string, caption?: string }} )[], notes?: '...' }}
    * **compare（対比）**: {{ type: 'compare', title: '...', subhead?: string, leftTitle: '...', rightTitle: '...', leftItems: string[], rightItems: string[], images?: string[], notes?: '...' }}
    * **process（手順・工程）**: {{ type: 'process', title: '...', subhead?: string, steps: string[], images?: string[], notes?: '...' }}
    * **timeline（時系列）**: {{ type: 'timeline', title: '...', subhead?: string, milestones: {{ label: string, date: string, state?: 'done'|'next'|'todo' }}[], images?: string[], notes?: '...' }}
    * **diagram（レーン図）**: {{ type: 'diagram', title: '...', subhead?: string, lanes: {{ title: string, items: string[] }}[], images?: string[], notes?: '...' }}
    * **cards（カードグリッド）**: {{ type: 'cards', title: '...', subhead?: string, columns?: 2|3, items: (string | {{ title: string, desc?: string }} )[], images?: string[], notes?: '...' }}
    * **table（表）**: {{ type: 'table', title: '...', subhead?: string, headers: string[], rows: string[][], notes?: '...' }}
    * **progress（進捗）**: {{ type: 'progress', title: '...', subhead?: string, items: {{ label: string, percent: number }}[], notes?: '...' }}

    ## **4.0 COMPOSITION_RULES（GooglePatternVer.） — 美しさと論理性を最大化する絶対規則**

    * **全体構成**:  
      1. title（表紙）  
      2. content（アジェンダ、※章が2つ以上のときのみ）  
      3. section  
      4. 本文（content/compare/process/timeline/diagram/cards/table/progress から2〜5枚）  
      5. （3〜4を章の数だけ繰り返し）  
      6. closing（結び）  
    * **テキスト表現・字数**（最大目安）:  
      * title.title: 全角35文字以内  
      * section.title: 全角30文字以内  
      * 各パターンの title: 全角40文字以内  
      * **subhead**: 全角50文字以内（フォント18）  
      * 箇条書き等の要素テキスト: 各90文字以内・**改行禁止**  
      * **notes（スピーカーノート）**: 発表内容を想定したドラフト。文字数制限は緩やかだが、要点を簡潔に。**プレーンテキスト**とし、強調記法は用いないこと。  
      * **禁止記号**: ■ / → を含めない（矢印や区切りはスクリプトが描画）  
      * 箇条書き文末の句点「。」**禁止**（体言止め推奨）  
      * **インライン強調記法**: **太字** と [[重要語]]（太字＋Googleブルー）を必要箇所に使用

    **【重要】最終出力形式:**
    あなたは、上記のルールとスキーマに従って生成したJSONオブジェクトを、そのまま出力してください。解説・前置き・後書き一切禁止。

    ## 5.0 完璧なJSONの出力例（この形式に厳密に従うこと）

    {{
      "title": "営業部門向け新AIツール導入提案",
      "author": "DX推進室",
      "slides": [
        {{
          "type": "title",
          "title": "営業部門向け 新AIツール導入のご提案",
          "date": "2025.09.10",
          "notes": "本日はお時間をいただきありがとうございます。DX推進室のXXです。本日は、営業部門の業務効率化を実現する新しいAIツールについてご提案します。"
        }},
        {{
          "type": "content",
          "title": "アジェンダ",
          "subhead": "本日お話しする内容",
          "points": [
            "1. 営業部門の現状と[[課題]]",
            "2. 新AIツール「SalesAI」のご紹介",
            "3. 導入効果の試算とロードマップ",
            "4. リスクと対策"
          ],
          "notes": "まず現状の課題を確認し、次に新しいツールの概要、そして具体的な導入効果と今後のスケジュールについてご説明します。"
        }},
        {{
          "type": "section",
          "title": "1. 営業部門の現状と課題",
          "sectionNo": 1,
          "notes": "最初のセクションとして、我々が認識している現状の課題についてご説明します。"
        }},
        {{
          "type": "compare",
          "title": "従来手法 (As-Is) と AI導入後 (To-Be)",
          "subhead": "AIがいかに業務を変革するか",
          "leftTitle": "従来 (As-Is)",
          "rightTitle": "新ツール (To-Be)",
          "leftItems": [
            "**手作業**での議事録作成（1件/60分）",
            "提案書の作成に丸1日",
            "顧客データの分析が[[属人的]]"
          ],
          "rightItems": [
            "商談音声から**自動**で議事録生成（5分）",
            "AIが提案書ドラフトを即時作成",
            "全顧客データをAIが分析し[[インサイト]]を提供"
          ],
          "notes": "ご覧の通り、これまで手作業で行っていた多くの業務が自動化され、営業担当者はより創造的な活動に集中できます。"
        }},
        {{
          "type": "process",
          "title": "導入までの3ステップ",
          "steps": [
            "PoC（概念実証）実施",
            "一部門での先行導入",
            "全部門への本格展開"
          ],
          "notes": "導入はスモールスタートで行います。まずPoCで効果を検証し、次に先行導入、最後に全社展開というステップを踏みます。"
        }},
        {{
          "type": "cards",
          "title": "AIが提供する3つのコア機能",
          "columns": 3,
          "items": [
            {{
              "title": "自動議事録作成",
              "desc": "商談音声を自動でテキスト化・要約"
            }},
            {{
              "title": "提案書ドラフト",
              "desc": "顧客課題に基づき最適な提案書を生成"
            }},
            {{
              "title": "失注分析",
              "desc": "過去の傾向から失注リスクを自動検知"
            }}
          ],
          "notes": "コア機能はこの3点です。議事録作成、提案書ドラフト、そして失注分析です。それぞれが営業活動を強力にサポートします。"
        }},
        {{
          "type": "closing",
          "notes": "ご清聴ありがとうございました。ぜひ前向きなご検討をお願いいたします。"
        }}
      ]
    }}

    ## 6.0 最終出力 (厳守事項)

    **最重要警告:** あなたの応答は、解説やMarkdownの```jsonフェンスを一切含まず、JSONオブジェクトそのもの（`{{` で始まり `}}` で終わる単一の文字列）でなければなりません。この出力は自動パーサー（json.loads）によって直接処理されます。
    
    * **禁止事項:** 「```json」, 「こちらがJSONです」, JSON以外のテキスト。
    * **必須事項:** JSON構文（コンマ、括弧、クォート）の完全な遵守。

    入力テキストに基づいて、上記のルールとスキーマに厳密に従った単一のJSONオブジェクトのみを出力してください。

    Here is the text to convert into a presentation JSON:
    ---
    {text_input}
    ---
    """

    logging.info("Calling LLM to generate slide data...")
    try:
        response = model.generate_content(prompt)
        llm_output_text = response.text
        logging.debug(f"Received raw response from LLM: {llm_output_text}")
    except Exception as e:
        logging.error(f"LLM call failed: {e}", exc_info=True)
        raise

    # Robustly parse the LLM output to extract the JSON payload
    try:
        # First, try to strip markdown code fences if present
        cleaned_llm_output = llm_output_text.strip()
        if cleaned_llm_output.startswith("```json") and cleaned_llm_output.endswith("```"):
            cleaned_llm_output = cleaned_llm_output[len("```json"): -len("```")].strip()
        elif cleaned_llm_output.startswith("```") and cleaned_llm_output.endswith("```"): # Generic code block
            cleaned_llm_output = cleaned_llm_output[len("```"): -len("```")].strip()
        logging.debug(f"Cleaned LLM output: {cleaned_llm_output}")

        # Attempt to parse the cleaned output directly as JSON
        try:
            data = json.loads(cleaned_llm_output)
        except json.JSONDecodeError:
            # If direct parsing fails, try to extract a JSON object using regex
            # Find the first '{' and the last '}'
            first_brace = cleaned_llm_output.find('{')
            last_brace = cleaned_llm_output.rfind('}')

            if first_brace == -1 or last_brace == -1 or first_brace >= last_brace:
                raise ValueError("Could not find a valid JSON object in the LLM response.")
            
            json_str = cleaned_llm_output[first_brace : last_brace + 1]
            logging.debug(f"Extracted JSON string: {json_str}")
            data = json.loads(json_str)

        if 'author' not in data:
            data['author'] = '作成者不明'
            logging.warning("LLM response was missing 'author' field. Using default value.")

        # Validate the structure with Pydantic
        validated_data = PresentationPayload(**data)
        return validated_data.dict()

    except (ValueError, json.JSONDecodeError, ValidationError) as e:
        logging.error(f"Failed to parse or validate JSON from LLM response: {e}", exc_info=True)
        raise ValueError(f"Failed to process LLM response: {llm_output_text}") from e

# --- Main Endpoint --- #
@app.post("/generate_from_text/", summary="Generate a PowerPoint from a text prompt")
async def generate_from_text_endpoint(payload: TextPayload):
    """Receives a text prompt, generates a presentation, and returns the GCS URL."""
    try:
        # 1. Call LLM to get structured data
        structured_data = generate_structured_data_from_text(payload.text)
        
        # 2. Generate the presentation
        logging.info(f"Generating presentation with title: {structured_data.get('title')}")

        # Reconstruct the slide data to match the format ppt_generator expects
        # This now directly uses the 'type' and other properties generated by the LLM
        final_slides_data = []

        today_date_str = datetime.now().strftime("%Y.%m.%d")
        # Add title and closing slides based on LLM's output, and process other slide types
        for slide_data in structured_data.get("slides", []):
            slide_type = slide_data.get("type")
            
            if slide_type == "title":
                slide_data["date"] = today_date_str
            # Handle specific transformations for ppt_generator if needed
            if slide_type == "content":
                # Convert 'content' string to 'points' list if 'points' is not present
                if "content" in slide_data and "points" not in slide_data:
                    slide_data["points"] = slide_data.pop("content").split('\n')
            elif slide_type == "cards":
                # Ensure 'items' are in the correct format for ppt_generator
                if "items" in slide_data:
                    new_items = []
                    for item in slide_data["items"]:
                        if isinstance(item, dict) and "title" in item:
                            new_items.append(item) # Already in {title: string, desc: string} format
                        elif isinstance(item, str):
                            new_items.append({"title": item}) # Convert string to {title: string} format
                    slide_data["items"] = new_items
            
            final_slides_data.append(slide_data)

        presentation_object = ppt_generator.create_presentation(final_slides_data)

        # 3. Upload to GCS
        buffer = io.BytesIO()
        presentation_object.save(buffer)
        buffer.seek(0)

        storage_client = storage.Client()
        bucket = storage_client.bucket(BUCKET_NAME)
        
        presentation_title = structured_data.get("title", "Untitled")
        safe_title = "".join(c for c in presentation_title if c.isalnum() or c in (' ', '_')).rstrip()
        if not safe_title:
            safe_title = "Untitled_Presentation"
        file_name = f"{safe_title.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d%H%M%S')}.pptx"
        
        blob = bucket.blob(file_name)
        
        logging.info(f"Uploading presentation to gs://{BUCKET_NAME}/{file_name}")
        blob.upload_from_string(
            buffer.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

        logging.info(f"File uploaded. Public URL: {blob.public_url}")

        return {
            "status": "success", 
            "message": f"Presentation uploaded to GCS successfully.",
            "file_url": blob.public_url
        }

    except Exception as e:
        logging.error(f"An error occurred in the generation process: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
async def root():
    return {"message": "PowerPoint Generation API v3 is running."}
