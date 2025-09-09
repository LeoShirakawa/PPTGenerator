import json
import re

def convert_to_slide_data(text_content: str) -> list:
    """
    Converts raw text content into a list of slide data dictionaries
    according to the specified presentation generation rules.
    """
    slide_data = []

    # --- 1. コンテキストの完全分解と正規化 ---
    # 入力前処理（タブ→スペース、連続スペース→1つ、スマートクォート→ASCIIクォート、改行コード→LF、用語統一）
    normalized_text = text_content.replace('\t', ' ')
    normalized_text = re.sub(r' +', ' ', normalized_text) # 連続スペースを1つに
    normalized_text = normalized_text.replace('“', '"').replace('”', '"') # スマートクォートをASCIIクォートに
    normalized_text = normalized_text.replace('‘', "'").replace('’', "'")
    normalized_text = normalized_text.replace('\r\n', '\n').replace('\r', '\n') # 改行コードをLFに統一
    
    # テキストをセクションに分割（簡易版：空行で分割）
    sections = [s.strip() for s in normalized_text.split('\n\n') if s.strip()]

    # --- 2. パターン選定と論理ストーリーの再構築 (簡易版) ---
    # 各セクションをスライドにマッピング

    # タイトルスライド
    title_text = sections[0] if sections else "プレゼンテーションタイトル"
    slide_data.append({
        "type": "title",
        "title": title_text,
        "date": "2025.09.05", # 仮の日付
        "notes": "これは自動生成されたプレゼンテーションの表紙です。"
    })

    # 残りのセクションをコンテンツスライドとして処理
    for i, section_content in enumerate(sections[1:]):
        # 最初の行をスライドタイトル、残りを箇条書きとして扱う
        lines = section_content.split('\n')
        slide_title = lines[0].strip() if lines else f"セクション {i+1}"
        
        # 箇条書きの抽出（ハイフン、アスタリスク、数字リストなど）
        points = []
        for line in lines[1:]:
            line = line.strip()
            if line.startswith('- ') or line.startswith('* '):
                points.append(line[2:].strip())
            elif re.match(r'^\d+\.\s', line): # 数字リスト (e.g., "1. item")
                points.append(re.sub(r'^\d+\.\s', '', line).strip())
            elif line: # その他の行も箇条書きとして追加
                points.append(line)

        # スピーカーノートの生成 (簡易版)
        notes = f"このスライドは「{slide_title}」について説明します。"

        slide_data.append({
            "type": "content",
            "title": slide_title,
            "points": points,
            "notes": notes
        })
    
    # --- 4. JSONオブジェクトの厳密な生成 (dict作成で対応) ---
    # --- 5. 自己検証と反復修正 (今後の課題) ---

    return slide_data