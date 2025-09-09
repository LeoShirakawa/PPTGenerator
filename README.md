# PPTGenerator: AI駆動プレゼンテーション生成サービス

## 1. 概要 (Overview)

このプロジェクトは、テキストプロンプトからPowerPointプレゼンテーション（.pptx）を自動生成するサービスです。生成AIモデル（Vertex AI Gemini）を利用してテキストからスライド構成をJSON形式で生成し、バックエンドサービスがそのJSONを基にPowerPointファイルを作成、Google Cloud Storage (GCS) にアップロードします。

ADKエージェントを介して、自然言語でプレゼンテーションの作成を指示できます。

## 2. アーキテクチャ (Architecture)

このシステムは以下のコンポーネントで構成されています。

- **ADK Agent (`agent.py`):** ユーザーとの対話インターフェース。`create_presentation_from_text` ツールを定義し、バックエンドサービスを呼び出します。
- **FastAPI Backend (`main.py`):** Cloud Run上で動作するコアサービス。プレゼンテーション生成プロセス全体を管理します。
- **Vertex AI (Gemini):** 入力されたテキストを、プレゼンテーションの構造化データ（JSON）に変換します。
- **Google Cloud Storage (GCS):** 生成されたPowerPointファイルを保存します。

**処理フロー:**
1.  ユーザーがADK Web UIでプレゼンテーション作成を指示します。
2.  ADKエージェントがリクエストを受け取り、Cloud Run上のFastAPIバックエンドを呼び出します。
3.  FastAPIバックエンドがVertex AI Geminiを呼び出し、テキストからスライド構成のJSONを生成します。
4.  生成されたJSONを基に`.pptx`ファイルを作成します。
5.  完成したファイルをGCSにアップロードし、公開URLを生成します。
6.  ADKエージェントがURLをユーザーに返します。

## 3. 利用手順 (Usage)

### ステップ1: ローカル環境でのセットアップ

1.  **リポジトリをクローンします。**
    ```bash
    git clone https://github.com/LeoShirakawa/PPTGenerator.git
    cd PPTGenerator
    ```

2.  **Python仮想環境を作成し、有効化します。**
    ```bash
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **必要なライブラリをインストールします。**
    ```bash
    pip install -r requirements.txt
    ```

4.  **環境変数を設定します。**
    `.env.example` ファイルをコピーして `.env` ファイルを作成し、お使いの環境に合わせて内容を編集します。
    ```bash
    cp .env.example .env
    ```
    **`.env` ファイルの編集項目:**
    - `GOOGLE_CLOUD_PROJECT`: あなたのGoogle CloudプロジェクトID
    - `GCS_BUCKET_NAME`: 生成したPPTXファイルを保存するGCSバケット名
    - `CLOUD_RUN_SERVICE_URL`: ステップ3でデプロイするCloud RunサービスのエンドポイントURL

### ステップ2: Google Cloudの事前設定

1.  **必要なAPIを有効化します。**
    ```bash
    gcloud services enable run.googleapis.com
    gcloud services enable cloudbuild.googleapis.com
    gcloud services enable storage.googleapis.com
    gcloud services enable aiplatform.googleapis.com
    ```

2.  **GCSバケットを作成します。**
    `uniform` バケットレベルのアクセスを有効にして作成することを推奨します。
    ```bash
    gcloud storage buckets create gs://YOUR_GCS_BUCKET_NAME --project=YOUR_PROJECT_ID --location=YOUR_REGION --uniform-bucket-level-access
    ```
    作成したバケットを一般公開して、生成されたファイルにアクセスできるようにします。
    ```bash
    gcloud storage buckets add-iam-policy-binding gs://YOUR_GCS_BUCKET_NAME --member=allUsers --role=roles/storage.objectViewer
    ```
    *`YOUR_GCS_BUCKET_NAME` は `.env` ファイルに設定したものと同じ名前にしてください。*

### ステップ3: Cloud Runへのデプロイ

1.  **gcloud CLIでGoogle Cloudにログインします。**
    ```bash
    gcloud auth login
    gcloud config set project YOUR_PROJECT_ID
    ```

2.  **ソースコードから直接Cloud Runにデプロイします。**
    以下のコマンドは、ソースコードのビルドとデプロイを一度に実行します。
    ```bash
    gcloud run deploy ppt-generator-service \
      --source . \
      --platform managed \
      --region YOUR_REGION \
      --allow-unauthenticated \
      --set-env-vars="GCS_BUCKET_NAME=YOUR_GCS_BUCKET_NAME,GOOGLE_CLOUD_PROJECT=YOUR_PROJECT_ID,GOOGLE_CLOUD_LOCATION=YOUR_REGION,GOOGLE_GENAI_USE_VERTEXAI=TRUE"
    ```
    - `YOUR_PROJECT_ID`: あなたのGoogle CloudプロジェクトID
    - `YOUR_REGION`: デプロイするリージョン (例: `us-central1`)
    - `YOUR_GCS_BUCKET_NAME`: ステップ2で作成したGCSバケット名

    デプロイが完了すると、サービスのURLが表示されます。このURLをコピーし、`.env` ファイルの `CLOUD_RUN_SERVICE_URL` に設定してください。


## 4. ADK Webでのテスト

1.  ADKプロジェクトのルートディレクトリで、環境変数（特に`CLOUD_RUN_SERVICE_URL`）が`.env`ファイルに正しく設定されていることを確認します。
2.  ADKプロジェクトのルートディレクトリ（この`PPTGenerator`ディレクトリの一つ上の階層）に移動し、ADK Webを起動します。
    ```bash
    adk web
    ```
3.  ブラウザでADK Web UIを開きます。
4.  エージェントリストから `presentation_generator_agent_v3` を選択します。
5.  チャット入力欄に、作成したいプレゼンテーションのトピックを入力します。
    **例:**
    `リモートワークの利点についてのプレゼンテーションを作成して`
6.  エージェントがツールを呼び出し、処理が完了すると、GCSにアップロードされたPowerPointファイルのURLが返されます。
