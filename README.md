# excel-vision-rag

Microsoft Graph API を使用して SharePoint にファイルをアップロードし、共有リンクを生成する Python ツール

## 📋 機能

- SharePoint サイトへのファイルアップロード（4MB未満のファイル対応）
- 複数の共有リンクタイプの自動生成
  - 組織内共有リンク（閲覧・編集）
  - 匿名共有リンク（閲覧）
  - 直接アクセスURL
  - ダウンロードURL
- アップロードしたファイルのダウンロード機能
- 設定可能なアップロード設定
- 詳細なログ出力

## 🛠️ 必要な環境

- Python 3.11以上
- Azure App Registration（Microsoft Graph API用）
- SharePoint Online サイト

## 📦 インストール

### 2. 依存関係のインストール

```bash
# uvを使用している場合
uv sync

# pipを使用している場合
pip install -e .
```

### 3. 環境設定

`.env` ファイルを作成し、以下の環境変数を設定してください：

```env
CLIENT_ID=your_azure_app_client_id
CLIENT_SECRET=your_azure_app_client_secret
TENANT_ID=your_azure_tenant_id
SITE_INFO_URL=https://graph.microsoft.com/v1.0/sites/your_site_url
```

## 🔧 Azure App Registration の設定

1. Azure Portal にログインし、「Azure Active Directory」を選択
2. 「App registrations」→「New registration」
3. アプリケーションを登録し、以下の API 権限を追加：
   - `Sites.ReadWrite.All`
   - `Files.ReadWrite.All`
4. 「Certificates & secrets」でクライアントシークレットを生成
5. `CLIENT_ID`、`CLIENT_SECRET`、`TENANT_ID` を `.env` ファイルに記録

## 📈 使用方法

### 基本的な使用例

```python
from sharepoint_uploader import SharePointUploader

# デフォルト設定でアップローダーを初期化
uploader = SharePointUploader()

# ファイルをアップロード
result = uploader.upload_file("sample_image.png")

if result:
    print(f"アップロード成功: {result['file_info']['name']}")
    uploader.print_links(result["links"])
```

### カスタム設定での使用

```python
from sharepoint_uploader import SharePointUploader, UploadConfig

# カスタム設定
config = UploadConfig(
    default_folder="uploads",           # デフォルトフォルダ
    enable_anonymous_sharing=False,     # 匿名共有の無効化
    request_timeout=60                  # リクエストタイムアウト（秒）
)

uploader = SharePointUploader(config=config)
```

### ファイルのダウンロード

```python
# アップロードしたファイルをダウンロード
file_id = result["file_info"]["id"]
success = uploader.download_file(file_id, "downloaded_file.png", overwrite=True)

if success:
    print("ダウンロード成功")
```

## 🔧 設定オプション

### UploadConfig クラス

| パラメータ | デフォルト値 | 説明 |
|-----------|------------|------|
| `max_small_file_size` | 4MB | 小さいファイルの最大サイズ |
| `default_folder` | "" | デフォルトのアップロードフォルダ |
| `enable_anonymous_sharing` | True | 匿名共有リンクの生成を有効化 |
| `request_timeout` | 30 | API リクエストのタイムアウト（秒） |

### SharePointCredentials クラス

環境変数から自動的に認証情報を取得しますが、直接指定することも可能です：

```python
from sharepoint_uploader import SharePointCredentials

credentials = SharePointCredentials(
    client_id="your_client_id",
    client_secret="your_client_secret",
    tenant_id="your_tenant_id",
    site_info_url="your_site_url"
)

uploader = SharePointUploader(credentials=credentials)
```

## 📊 リンクタイプ

アップロード成功時に以下のリンクが生成されます：

- **直接アクセスURL**: SharePoint上でファイルを直接開く
- **ダウンロードURL**: ファイルを直接ダウンロード
- **組織内共有（閲覧）**: 組織内のメンバーが閲覧可能
- **組織内共有（編集）**: 組織内のメンバーが編集可能
- **匿名共有（閲覧）**: 認証なしでアクセス可能（設定で有効化されている場合）

## 🚀 実行

### コマンドラインでの実行

```bash
python sharepoint_uploader.py
```

プロジェクトディレクトリに `sample_image.png` ファイルを配置してから実行してください。

### スクリプトとしての使用

```python
from sharepoint_uploader import SharePointUploader

def upload_my_file():
    uploader = SharePointUploader()
    
    # ファイルをアップロード
    result = uploader.upload_file(
        local_file_path="my_file.pdf",
        remote_file_name="uploaded_file.pdf",
        folder_path="documents"
    )
    
    if result:
        return result["links"]
    return None
```

## 📝 ログ出力

アプリケーションは詳細なログを出力します：

- 認証状況
- サイト情報取得
- アップロード進行状況
- エラー詳細

## ⚠️ 制限事項

- 現在、4MB未満のファイルのみサポート
- 大きなファイルのアップロードには未対応
- SharePoint Online のみサポート

## 🐛 トラブルシューティング

### よくある問題

1. **認証エラー**
   - 環境変数が正しく設定されているか確認
   - Azure App Registration の権限設定を確認

2. **ファイルが見つからない**
   - ファイルパスが正確か確認
   - ファイルの存在を確認

3. **アップロード失敗**
   - ファイルサイズが4MB未満か確認
   - ネットワーク接続を確認
