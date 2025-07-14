"""
SharePoint File Uploader

Microsoft Graph APIを使用してSharePointにファイルをアップロードし、
共有リンクを生成するツール

必要な環境変数:
- CLIENT_ID: Azure App RegistrationのクライアントID
- CLIENT_SECRET: Azure App Registrationのクライアントシークレット
- TENANT_ID: Azure ADのテナントID
- SITE_INFO_URL: SharePointサイトのGraph API URL

使用例:
    uploader = SharePointUploader()
    result = uploader.upload_file("sample_image.png")
    if result:
        uploader.print_links(result["links"])
"""

import os
import logging
import mimetypes
from typing import Dict, Optional, Any
from dataclasses import dataclass
import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv


@dataclass
class UploadConfig:
    """アップロード設定クラス"""

    max_small_file_size: int = 4 * 1024 * 1024  # 4MB
    default_folder: str = ""
    enable_anonymous_sharing: bool = True
    request_timeout: int = 30


@dataclass
class SharePointCredentials:
    """SharePoint認証情報クラス"""

    client_id: str
    client_secret: str
    tenant_id: str
    site_info_url: str

    @classmethod
    def from_env(cls) -> "SharePointCredentials":
        """環境変数から認証情報を取得"""
        load_dotenv()
        return cls(
            client_id=os.environ.get("CLIENT_ID") or "",
            client_secret=os.environ.get("CLIENT_SECRET") or "",
            tenant_id=os.environ.get("TENANT_ID") or "",
            site_info_url=os.environ.get("SITE_INFO_URL") or "",
        )

    def validate(self) -> None:
        """認証情報の妥当性をチェック"""
        missing = []
        if not self.client_id:
            missing.append("CLIENT_ID")
        if not self.client_secret:
            missing.append("CLIENT_SECRET")
        if not self.tenant_id:
            missing.append("TENANT_ID")
        if not self.site_info_url:
            missing.append("SITE_INFO_URL")

        if missing:
            raise ValueError(
                f"以下の環境変数が設定されていません: {', '.join(missing)}"
            )


class SharePointUploader:
    """SharePointファイルアップローダー"""

    def __init__(
        self,
        credentials: Optional[SharePointCredentials] = None,
        config: Optional[UploadConfig] = None,
    ):
        """
        初期化

        Args:
            credentials: SharePoint認証情報（Noneの場合は環境変数から取得）
            config: アップロード設定（Noneの場合はデフォルト設定を使用）
        """
        self.credentials = credentials or SharePointCredentials.from_env()
        self.config = config or UploadConfig()

        # ログ設定
        self._setup_logging()

        # 認証情報検証
        self.credentials.validate()

        # Graph API設定
        self.authority_url = (
            f"https://login.microsoftonline.com/{self.credentials.tenant_id}"
        )
        self.scope = ["https://graph.microsoft.com/.default"]

        # 初期化時に認証とサイト情報を取得
        self._access_token = None
        self._site_id = None
        self._drive_id = None

        self._authenticate()
        self._get_site_info()

    def _setup_logging(self) -> None:
        """ログ設定"""
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        )
        self.logger = logging.getLogger(__name__)

    def _authenticate(self) -> None:
        """Microsoft Graph APIの認証"""
        try:
            app = ConfidentialClientApplication(
                self.credentials.client_id,
                authority=self.authority_url,
                client_credential=self.credentials.client_secret,
            )

            token_result = app.acquire_token_for_client(scopes=self.scope)
            if token_result is None:
                raise Exception(f"認証に失敗しました")
            self._access_token = token_result.get("access_token")
            self.logger.info("Microsoft Graph API認証成功")

        except Exception as e:
            self.logger.error(f"認証エラー: {e}")
            raise

    def _get_site_info(self) -> None:
        """サイト情報とドライブIDを取得"""
        try:
            headers = {
                "Authorization": f"Bearer {self._access_token}",
                "Content-Type": "application/json",
            }

            # サイトID取得
            response = requests.get(
                self.credentials.site_info_url,
                headers=headers,
                timeout=self.config.request_timeout,
            )
            response.raise_for_status()

            self._site_id = response.json().get("id").split(",")[1]
            self.logger.info(f"サイトID取得成功: {self._site_id}")

            # ドライブ一覧取得
            drive_url = f"https://graph.microsoft.com/v1.0/sites/{self._site_id}/drives"
            response = requests.get(
                drive_url, headers=headers, timeout=self.config.request_timeout
            )
            response.raise_for_status()

            drives = response.json().get("value", [])
            if not drives:
                raise Exception("利用可能なドライブが見つかりません")

            self._drive_id = drives[0]["id"]
            self.logger.info(f"ドライブID取得成功: {self._drive_id}")

        except requests.RequestException as e:
            self.logger.error(f"サイト情報取得エラー: {e}")
            raise
        except Exception as e:
            self.logger.error(f"サイト情報処理エラー: {e}")
            raise

    @property
    def headers(self) -> Dict[str, str]:
        """API呼び出し用ヘッダー"""
        return {
            "Authorization": f"Bearer {self._access_token}",
            "Content-Type": "application/json",
        }

    def upload_file(
        self,
        local_file_path: str,
        remote_file_name: Optional[str] = None,
        folder_path: Optional[str] = None,
    ) -> Optional[Dict[str, Any]]:
        """
        ファイルをSharePointにアップロード

        Args:
            local_file_path: ローカルファイルのパス
            remote_file_name: アップロード先のファイル名（Noneの場合はローカルファイル名を使用）
            folder_path: フォルダパス（Noneの場合は設定のデフォルトフォルダを使用）

        Returns:
            アップロード結果とリンク情報、失敗時はNone
        """
        try:
            # ファイル存在確認
            if not os.path.exists(local_file_path):
                self.logger.error(f"ファイルが見つかりません: {local_file_path}")
                return None

            # ファイル名決定
            if remote_file_name is None:
                remote_file_name = os.path.basename(local_file_path)

            # フォルダパス決定
            if folder_path is None:
                folder_path = self.config.default_folder

            # ファイル情報取得
            file_size = os.path.getsize(local_file_path)
            mime_type, _ = mimetypes.guess_type(local_file_path)
            if mime_type is None:
                mime_type = "application/octet-stream"

            self.logger.info(f"アップロード開始: {local_file_path}")
            self.logger.info(f"ファイルサイズ: {file_size:,} bytes")
            self.logger.info(f"MIMEタイプ: {mime_type}")

            # ファイルサイズに応じてアップロード方法を選択
            if file_size < self.config.max_small_file_size:
                return self._upload_small_file(
                    local_file_path, remote_file_name, folder_path, mime_type
                )
            else:
                self.logger.error(f"4MB以上のファイルは未対応です: {file_size:,} bytes")
                return None

        except Exception as e:
            self.logger.error(f"ファイルアップロードエラー: {e}")
            return None

    def _upload_small_file(
        self,
        local_file_path: str,
        remote_file_name: str,
        folder_path: str,
        mime_type: str,
    ) -> Optional[Dict[str, Any]]:
        """4MB未満のファイルを直接アップロード"""
        try:
            # アップロードパス構築
            if folder_path:
                full_path = f"{folder_path}/{remote_file_name}"
            else:
                full_path = remote_file_name

            # ファイル読み込み
            with open(local_file_path, "rb") as f:
                file_content = f.read()

            # アップロード実行
            upload_url = f"https://graph.microsoft.com/v1.0/sites/{self._site_id}/drives/{self._drive_id}/root:/{full_path}:/content"
            upload_headers = {
                "Authorization": f"Bearer {self._access_token}",
                "Content-Type": mime_type,
            }

            response = requests.put(
                upload_url,
                headers=upload_headers,
                data=file_content,
                timeout=self.config.request_timeout,
            )
            response.raise_for_status()

            file_info = response.json()
            file_id = file_info["id"]

            self.logger.info(f"ファイルアップロード成功: {remote_file_name}")

            # リンク取得
            links = self._get_file_links(file_id)

            return {"file_info": file_info, "links": links}

        except requests.RequestException as e:
            self.logger.error(f"アップロードAPIエラー: {e}")
            return None
        except Exception as e:
            self.logger.error(f"アップロード処理エラー: {e}")
            return None

    def _get_file_links(self, file_id: str) -> Dict[str, str]:
        """ファイルの各種リンクを取得"""
        links = {}

        try:
            # ファイル情報取得
            file_info_url = f"https://graph.microsoft.com/v1.0/sites/{self._site_id}/drives/{self._drive_id}/items/{file_id}"
            response = requests.get(
                file_info_url, headers=self.headers, timeout=self.config.request_timeout
            )
            response.raise_for_status()

            file_data = response.json()
            links["direct_url"] = file_data.get("webUrl")
            links["download_url"] = file_data.get("@microsoft.graph.downloadUrl")

            # 共有リンク作成
            sharing_links = self._create_sharing_links(file_id)
            links.update(sharing_links)

            self.logger.info("リンク取得成功")

        except Exception as e:
            self.logger.error(f"リンク取得エラー: {e}")

        return links

    def _create_sharing_links(self, file_id: str) -> Dict[str, str]:
        """共有リンクを作成"""
        sharing_links = {}

        # 組織内共有リンク（閲覧）
        view_link = self._create_sharing_link(file_id, "view", "organization")
        if view_link:
            sharing_links["organization_view_link"] = view_link

        # 組織内共有リンク（編集）
        edit_link = self._create_sharing_link(file_id, "edit", "organization")
        if edit_link:
            sharing_links["organization_edit_link"] = edit_link

        # 匿名共有リンク（設定で有効な場合のみ）
        if self.config.enable_anonymous_sharing:
            anonymous_link = self._create_sharing_link(file_id, "view", "anonymous")
            if anonymous_link:
                sharing_links["anonymous_view_link"] = anonymous_link

        return sharing_links

    def _create_sharing_link(
        self, file_id: str, permission_type: str, scope: str
    ) -> Optional[str]:
        """特定の権限で共有リンクを作成"""
        try:
            create_link_url = f"https://graph.microsoft.com/v1.0/sites/{self._site_id}/drives/{self._drive_id}/items/{file_id}/createLink"

            body = {"type": permission_type, "scope": scope}

            response = requests.post(
                create_link_url,
                headers=self.headers,
                json=body,
                timeout=self.config.request_timeout,
            )
            response.raise_for_status()

            link_data = response.json()
            return link_data.get("link", {}).get("webUrl")

        except requests.RequestException as e:
            self.logger.warning(f"共有リンク作成失敗 ({permission_type}, {scope}): {e}")
            return None

    def print_links(self, links: Dict[str, str]) -> None:
        """取得したリンクを整理して表示"""
        print("\n=== ファイルリンク情報 ===")

        link_types = [
            ("direct_url", "直接アクセスURL"),
            ("download_url", "ダウンロードURL"),
            ("organization_view_link", "組織内共有（閲覧）"),
            ("organization_edit_link", "組織内共有（編集）"),
            ("anonymous_view_link", "匿名共有（閲覧）"),
        ]

        for key, label in link_types:
            if key in links and links[key]:
                print(f"{label}: {links[key]}")

        print("========================\n")

    def get_site_info(self) -> Dict[str, str]:
        """サイト情報を取得"""
        if not self._site_id or not self._drive_id:
            raise ValueError(
                "サイト情報が取得されていません。_get_site_info()を実行してください。"
            )
        return {"site_id": self._site_id, "drive_id": self._drive_id}

    def download_file(
        self,
        file_id: str,
        local_file_path: str,
        overwrite: bool = False,
    ) -> bool:
        """
        SharePoint 上のファイルをダウンロードし、任意の名前で保存する

        Args:
            file_id        : ダウンロード対象ファイルの ID
            local_file_path: 保存先（別名を含む）フルパス
            overwrite      : True なら既存ファイルを上書き

        Returns:
            成功時 True / 失敗時 False
        """
        try:
            if os.path.exists(local_file_path) and not overwrite:
                self.logger.error(f"既に存在します: {local_file_path}")
                return False

            # /content エンドポイントは 302 で実体 URL にリダイレクトする
            content_url = (
                f"https://graph.microsoft.com/v1.0/"
                f"sites/{self._site_id}/drives/{self._drive_id}/items/{file_id}/content"
            )
            with requests.get(
                content_url,
                headers={"Authorization": f"Bearer {self._access_token}"},
                stream=True,
                timeout=self.config.request_timeout,
            ) as resp:
                resp.raise_for_status()
                with open(local_file_path, "wb") as fp:
                    for chunk in resp.iter_content(chunk_size=8192):
                        fp.write(chunk)

            self.logger.info(f"ダウンロード成功: {local_file_path}")
            return True

        except requests.RequestException as e:
            self.logger.error(f"ダウンロード API エラー: {e}")
        except Exception as e:
            self.logger.error(f"ダウンロード処理エラー: {e}")
        return False

def main():
    """メイン実行関数"""
    try:
        # アップローダー初期化
        config = UploadConfig(
            default_folder="uploads",  # デフォルトフォルダを設定
            enable_anonymous_sharing=False,  # 匿名共有を無効化
        )

        uploader = SharePointUploader(config=config)

        # サイト情報表示
        site_info = uploader.get_site_info()
        print(f"接続先 - Site ID: {site_info['site_id']}")
        print(f"接続先 - Drive ID: {site_info['drive_id']}")

        # ファイルアップロード
        local_file = "sample_image.png"

        if not os.path.exists(local_file):
            print(f"ファイルが見つかりません: {local_file}")
            print("同じディレクトリにsample_image.pngを配置してください")
            return

        result = uploader.upload_file(local_file)

        if result:
            print(f"アップロード成功: {result['file_info']['name']}")
            uploader.print_links(result["links"])

            # 画像の場合は追加情報表示
            file_info = result["file_info"]
            if "image" in file_info:
                image_info = file_info["image"]
                width = image_info.get("width", "N/A")
                height = image_info.get("height", "N/A")
                print(f"画像サイズ: {width} x {height}")
        else:
            print("アップロードに失敗しました")

        # …アップロード直後
        if result:
            print(f"アップロード成功: {result['file_info']['name']}")
            uploader.print_links(result["links"])

            # 追加: アップロードしたファイルを別名でダウンロード
            file_id = result["file_info"]["id"]
            save_as = "renamed_sample.png"
            if uploader.download_file(file_id, save_as, overwrite=True):
                print(f"ダウンロード完了: {save_as}")
            else:
                print("ダウンロードに失敗しました")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        logging.error(f"メイン処理エラー: {e}")


if __name__ == "__main__":
    main()
