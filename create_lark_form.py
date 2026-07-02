import os, json, urllib.request, urllib.parse
from pathlib import Path

# .envを読み込む
def load_env():
    env_path = Path(__file__).parent / ".env"
    if not env_path.exists():
        # 親フォルダも探す
        env_path = Path.home() / ".env"
    if env_path.exists():
        for line in env_path.read_text().splitlines():
            if "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"'))

load_env()

APP_ID     = os.environ.get("LARK_APP_ID")
APP_SECRET = os.environ.get("LARK_APP_SECRET")

# LarkのBaseURL（URLから取得）
APP_TOKEN  = "GENzb5xNgaRAtusrhvRjGwbopYc"
TABLE_ID   = "tblCnORjLjVJYDiv"
BASE_URL   = "https://open.larksuite.com"

if not APP_ID or not APP_SECRET:
    print("❌ LARK_APP_ID または LARK_APP_SECRET が .env に見つかりません")
    print("   .env に以下を追加してください：")
    print("   LARK_APP_ID=cli_xxxxxxxx")
    print("   LARK_APP_SECRET=xxxxxxxxxxxxxxxx")
    exit(1)

def api(method, path, body=None, token=None):
    url = BASE_URL + path
    data = json.dumps(body).encode() if body else None
    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req) as res:
        return json.loads(res.read())

# 1. テナントアクセストークン取得
print("🔑 Larkにログイン中...")
res = api("POST", "/open-apis/auth/v3/tenant_access_token/internal", {
    "app_id": APP_ID,
    "app_secret": APP_SECRET
})
if res.get("code") != 0:
    print(f"❌ 認証失敗: {res.get('msg')}")
    exit(1)

token = res["tenant_access_token"]
print("✅ 認証成功")

# 2. フォームビューを作成
print("📋 フォームを作成中...")
res = api("POST", f"/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/views", {
    "view_name": "AI経理講座 受講1ヶ月アンケート",
    "view_type": "form"
}, token)

if res.get("code") != 0:
    print(f"❌ フォーム作成失敗: {res.get('msg')}")
    print(f"   詳細: {res}")
    exit(1)

view_id = res["data"]["view"]["view_id"]
print(f"✅ フォーム作成成功！")
print(f"")
print(f"🔗 フォームURL（Lark Base内）:")
print(f"   https://djp4tsqqyucj.jp.larksuite.com/base/{APP_TOKEN}?table={TABLE_ID}&view={view_id}")
