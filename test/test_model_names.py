"""测试 Atlas 网关接受哪种模型名格式，帮助排查 Claude Code 400 not found 问题。"""
import requests

BASE_URL = "https://api.atlascloud.ai/v1"
API_KEY = "apikey-ba28bd46b2bf470aa8a6718db4b6839a"

HEADERS_BEARER = {
    "Content-Type": "application/json",
    "Authorization": f"Bearer {API_KEY}",
}

HEADERS_ANTHROPIC = {
    "Content-Type": "application/json",
    "x-api-key": API_KEY,
    "anthropic-version": "2023-06-01",
}

MODEL_NAMES = [
    "anthropic/claude-opus-4.6",
    "anthropic/claude-opus-4-6",
    "claude-opus-4-6",
    "claude-opus-4.6",
    "anthropic/claude-opus-4-20250514",
    "anthropic/claude-opus-4-20250514-developer",
]

MESSAGES = [{"role": "user", "content": "say hi"}]

def test_model(model_name, headers, header_label):
    url = f"{BASE_URL}/messages"
    body = {"model": model_name, "max_tokens": 32, "messages": MESSAGES}
    try:
        resp = requests.post(url, json=body, headers=headers, timeout=30)
        status = resp.status_code
        text = resp.text[:200]
    except Exception as e:
        status = "ERR"
        text = str(e)[:200]
    ok = "OK" if status == 200 else "FAIL"
    print(f"  [{ok}] {status} | {header_label:12s} | model={model_name}")
    if status != 200:
        print(f"         response: {text}")
    return status == 200

# 同时测试不同 base URL 前缀
BASE_URLS = [
    "https://api.atlascloud.ai/v1",
    "https://api.atlascloud.ai/api/v1",
]

if __name__ == "__main__":
    for base in BASE_URLS:
        print(f"\n===== BASE: {base} =====")
        for model in MODEL_NAMES[:2]:
            url = f"{base}/messages"
            body = {"model": model, "max_tokens": 32, "messages": MESSAGES}
            try:
                resp = requests.post(url, json=body, headers=HEADERS_BEARER, timeout=30)
                status = resp.status_code
                text = resp.text[:200]
            except Exception as e:
                status = "ERR"
                text = str(e)[:200]
            ok = "OK" if status == 200 else "FAIL"
            print(f"  [{ok}] {status} | Bearer       | model={model}")
            if status != 200:
                print(f"         response: {text}")

    print(f"\n===== 模型名格式测试 (Bearer, base={BASE_URL}) =====")
    for model in MODEL_NAMES:
        test_model(model, HEADERS_BEARER, "Bearer")

    print(f"\n===== 模型名格式测试 (x-api-key, base={BASE_URL}) =====")
    for model in MODEL_NAMES[:3]:
        test_model(model, HEADERS_ANTHROPIC, "x-api-key")
