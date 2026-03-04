import requests

BASE_URL = "https://api.atlascloud.ai/v1"
API_KEY = "apikey-ba28bd46b2bf470aa8a6718db4b6839a"

# Atlas 使用 Bearer；若直连 Anthropic 需改用 x-api-key
DEFAULT_HEADERS = {
    "Content-Type": "application/json",
    "Authorization": f"Bearer {API_KEY}",
}

# Anthropic 格式请求头（部分网关需要）
ANTHROPIC_HEADERS = {
    "Content-Type": "application/json",
    "x-api-key": API_KEY,
    "anthropic-version": "2023-06-01",
}


def get_models():
    """获取模型列表，返回 API 原始 JSON；若请求失败返回 None。"""
    url = f"{BASE_URL}/models"
    response = requests.get(url, headers=DEFAULT_HEADERS)
    print("Status:", response.status_code)
    print(response.text)
    if not response.ok:
        return None
    data = response.json()
    if "data" in data:
        for m in data["data"]:
            print("-", m.get("id", m.get("model", m)))
    return data


def call_opus_4_6(
    messages,
    max_tokens=4096,
    system=None,
    stream=False,
):
    """
    使用 Anthropic Messages API 格式调用 Claude Opus 4.6。

    :param messages: 消息列表，如 [{"role": "user", "content": "你好"}]
    :param max_tokens: 最大生成 token 数
    :param system: 可选，系统提示词
    :param stream: 是否流式返回
    :return: 非流式时返回回复文本；流式时返回原始 response
    """
    url = f"{BASE_URL}/messages"
    body = {
        "model": "anthropic/claude-opus-4.6",
        "max_tokens": max_tokens,
        "messages": messages,
    }
    if system is not None:
        body["system"] = system
    # 使用 Bearer（Atlas 常用）；若网关要求 Anthropic 头可改为 ANTHROPIC_HEADERS
    response = requests.post(url, json=body, headers=DEFAULT_HEADERS)
    if not response.ok:
        print("Status:", response.status_code)
        print("Response:", response.text)
        return None
    if stream:
        return response
    data = response.json()
    # Anthropic 返回 content: [ {"type": "text", "text": "..."}, ... ]
    if "content" in data and isinstance(data["content"], list):
        parts = [
            b["text"]
            for b in data["content"]
            if isinstance(b, dict) and b.get("type") == "text" and "text" in b
        ]
        return "".join(parts) if parts else None
    return None


if __name__ == "__main__":
    # 示例：获取模型列表
    print("=== 模型列表 ===")
    # get_models()

    # 示例：Anthropic 格式调用 Opus 4.6（POST /messages，body 含 model/max_tokens/messages）
    print("\n=== 调用 Claude Opus 4.6 (Anthropic 格式) ===")
    reply = call_opus_4_6([{"role": "user", "content": "你好"}])
    if reply:
        print("回复:", reply)
