#!/usr/bin/env python3
"""
智谱 GLM-4.6V 视觉模型 · 图像识别工具

用法:
    python3 glm_vision.py <图片路径或URL> ["提示词"]

示例:
    python3 glm_vision.py /tmp/screenshot.png
    python3 glm_vision.py /tmp/screenshot.png "图中有什么数据？"
    python3 glm_vision.py https://example.com/img.png "描述这张图"

API Key 读取顺序:
    1. 环境变量 ZHIPU_API_KEY
    2. 本文件同目录 .zhipu_key 文件（内容为 Key，第一行）
"""
import base64
import json
import os
import sys
import urllib.request
from pathlib import Path

API_URL = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
MODEL = os.environ.get("GLM_VISION_MODEL", "glm-4.6v")
DEFAULT_PROMPT = "请详细描述这张图片的内容。"


def get_api_key():
    """按优先级获取 API Key。"""
    key = os.environ.get("ZHIPU_API_KEY", "").strip()
    if key:
        return key
    key_file = Path(__file__).parent / ".zhipu_key"
    if key_file.exists():
        k = key_file.read_text(encoding="utf-8").strip().splitlines()
        if k and k[0].strip():
            return k[0].strip()
    raise RuntimeError(
        "未找到智谱 API Key：请设置环境变量 ZHIPU_API_KEY，"
        "或在工具同目录创建 .zhipu_key 文件"
    )


def load_image_b64(src):
    """图片路径或 URL -> base64。URL 图片直接返回原 URL。"""
    if src.startswith("http://") or src.startswith("https://"):
        return src
    p = Path(src)
    if not p.is_file():
        raise FileNotFoundError(f"图片不存在: {src}")
    size = p.stat().st_size
    if size > 4 * 1024 * 1024:
        raise ValueError(f"图片过大({size/1024/1024:.1f}MB)，请压缩到 4MB 以内")
    with open(p, "rb") as f:
        return base64.b64encode(f.read()).decode()


def recognize(image_src, prompt=DEFAULT_PROMPT, temperature=0.7, max_tokens=2048):
    """调用 GLM-4.6V 识别图片，返回模型文本。"""
    key = get_api_key()
    img = load_image_b64(image_src)
    payload = {
        "model": MODEL,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": img}},
                {"type": "text", "text": prompt},
            ],
        }],
        "temperature": temperature,
        "max_tokens": max_tokens,
        "stream": False,
    }
    req = urllib.request.Request(API_URL, data=json.dumps(payload).encode(), method="POST")
    req.add_header("Authorization", f"Bearer {key}")
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            d = json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        try:
            err = json.loads(e.read().decode()).get("error", {})
            raise RuntimeError(
                f"智谱 API 错误 HTTP {e.code}: "
                f"[{err.get('code')}] {err.get('message','')}"
            ) from e
        except json.JSONDecodeError:
            raise RuntimeError(f"智谱 API 错误 HTTP {e.code}: {e.read().decode()[:200]}") from e
    if "choices" not in d:
        raise RuntimeError(f"智谱 API 异常响应: {json.dumps(d, ensure_ascii=False)[:300]}")
    return d["choices"][0]["message"]["content"]


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    image_src = sys.argv[1]
    prompt = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_PROMPT
    try:
        result = recognize(image_src, prompt)
        print(result)
    except Exception as e:
        print(f"[错误] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
