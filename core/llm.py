# OpenAI 兼容 LLM 客户端 —— 不引框架，直接 requests 打 HTTP。
# 配置走环境变量（也可写在项目根目录 .env 文件里，.env 已被 gitignore）：
#   LLM_BASE_URL / LLM_API_KEY / LLM_MODEL / LLM_TIMEOUT
#   LLM_ALLOW_PUBLIC_IMAGE_UPLOAD（默认 false）
#   中转站（VibeToken）: https://vibetoken.cn/v1
# 约定：temperature=0；超时默认 120s；指数退避重试 <=2 次。
#
# 已知的中转站适配（实测踩出来的坑）：
#   1. 某些模型不支持 response_format JSON 模式（400）→ 自动去掉重试
#   2. VibeToken 的上游无法拉取 data:base64 内联图片（500 "Error fetching file"）
#      → 视觉调用自动把图片上传到临时图床，改用 https URL 重试（见 _upload_image）
#   3. 推理模型（如 kimi-k3）可能只回 reasoning_content、content 为空 → 视为失败重试

import base64
import json
import os
import time

import requests


_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_DEFAULT_ENV_PATH = os.path.join(_PROJECT_ROOT, ".env")


def load_dotenv(path=None, override=False):
    """极简 .env 加载：KEY=VALUE 逐行读。
    override=False（默认）：不覆盖已存在的环境变量；
    override=True：以 .env 为准（Streamlit 这类长驻进程每次重跑时刷新配置用）。
    """
    path = path or _DEFAULT_ENV_PATH
    if not os.path.exists(path):
        return
    setter = os.environ.__setitem__ if override else os.environ.setdefault
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, _, v = line.partition("=")
            setter(k.strip(), v.strip().strip('"').strip("'"))


load_dotenv()


class LLMError(Exception):
    """LLM 调用失败（网络/HTTP/解析），重试耗尽后抛出。"""


class LLMHTTPError(LLMError):
    """OpenAI 兼容接口返回了非 2xx 状态。"""

    def __init__(self, status_code, message, request_id=None):
        self.status_code = int(status_code)
        self.request_id = request_id
        suffix = f" (request_id={request_id})" if request_id else ""
        super().__init__(f"HTTP {self.status_code}{suffix}: {message}")

    @property
    def retryable(self):
        return self.status_code in {408, 409, 425, 429} or self.status_code >= 500


def _http_error(resp):
    """构造带响应体的错误信息（中转站的错误细节都在 body 里，不看 body 没法排查）。"""
    request_id = (
        resp.headers.get("x-request-id")
        or resp.headers.get("request-id")
        or resp.headers.get("cf-ray")
    )
    return LLMHTTPError(resp.status_code, resp.text[:600], request_id=request_id)


def _should_retry(error):
    return not isinstance(error, LLMHTTPError) or error.retryable


def _inline_image_was_rejected(error):
    message = str(error).lower()
    markers = (
        "error fetching file",
        "invalid_image_url",
        "failed to fetch image",
        "unable to download image",
        "data:image",
        "data url",
    )
    return any(marker in message for marker in markers)


def _message_text(content):
    """兼容返回字符串或 OpenAI 新式 content parts 的中转站。"""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for item in content:
            if not isinstance(item, dict):
                continue
            value = item.get("text")
            if isinstance(value, str):
                parts.append(value)
        return "\n".join(parts)
    return ""


def _upload_image(path, timeout=60):
    """把本地图片传到临时公开图床，返回可直接拉取的 https URL。
    供不支持 data: 内联图片的中转站用。
    注意：图片会变成任何人可访问的公网 URL，含敏感内容的文档不要用这条路。
    依次尝试 litterbox(catbox 临时版) → catbox.moe → 0x0.st。
    """
    errors = []
    # litterbox.catbox.moe（1 小时有效期，给视觉调用用正好；返回直链）
    try:
        with open(path, "rb") as f:
            r = requests.post("https://litterbox.catbox.moe/resources/internals/api.php",
                              data={"reqtype": "fileupload", "time": "1h"},
                              files={"fileToUpload": f}, timeout=timeout)
        if r.ok and r.text.strip().startswith("http"):
            return r.text.strip()
        errors.append(f"litterbox: HTTP {r.status_code} {r.text[:100]}")
    except Exception as e:  # noqa: BLE001
        errors.append(f"litterbox: {e}")
    # catbox.moe
    try:
        with open(path, "rb") as f:
            r = requests.post("https://catbox.moe/user/api.php",
                              data={"reqtype": "fileupload"},
                              files={"fileToUpload": f}, timeout=timeout)
        if r.ok and r.text.strip().startswith("http"):
            return r.text.strip()
        errors.append(f"catbox: HTTP {r.status_code} {r.text[:100]}")
    except Exception as e:  # noqa: BLE001
        errors.append(f"catbox: {e}")
    # 0x0.st
    try:
        with open(path, "rb") as f:
            r = requests.post("https://0x0.st",
                              files={"file": f},
                              headers={"User-Agent": "format-agent/1.0"}, timeout=timeout)
        if r.ok and r.text.strip().startswith("http"):
            return r.text.strip()
        errors.append(f"0x0.st: HTTP {r.status_code} {r.text[:100]}")
    except Exception as e:  # noqa: BLE001
        errors.append(f"0x0.st: {e}")
    raise LLMError("图片上传图床失败（" + "；".join(errors) + "）")


class LLMClient:
    def __init__(self, base_url=None, api_key=None, model=None,
                 timeout=None, max_retries=2, on_event=None):
        self.base_url = (base_url or os.environ.get("LLM_BASE_URL") or "").rstrip("/")
        self.api_key = api_key or os.environ.get("LLM_API_KEY") or ""
        self.model = model or os.environ.get("LLM_MODEL") or ""
        self.timeout = timeout if timeout is not None else int(
            os.environ.get("LLM_TIMEOUT", "120"))
        # 某些端点的模型（如 kimi coding）只允许 temperature=1，用环境变量适配。
        raw_temperature = os.environ.get("LLM_TEMPERATURE", "0").strip()
        try:
            self.temperature = float(raw_temperature)
        except ValueError:
            raise LLMError(f"LLM_TEMPERATURE 必须是数字，当前值为 {raw_temperature!r}")
        max_tokens = os.environ.get("LLM_MAX_TOKENS", "4096").strip()
        try:
            self.max_tokens = max(256, int(max_tokens))
        except ValueError:
            self.max_tokens = 4096
        self.allow_public_image_upload = (
            os.environ.get("LLM_ALLOW_PUBLIC_IMAGE_UPLOAD", "false").strip().lower()
            in {"1", "true", "yes", "on"}
        )
        self.max_retries = max_retries
        self.on_event = on_event or (lambda msg: None)
        self.last_response_meta = {}
        if not self.base_url or not self.api_key or not self.model:
            raise LLMError(
                "缺少 LLM 配置：请设置环境变量 LLM_BASE_URL / LLM_API_KEY / LLM_MODEL"
            )

    def chat_json(self, prompt, model=None):
        """发单轮 user prompt，要求模型输出 JSON，返回解析后的对象。
        失败（HTTP 错误/超时/JSON 解析失败）指数退避重试，耗尽后抛 LLMError。
        """
        last_err = None
        for attempt in range(self.max_retries + 1):
            try:
                text = self._chat(prompt, model=model, json_mode=True)
                return _parse_json(text)
            except Exception as e:  # noqa: BLE001 —— 统一收口为重试
                last_err = e
                if not _should_retry(e):
                    raise LLMError(f"LLM 调用失败（不可重试）: {e}") from e
                if attempt < self.max_retries:
                    self.on_event(f"LLM 调用失败（{e}），{2 ** attempt}s 后重试（第 {attempt + 1} 次）")
                    time.sleep(2 ** attempt)  # 1s, 2s
        raise LLMError(f"LLM 调用失败（重试 {self.max_retries} 次后放弃）: {last_err}")

    def chat_vision_json(self, prompt, image_paths, model=None):
        """使用同一个 LLM_MODEL 做带图调用：prompt + png 列表 → JSON 对象。
        默认先用 data:base64 内联图片。只有显式启用 LLM_ALLOW_PUBLIC_IMAGE_UPLOAD
        时，才允许在中转站不接受内联图片的情况下上传临时公开图床。
        """
        image_paths = list(image_paths or [])
        if not image_paths:
            raise LLMError("视觉调用没有收到任何图片，已拒绝降级为纯文本请求")
        for path in image_paths:
            if not os.path.isfile(path) or os.path.getsize(path) <= 0:
                raise LLMError(f"视觉调用的图片不存在或为空: {path}")

        use_urls = False  # 中转站明确拒绝 data URL 后才改用图床 URL
        last_err = None
        for attempt in range(self.max_retries + 1):
            try:
                text = self._chat(prompt, model=model or self.model,
                                  json_mode=True, image_paths=image_paths,
                                  image_as_url=use_urls)
                return _parse_json(text)
            except Exception as e:  # noqa: BLE001
                last_err = e
                switched_to_urls = False
                if not use_urls and _inline_image_was_rejected(e):
                    if not self.allow_public_image_upload:
                        raise LLMError(
                            "接口不接受 data:base64 内联图片；为保护文档隐私，程序没有上传"
                            "公开图床。请换用支持内联图片的多模态接口，或在确认文档不敏感后"
                            "设置 LLM_ALLOW_PUBLIC_IMAGE_UPLOAD=true。"
                        ) from e
                    use_urls = True
                    switched_to_urls = True
                    self.on_event("中转站不支持内联图片，改为上传临时图床后用 URL 调用"
                                  "（注意：渲染图会变成公网可访问的临时链接）")
                if not switched_to_urls and not _should_retry(e):
                    raise LLMError(f"视觉模型调用失败（不可重试）: {e}") from e
                if attempt < self.max_retries:
                    self.on_event(f"视觉模型调用失败（{e}），{2 ** attempt}s 后重试（第 {attempt + 1} 次）")
                    time.sleep(2 ** attempt)
        raise LLMError(f"视觉模型调用失败（重试 {self.max_retries} 次后放弃）: {last_err}")

    def _chat(self, prompt, model=None, json_mode=True, image_paths=None, image_as_url=False):
        url = f"{self.base_url}/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        if image_paths:
            content = [{"type": "text", "text": prompt}]
            for path in image_paths:
                if image_as_url:
                    img_url = _upload_image(path, timeout=self.timeout)
                else:
                    with open(path, "rb") as f:
                        b64 = base64.b64encode(f.read()).decode("ascii")
                    img_url = f"data:image/png;base64,{b64}"
                content.append({"type": "image_url", "image_url": {"url": img_url}})
            messages = [{"role": "user", "content": content}]
        else:
            messages = [{"role": "user", "content": prompt}]
        body = {
            "model": model or self.model,
            "messages": messages,
            "temperature": self.temperature,
            "max_tokens": self.max_tokens,
        }
        if json_mode:
            body["response_format"] = {"type": "json_object"}
        resp = requests.post(url, headers=headers, json=body, timeout=self.timeout)
        # 仅对“字段不兼容”的 400 做一次精确降级；其他 4xx 不盲目重试。
        if resp.status_code == 400:
            details = resp.text.lower()
            changed = False
            if json_mode and "response_format" in details and "response_format" in body:
                body.pop("response_format", None)
                changed = True
            if "max_tokens" in details and "max_tokens" in body:
                body.pop("max_tokens", None)
                changed = True
            # 模型强制 temperature=1 时，按错误提示自适应一次。
            if "temperature" in details and "only 1 is allowed" in details:
                body["temperature"] = 1
                changed = True
            if changed:
                resp = requests.post(url, headers=headers, json=body, timeout=self.timeout)
        if not resp.ok:
            raise _http_error(resp)
        payload = resp.json()
        choice = payload["choices"][0]
        message = choice["message"]
        text = _message_text(message.get("content"))
        self.last_response_meta = {
            "model": payload.get("model") or body["model"],
            "finish_reason": choice.get("finish_reason"),
            "usage": payload.get("usage"),
            "request_id": (
                resp.headers.get("x-request-id")
                or resp.headers.get("request-id")
                or resp.headers.get("cf-ray")
            ),
        }
        if not text:
            # 推理模型可能只回 reasoning_content；拿不到正文视为失败，让重试兜底
            raise LLMError("模型返回了空 content（可能 token 被推理过程耗尽）")
        return text


def _parse_json(text):
    """从模型输出里抠出 JSON。容忍 ```json 围栏和前后杂文本。"""
    text = text.strip()
    if text.startswith("```"):
        # 去掉首尾围栏行
        lines = text.splitlines()
        lines = [l for l in lines if not l.strip().startswith("```")]
        text = "\n".join(lines).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        # 找第一个 { 或 [ 到最后一个 } 或 ]
        start = min([i for i in (text.find("{"), text.find("[")) if i >= 0], default=-1)
        end = max(text.rfind("}"), text.rfind("]"))
        if start >= 0 and end > start:
            try:
                return json.loads(text[start:end + 1])
            except json.JSONDecodeError:
                pass
        excerpt = text[:500].replace("\n", " ")
        raise ValueError(f"模型返回内容不是合法 JSON: {excerpt!r}")
