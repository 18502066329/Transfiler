import json
import re
import httpx
import asyncio
from typing import List, Dict, Any, Callable, Optional
from backend.core.glossary_matcher import GlossaryMatcher

def clean_api_key(key: str) -> str:
    """清理 API Key 中的空格、不可见字符、引号及前缀"""
    if not key:
        return ""
    k = key.strip().strip("\ufeff\u200b\u200c\u200d\u3000\r\n\t ")
    # 去除外层引号
    if (k.startswith('"') and k.endswith('"')) or (k.startswith("'") and k.endswith("'")):
        k = k[1:-1].strip()
    # 去除多余的 "Bearer " 前缀
    if k.lower().startswith("bearer "):
        k = k[7:].strip()
    return k

def clean_base_url(url: str) -> str:
    """标准化 Base URL"""
    if not url:
        return "https://api.deepseek.com/v1"
    u = url.strip().strip("\ufeff\u200b\u3000\r\n\t ")
    # 自动补齐协议头
    if not u.startswith("http://") and not u.startswith("https://"):
        u = "https://" + u
    u = u.rstrip("/")
    # 如果用户填入了完整的 /chat/completions 端点，自动去除
    if u.endswith("/chat/completions"):
        u = u[:-len("/chat/completions")].rstrip("/")
    return u

def parse_json_from_response(content: str) -> Dict[str, Any]:
    """从大模型返回的文本中稳健提取 JSON 字典"""
    if not content:
        return {}
    content = content.strip()
    # 1. 直接尝试 json.loads
    try:
        return json.loads(content)
    except Exception:
        pass

    # 2. 尝试从 ```json ... ``` 代码块中提取
    match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", content, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except Exception:
            pass

    # 3. 尝试贪婪匹配最外层的 { ... }
    start = content.find("{")
    end = content.rfind("}")
    if start != -1 and end != -1 and end > start:
        try:
            return json.loads(content[start:end + 1])
        except Exception:
            pass

    return {}

class LLMTranslator:
    def __init__(self, api_key: str, base_url: str = "https://api.deepseek.com/v1", model_name: str = "deepseek-chat"):
        self.api_key = clean_api_key(api_key)
        self.base_url = clean_base_url(base_url)
        self.model_name = (model_name or "deepseek-chat").strip()

    def _get_headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }

    async def _post_request(self, payload: Dict[str, Any], timeout: float = 30.0) -> httpx.Response:
        """带 SSL 自动降级和多环境兼容的 POST 请求"""
        url = f"{self.base_url}/chat/completions"
        headers = self._get_headers()

        try:
            async with httpx.AsyncClient(timeout=timeout, follow_redirects=True, trust_env=True) as client:
                return await client.post(url, headers=headers, json=payload)
        except (httpx.ConnectError, httpx.SecurityError, Exception) as e:
            # 在某些工业内网/防火墙/代理环境下，重试关闭 SSL 校验
            async with httpx.AsyncClient(timeout=timeout, follow_redirects=True, verify=False, trust_env=True) as client:
                return await client.post(url, headers=headers, json=payload)

    async def test_connection(self) -> Dict[str, Any]:
        """测试 API Key 与服务连通性（带详细诊断）"""
        if not self.api_key:
            return {"success": False, "message": "API Key 为空，请输入有效的 API 密钥"}

        # 使用最广泛兼容的测试 payload (仅包含 user 消息，不加 system，避免某些推理模型拒绝)
        payload = {
            "model": self.model_name,
            "messages": [
                {"role": "user", "content": "Hi"}
            ],
            "max_tokens": 16
        }

        try:
            resp = await self._post_request(payload, timeout=20.0)
            if resp.status_code == 200:
                return {
                    "success": True,
                    "message": f"连接成功！(Status: 200 OK) 模型「{self.model_name}」响应正常",
                    "cleaned_base_url": self.base_url
                }
            
            # 解析常见 HTTP 状态码并给出针对性提示
            status = resp.status_code
            err_body = resp.text
            try:
                err_json = resp.json()
                err_msg = err_json.get("error", {}).get("message", err_body)
            except Exception:
                err_msg = err_body

            if status == 401:
                return {
                    "success": False,
                    "message": f"API Key 认证失败 (401 Unauthorized)。\n请检查：\n1. 复制的 API Key 是否完整无误；\n2. 当前模型提供商是否与 Key 匹配；\n3. 新创建的 Key 是否已生效。\n接口返回: {err_msg}"
                }
            elif status in [402, 429]:
                return {
                    "success": False,
                    "message": f"请求受限或账户余额不足 ({status})。\n请登录服务商后台检查账户余额与充值状态。\n接口返回: {err_msg}"
                }
            elif status == 404:
                return {
                    "success": False,
                    "message": f"接口地址或模型未找到 (404 Not Found)。\n当前请求地址: {self.base_url}/chat/completions\n模型名称: {self.model_name}\n请核对 Base URL 与模型名称是否正确。"
                }
            elif status == 400:
                return {
                    "success": False,
                    "message": f"请求参数错误 (400 Bad Request)。\n可能原因: 模型「{self.model_name}」不存在或不支持该参数。\n接口返回: {err_msg}"
                }
            else:
                return {
                    "success": False,
                    "message": f"连接失败 (HTTP {status})。\n接口返回: {err_msg}"
                }
        except Exception as e:
            return {
                "success": False,
                "message": f"网络连接异常: 无法访问 {self.base_url}。\n请检查：\n1. 电脑是否能正常联网；\n2. 如需代理请确认代理状态；\n3. 工厂局域网是否有防火墙拦截。\n错误详情: {str(e)}"
            }

    async def translate_batch(
        self,
        items: List[Dict[str, Any]],
        target_lang_name: str,
        glossary_matcher: GlossaryMatcher,
        source_lang_name: str = "中文",
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> List[Dict[str, Any]]:
        """批量翻译条目"""
        # 1. 收集批次内的术语命中
        all_matched_terms = []
        for item in items:
            src = item.get("source_text", "").strip()
            if src:
                matched = glossary_matcher.match_terms_in_text(src)
                item["matched_terms"] = matched
                all_matched_terms.extend(matched)
            else:
                item["matched_terms"] = []
                item["target_text"] = ""

        # 去重
        unique_terms = []
        seen = set()
        for t in all_matched_terms:
            if t["source"] not in seen:
                unique_terms.append(t)
                seen.add(t["source"])

        term_constraints = glossary_matcher.build_prompt_constraint(unique_terms)

        # 2. 如果没有配置有效 API Key，则使用智能离线翻译规则
        if not self.api_key or self.api_key == "sk-mock-key":
            for idx, item in enumerate(items):
                src = item.get("source_text", "").strip()
                if not src:
                    item["target_text"] = ""
                elif item.get("is_person_name"):
                    # 人名免翻译保护：直接保留原姓名
                    item["target_text"] = src
                else:
                    if item.get("matched_terms"):
                        item["target_text"] = f"[Mock {target_lang_name}] " + item["matched_terms"][0]["target"]
                    else:
                        item["target_text"] = f"[Trans {target_lang_name}] {src}"
                if progress_callback:
                    progress_callback(idx + 1, len(items))
            return items

        # 3. 构造请求给大模型进行批次翻译 (按 20 条分块)
        chunk_size = 20
        total_items = len(items)
        processed_count = 0

        for i in range(0, total_items, chunk_size):
            chunk = items[i:i + chunk_size]
            
            # 过滤出非空且非纯人名的待翻译文本 (纯人名直接保留原姓名)
            query_payload = {}
            for it in chunk:
                s_text = it.get("source_text", "").strip()
                if not s_text:
                    continue
                if it.get("is_person_name"):
                    it["target_text"] = s_text
                else:
                    query_payload[it["id"]] = s_text

            if query_payload:
                system_prompt = f"""你是一名精通制造业工程、工业制造 SOP 及质量体系（IATF 16949 / ISO）的资深工业翻译专家。
请将用户提供的 JSON 键值对中的【{source_lang_name}】内容翻译为【{target_lang_name}】。

要求：
1. 语言表达必须严谨、工业化、符合工程作业规范；
2. 保持简洁专业，不要添加任何额外的解释或引导词；
3. 【人名与签署保护】：文档中的人员姓名（如编制、审核、批准、签名栏中的中文姓名，例如：罗成耿、张三、李四等）请一律保持原样中文姓名不变，严禁将中文姓名翻译为拼音、拼音缩写或英文单词；
4. {term_constraints}
5. 必须直接返回与输入 JSON key 一一对应的 JSON 格式结果，格式如下：
{{
  "key1": "翻译结果1",
  "key2": "翻译结果2"
}}
"""
                # 首先尝试标准 JSON Mode 请求
                body = {
                    "model": self.model_name,
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": json.dumps(query_payload, ensure_ascii=False)}
                    ],
                    "temperature": 0.2,
                    "response_format": {"type": "json_object"}
                }

                parsed = {}
                try:
                    resp = await self._post_request(body, timeout=60.0)
                    if resp.status_code == 200:
                        res_json = resp.json()
                        content = res_json["choices"][0]["message"]["content"]
                        parsed = parse_json_from_response(content)
                    elif resp.status_code == 400:
                        # 降级重试：去掉 response_format 重新请求（兼容不支持 json_object 的模型）
                        body_fallback = {
                            "model": self.model_name,
                            "messages": [
                                {"role": "system", "content": system_prompt},
                                {"role": "user", "content": json.dumps(query_payload, ensure_ascii=False)}
                            ],
                            "temperature": 0.2
                        }
                        resp2 = await self._post_request(body_fallback, timeout=60.0)
                        if resp2.status_code == 200:
                            content2 = resp2.json()["choices"][0]["message"]["content"]
                            parsed = parse_json_from_response(content2)
                        else:
                            for it in chunk:
                                it["target_text"] = f"[翻译失败 {resp2.status_code}] " + it.get("source_text", "")
                    else:
                        for it in chunk:
                            it["target_text"] = f"[翻译失败 {resp.status_code}] " + it.get("source_text", "")
                except Exception as e:
                    for it in chunk:
                        it["target_text"] = f"[翻译异常: {str(e)[:30]}] " + it.get("source_text", "")

                # 填充解析出的结果
                if parsed:
                    for it in chunk:
                        it["target_text"] = parsed.get(it["id"], it["source_text"])

            processed_count += len(chunk)
            if progress_callback:
                progress_callback(min(processed_count, total_items), total_items)

        return items
