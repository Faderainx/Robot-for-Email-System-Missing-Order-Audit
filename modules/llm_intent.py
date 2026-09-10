"""LLM 意图识别模块 — DeepSeek API 辅助邮件分类与字段提取

两个核心功能:
1. classify_email: 对规则过滤掉的邮件做二次意图识别, 找回被误杀的注册类邮件
2. extract_fields_llm: 对规则提取失败的字段做 LLM 补充提取
"""
import json
import time
from typing import Dict, List, Optional

from openai import OpenAI


SYSTEM_PROMPT_CLASSIFY = """你是一个邮件分类助手，服务于一家处理欧盟环保合规注册的公司（业务包括 WEEE、电池法、包装法、EPR、一次性塑料等）。

你的任务是判断邮件是否属于以下类型之一：
- "注册类"：客户/代理发来要求注册、新增、撤销环保合规项目
- "账单"：关于费用账单、发票
- "合同"：合同相关
- "咨询"：业务咨询、问题反馈
- "下证"：证书下发通知
- "购买"：购买担保、购买回收费
- "其他"：不属于以上类型

判断标准：
- 邮件主题或正文中提到"注册""新增""撤单"等关键词，且涉及环保合规项目（WEEE/电池法/包装法/EPR等）→ 注册类
- 邮件提到"账单""invoice""费用"→ 账单
- 邮件提到"合同""协议"→ 合同
- 邮件提到"咨询""问题""反馈"→ 咨询
- 邮件提到"下证""证书""下号"→ 下证
- 邮件提到"购买""担保""回收费"→ 购买
- 其他情况 → 其他

你必须只返回 JSON，不要输出任何其他文字。"""


SYSTEM_PROMPT_EXTRACT = """你是一个信息提取助手，服务于一家处理欧盟环保合规注册的公司。

从邮件内容中提取以下字段：
1. "代理"：处理注册的中间代理机构名称（不是客户公司名，是代理/中介公司名）。如果邮件中没有明确提到代理名称，返回空字符串。
2. "客户"：需要注册的客户公司名称（通常含"公司""有限""科技""电商""贸易"等字样）。
3. "项目"：邮件中提到的环保合规项目名称数组。标准格式为"国家+项目类型"，例如"德国WEEE""法国电池法""波兰包装法""荷兰包装法"等。组合项目要拆开，例如"德国WEEE+电池法"应拆为["德国WEEE","电池法"]。注意"电池法"可能指"德国电池法"等，根据上下文判断国家。
4. "需求"：邮件意图类型，"注册"/"新增"/"撤单"之一。如果无法判断，返回空字符串。

你必须只返回 JSON，不要输出任何其他文字。"""


class LLMIntentClient:
    """DeepSeek LLM 客户端"""

    def __init__(self, config: dict, logger=None):
        """
        config: {
            "api_key": str,
            "base_url": str,
            "model": str,
            "timeout": int,
            "max_tokens": int,
        }
        """
        self.api_key = config.get("api_key", "")
        self.base_url = config.get("base_url", "https://api.deepseek.com")
        self.model = config.get("model", "deepseek-chat")
        self.timeout = config.get("timeout", 30)
        self.max_tokens = config.get("max_tokens", 1000)
        self.temperature = config.get("temperature", 0.1)
        self.enabled = bool(self.api_key)
        self.logger = logger

        if self.enabled:
            self.client = OpenAI(
                api_key=self.api_key,
                base_url=self.base_url,
                timeout=self.timeout,
            )
        else:
            self.client = None

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def _call_api(self, system_prompt: str, user_content: str) -> Optional[dict]:
        """调用 LLM API，返回解析后的 JSON dict"""
        if not self.enabled:
            return None

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_content},
                ],
                response_format={"type": "json_object"},
                temperature=self.temperature,
                max_tokens=self.max_tokens,
            )
            raw = response.choices[0].message.content.strip()
            if not raw:
                self._log("LLM 返回空内容", "warning")
                return None
            return json.loads(raw)
        except json.JSONDecodeError as e:
            self._log(f"LLM 返回 JSON 解析失败: {e} | 原始: {raw[:200]}", "error")
            return None
        except Exception as e:
            self._log(f"LLM API 调用失败: {e}", "error")
            return None

    @staticmethod
    def _build_email_content(subject: str, body: str, attachments: list) -> str:
        """构造发给 LLM 的邮件内容文本"""
        parts = [f"【邮件主题】\n{subject}"]
        if body:
            truncated_body = body[:2000] if len(body) > 2000 else body
            parts.append(f"【邮件正文】\n{truncated_body}")

        att_texts = []
        for att in attachments:
            fname = att.get("filename", "")
            text = att.get("text_content", "")
            if text:
                truncated = text[:1000] if len(text) > 1000 else text
                att_texts.append(f"[{fname}] {truncated}")
        if att_texts:
            parts.append(f"【附件内容】\n" + "\n".join(att_texts))

        return "\n\n".join(parts)

    def classify_email(
        self, subject: str, body: str, attachments: list
    ) -> Optional[dict]:
        """
        对单封邮件做意图分类
        返回: {
            "is_target": bool,      # 是否注册类邮件
            "intent": str,          # "注册"/"新增"/"撤单"/""
            "email_type": str,      # "注册类"/"账单"/"合同"/"咨询"/"下证"/"购买"/"其他"
            "reason": str,          # 简要原因
        }
        失败返回 None
        """
        content = self._build_email_content(subject, body, attachments)
        result = self._call_api(
            SYSTEM_PROMPT_CLASSIFY,
            content + '\n\n请分析以上邮件，返回 JSON：{"is_target": true/false, "intent": "注册/新增/撤单/空", "email_type": "类型", "reason": "简要原因"}',
        )
        if result is None:
            return None

        return {
            "is_target": bool(result.get("is_target", False)),
            "intent": result.get("intent", ""),
            "email_type": result.get("email_type", ""),
            "reason": result.get("reason", ""),
        }

    def batch_classify(self, mails: List[dict]) -> Dict[str, dict]:
        """
        批量分类邮件
        mails: [{uid, subject, body_text, attachments, ...}, ...]
        返回: {uid: classify_result}
        """
        results = {}
        total = len(mails)
        for idx, mail in enumerate(mails, 1):
            uid = mail.get("uid", str(idx))
            subject = mail.get("subject", "") or ""
            body = mail.get("body_text", "") or ""
            attachments = mail.get("attachments", [])

            self._log(f"LLM 分类 {idx}/{total}: {subject[:40]}")
            result = self.classify_email(subject, body, attachments)
            if result:
                results[uid] = result
            else:
                results[uid] = {"is_target": None, "intent": "", "email_type": "LLM失败", "reason": "API调用失败"}

            time.sleep(0.5)

        return results

    def extract_fields_llm(
        self, subject: str, body: str, attachments: list, sender_email: str = ""
    ) -> Optional[dict]:
        """
        LLM 提取结构化字段
        返回: {
            "代理": str,
            "客户": str,
            "项目": [str, ...],
            "需求": str,
            "confidence": str,
        }
        失败返回 None
        """
        content = self._build_email_content(subject, body, attachments)
        if sender_email:
            content = f"【发件人邮箱】{sender_email}\n\n" + content

        result = self._call_api(
            SYSTEM_PROMPT_EXTRACT,
            content + '\n\n请提取以上邮件中的字段，返回 JSON：{"代理": "", "客户": "", "项目": [], "需求": "", "confidence": "high/medium/low"}',
        )
        if result is None:
            return None

        return {
            "代理": result.get("代理", ""),
            "客户": result.get("客户", ""),
            "项目": result.get("项目", []),
            "需求": result.get("需求", ""),
            "confidence": result.get("confidence", "medium"),
        }
