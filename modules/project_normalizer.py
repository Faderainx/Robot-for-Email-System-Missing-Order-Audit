"""M4 项目标准化与拆分模块 — 将邮件中的项目简写/组合拆分为标准项目名称"""
import re
from typing import List, Dict

from utils.fuzzy_match import normalize_text, fuzzy_search


# 国家列表(从附件四提取)
COUNTRIES = [
    "德国", "荷兰", "爱尔兰", "意大利", "比利时", "波兰", "丹麦",
    "法国", "捷克", "葡萄牙", "瑞典", "西班牙", "卢森堡", "奥地利",
    "匈牙利", "芬兰", "罗马尼亚", "挪威", "爱沙尼亚", "瑞士",
    "希腊", "英国", "加拿大", "拉脱维亚",
]

# 业务类型
BUSINESS_TYPES = ["WEEE", "电池法", "包装法", "一次性塑料"]

# 组合项目映射 (附件四中以11/12/13/14结尾的组合项目)
COMBO_MAP = {
    "荷兰WEEE+包装法": ["荷兰WEEE", "荷兰包装法"],
    "荷兰电池法+包装法": ["荷兰电池法", "荷兰包装法"],
    "荷兰WEEE+电池法+包装法": ["荷兰WEEE", "荷兰电池法", "荷兰包装法"],
    "爱尔兰WEEE+电池法": ["爱尔兰WEEE", "爱尔兰电池法"],
    "意大利WEEE+电池法": ["意大利WEEE", "意大利电池法"],
    "波兰WEEE+波兰包装法": ["波兰WEEE", "波兰包装法"],
    "波兰电池法+波兰包装法": ["波兰电池法", "波兰包装法"],
    "波兰WEEE+波兰电池法+波兰包装法": ["波兰WEEE", "波兰电池法", "波兰包装法"],
    "捷克WEEE+电池法": ["捷克WEEE", "捷克电池法"],
    "匈牙利WEEE+匈牙利电池法": ["匈牙利WEEE", "匈牙利电池法"],
    "匈牙利WEEE+匈牙利包装法": ["匈牙利WEEE", "匈牙利包装法"],
    "匈牙利电池法+匈牙利包装法": ["匈牙利电池法", "匈牙利包装法"],
    "匈牙利WEEE+匈牙利电池法+匈牙利包装法": ["匈牙利WEEE", "匈牙利电池法", "匈牙利包装法"],
    "挪威WEEE/电池法/包装法": ["挪威WEEE", "挪威电池法", "挪威包装法"],
}


class ProjectNormalizer:
    def __init__(self, project_table: List[Dict], logger=None):
        """
        project_table: 附件四完整列表
        [{"项目编号": str, "项目名称": str, "国家": str, "业务类型": str}, ...]
        """
        self.project_table = project_table
        self.standard_names = [p["项目名称"] for p in project_table if p.get("项目名称")]
        self.project_lookup = {p["项目名称"]: p for p in project_table if p.get("项目名称")}
        self.logger = logger

        self.project_candidates = [
            (name, name) for name in self.standard_names
        ]

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def normalize_and_split(self, raw_value: str, subject: str = "", body: str = "") -> List[Dict]:
        """
        标准化并拆分项目
        raw_value: 邮件中提取的项目原始文本
        subject/body: 用于推断国家上下文
        返回: [{"raw_value": str, "standard_name": str, "matched": bool}, ...]
        """
        if not raw_value:
            return []

        # 如果已经是标准名称，直接返回
        if raw_value in self.project_lookup:
            return [{"raw_value": raw_value, "standard_name": raw_value, "matched": True}]

        # 尝试组合项目映射
        combo_result = self._check_combo(raw_value)
        if combo_result:
            return combo_result

        # 尝试分隔符拆分
        split_parts = self._split_by_delimiters(raw_value)
        if len(split_parts) > 1:
            results = []
            # 用原始值作为额外上下文, 这样 "荷兰" 段能从 "荷兰、瑞典、波兰、比利时4国包装法" 知道业务类型是"包装法"
            full_context = raw_value + " " + subject + " " + body
            for part in split_parts:
                results.extend(self._normalize_single(part, full_context, body))
            if results:
                return results

        # 单个值标准化
        results = self._normalize_single(raw_value, subject, body)
        if results:
            return results

        # 模糊匹配
        match = fuzzy_search(raw_value, self.project_candidates, high_threshold=85, medium_threshold=75)
        if match["matched"]:
            return [{
                "raw_value": raw_value,
                "standard_name": match["results"][0]["text"],
                "matched": match["certain"],
            }]

        # 无法匹配
        self._log(f"项目标准化失败: {raw_value}", "warning")
        return [{"raw_value": raw_value, "standard_name": raw_value, "matched": False}]

    def _check_combo(self, raw_value: str) -> List[Dict]:
        """检查是否是已知的组合项目"""
        norm = normalize_text(raw_value)
        for combo_name, sub_items in COMBO_MAP.items():
            if normalize_text(combo_name) == norm:
                return [
                    {"raw_value": raw_value, "standard_name": item, "matched": True}
                    for item in sub_items
                ]
        return []

    def _split_by_delimiters(self, raw_value: str) -> List[str]:
        """按分隔符拆分"""
        # 按多种分隔符拆分
        parts = re.split(r"[+＋/／、，,]", raw_value)
        parts = [p.strip() for p in parts if p.strip()]

        if len(parts) > 1:
            return parts

        # 尝试国家组合模式: "波兰荷兰包装法" → ["波兰包装法", "荷兰包装法"]
        # 或 "荷兰、瑞典、波兰、比利时4国包装法" → 4条
        countries_found = []
        for c in COUNTRIES:
            if c in raw_value and c not in countries_found:
                countries_found.append(c)

        # 去掉国家后剩余部分作为业务类型
        if len(countries_found) > 1:
            remaining = raw_value
            for c in countries_found:
                remaining = remaining.replace(c, "")
            # 去除"N国"模式
            remaining = re.sub(r"\d+国", "", remaining)
            remaining = remaining.strip()

            if remaining and any(bt in remaining for bt in BUSINESS_TYPES):
                results = []
                for c in countries_found:
                    candidate = c + remaining
                    # 检查是否是标准名称
                    if candidate in self.project_lookup:
                        results.append(candidate)
                    else:
                        # 模糊匹配
                        match = fuzzy_search(
                            candidate, self.project_candidates,
                            high_threshold=85, medium_threshold=75
                        )
                        if match["matched"] and match["certain"]:
                            results.append(match["results"][0]["text"])
                        else:
                            results.append(candidate)

                if len(results) > 1:
                    return results

        return parts

    def _normalize_single(self, value: str, subject: str = "", body: str = "") -> List[Dict]:
        """标准化单个项目值"""
        value = value.strip()
        if not value:
            return []

        # 精确匹配
        if value in self.project_lookup:
            return [{"raw_value": value, "standard_name": value, "matched": True}]

        # 模糊匹配
        match = fuzzy_search(value, self.project_candidates, high_threshold=90, medium_threshold=80)
        if match["matched"] and match["certain"]:
            return [{
                "raw_value": value,
                "standard_name": match["results"][0]["text"],
                "matched": True,
            }]

        # 检查是否是纯业务类型 (如"电池"、"包装"、"weee")，需要补国家前缀
        context = subject + " " + body
        value_lower = value.lower()

        # 判断 value 是否是某个业务类型(或其简写)
        bt_aliases = {
            "WEEE": ["weee", "weee"],
            "电池法": ["电池", "电池法", "battery"],
            "包装法": ["包装", "包装法", "packaging"],
            "一次性塑料": ["一次性塑料", "塑料"],
        }

        matched_bt = None
        for bt, aliases in bt_aliases.items():
            for alias in aliases:
                if alias.lower() == value_lower:
                    matched_bt = bt
                    break
            if matched_bt:
                break

        if matched_bt:
            # 从上下文中找国家
            countries_in_context = [c for c in COUNTRIES if c in context]
            if not countries_in_context:
                # 从 value 自身找国家 (如 "德国电池" 中的 "德国")
                for c in COUNTRIES:
                    if c in value:
                        countries_in_context = [c]
                        break

            results = []
            for country in countries_in_context:
                candidate = country + matched_bt
                if candidate in self.project_lookup:
                    results.append({"raw_value": value, "standard_name": candidate, "matched": True})
            if results:
                return results

            # 模糊匹配补全
            for country in countries_in_context:
                candidate = country + matched_bt
                match2 = fuzzy_search(candidate, self.project_candidates, high_threshold=80, medium_threshold=70)
                if match2["matched"]:
                    results.append({
                        "raw_value": value,
                        "standard_name": match2["results"][0]["text"],
                        "matched": match2["certain"],
                    })
            if results:
                return results

        # 特殊处理: value 是纯国家名 (如 "荷兰" "瑞典"), 从上下文推断业务类型
        is_pure_country = False
        for c in COUNTRIES:
            if value == c:
                is_pure_country = True
                break

        if is_pure_country:
            # 从上下文中找业务类型
            context_bt = None
            for bt in BUSINESS_TYPES:
                if bt in context:
                    context_bt = bt
                    break
            if context_bt:
                candidate = value + context_bt
                if candidate in self.project_lookup:
                    return [{"raw_value": value, "standard_name": candidate, "matched": True}]
                match3 = fuzzy_search(candidate, self.project_candidates, high_threshold=80, medium_threshold=70)
                if match3["matched"]:
                    return [{
                        "raw_value": value,
                        "standard_name": match3["results"][0]["text"],
                        "matched": match3["certain"],
                    }]

        # 尝试从主题推断国家 (兼容旧逻辑)
        country = None
        for c in COUNTRIES:
            if c in context:
                country = c
                break

        if country:
            for bt in BUSINESS_TYPES:
                if bt.lower() in value_lower or bt in value:
                    candidate = country + bt
                    if candidate in self.project_lookup:
                        return [{"raw_value": value, "standard_name": candidate, "matched": True}]

        return []
