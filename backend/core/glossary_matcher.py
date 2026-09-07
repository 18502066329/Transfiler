import re
from typing import Dict, List, Tuple, Any

class GlossaryMatcher:
    def __init__(self, glossary_dict: Dict[str, str] = None):
        """
        glossary_dict: { "射胶压力": "Injection Pressure", "合模机构": "Clamping Mechanism", ... }
        """
        self.glossary = glossary_dict or {}
        # 按词长倒序排列，优先匹配最长专有名词
        self.sorted_terms = sorted(self.glossary.keys(), key=lambda x: len(x), reverse=True)

    def update_glossary(self, glossary_dict: Dict[str, str]):
        self.glossary = glossary_dict or {}
        self.sorted_terms = sorted(self.glossary.keys(), key=lambda x: len(x), reverse=True)

    def match_terms_in_text(self, text: str) -> List[Dict[str, str]]:
        """
        在文本中查找所有匹配到的术语
        返回: [ {"source": "射胶压力", "target": "Injection Pressure"}, ... ]
        """
        if not text or not self.glossary:
            return []
        
        matched = []
        seen_sources = set()
        for term in self.sorted_terms:
            if term in text and term not in seen_sources:
                matched.append({
                    "source": term,
                    "target": self.glossary[term]
                })
                seen_sources.add(term)
        return matched

    def build_prompt_constraint(self, matched_terms: List[Dict[str, str]]) -> str:
        """
        构造针对大模型的专属术语约束指令
        """
        if not matched_terms:
            return ""
        
        constraint = "\n【强制专业术语对照规则】在翻译中遇到以下中文术语时，必须严格使用对应的外文翻译，不得自作主张同义改写：\n"
        for item in matched_terms:
            constraint += f"- 「{item['source']}」 => 「{item['target']}」\n"
        return constraint

    def apply_exact_replacement(self, text: str, translation: str) -> str:
        """
        后置校验：如果特定非常精确的专有名词译文中缺失，可做温和后处理校准
        """
        return translation
