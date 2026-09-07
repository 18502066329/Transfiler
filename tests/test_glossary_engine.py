import pytest
from backend.core.glossary_matcher import GlossaryMatcher
from backend.db.database import (
    init_db,
    add_glossary_term,
    get_glossary_terms,
    delete_glossary_term,
    get_enabled_glossary_dict
)

def test_glossary_matcher():
    glossary = {
        "射胶压力": "Injection Pressure",
        "防错治具": "Poka-Yoke Fixture",
        "合模机构": "Clamping Mechanism"
    }
    matcher = GlossaryMatcher(glossary)

    sample_text = "操作员必须检查射胶压力并在工位安装防错治具。"
    matched = matcher.match_terms_in_text(sample_text)

    assert len(matched) == 2
    sources = [m["source"] for m in matched]
    assert "射胶压力" in sources
    assert "防错治具" in sources

    prompt_constraint = matcher.build_prompt_constraint(matched)
    assert "射胶压力" in prompt_constraint
    assert "Injection Pressure" in prompt_constraint

def test_glossary_db_crud():
    init_db()
    
    # 增加
    term_id = add_glossary_term(source_term="测试工艺词条", target_term="Test Process Term", target_lang="en", category="测试分类")
    assert term_id > 0

    # 查询
    terms = get_glossary_terms(keyword="测试工艺词条")
    assert len(terms) > 0
    assert terms[0]["source_term"] == "测试工艺词条"

    # 删除
    delete_glossary_term(term_id)
    terms_after = get_glossary_terms(keyword="测试工艺词条")
    assert len(terms_after) == 0
