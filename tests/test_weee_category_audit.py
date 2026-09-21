import json

from modules.weee_category_audit import (
    classify_weee_product,
    compare_weee_categories,
    extract_weee_items,
)


def test_classification_table_rules_are_conservative():
    assert classify_weee_product("手机")["category_class"] == "6"
    assert classify_weee_product("空调")["category_class"] == "1"
    assert classify_weee_product("LED灯")["category_class"] == "3"
    # 传统白炽灯在文件中明确不属于第 3 类，不能被灯具关键词误收。
    assert classify_weee_product("传统白炽灯")["status"] == "unmatched"
    # 未提供尺寸的“打印机”无法在大/小/IT 设备之间确定，不自动猜类别。
    assert classify_weee_product("打印机")["status"] == "ambiguous"


def test_mail_brand_and_category_are_paired():
    result = extract_weee_items(
        "德国 WEEE 品类清单",
        "品牌：ACME 类别：手机",
        [],
        "德国WEEE",
    )
    assert result["enabled"] is True
    assert result["status"] == "ready"
    assert result["items"] == [
        {
            "brand": "ACME",
            "category": "手机",
            "sources": ["邮件正文/标题"],
            "evidences": ["品牌：ACME；类别：手机"],
            "confidence": "medium",
            "evidence": "品牌：ACME；类别：手机",
            "category_original": "手机",
            "category_class": "6",
            "category_class_name": "小型信息和电信设备",
            "category_class_status": "matched",
            "category_candidates": [
                {"category_class": "6", "category_class_name": "小型信息和电信设备", "score": 12}
            ],
            "category_class_evidence": ["手机"],
            "category_rule_source": "产品分类表 中文版.docx（德国 ElektroG/WEEE 六类）",
        }
    ]


def test_workorder_category_detail_is_compared_by_class_and_brand():
    mail = extract_weee_items(
        "德国 WEEE",
        "品牌：ACME；品类：手机",
        [],
        "德国WEEE",
    )
    matched = compare_weee_categories(
        mail["items"],
        [{"品类明细": [{"品类": "小型信息和电信设备"}], "品牌": "ACME"}],
    )
    assert matched["status"] == "matched"
    assert matched["items"][0]["status"] == "matched"

    missing = compare_weee_categories(
        mail["items"],
        [{"品类明细": [{"品类": "屏幕和显示设备"}], "品牌": "ACME"}],
    )
    assert missing["status"] == "missing"

    pending = compare_weee_categories(mail["items"], [{"工单编号": "WO-1"}])
    assert pending["status"] == "pending"
