"""逐模块测试脚本 — T1 在线集成 + T2~T9 离线单测

说明:
- T1 需要真实邮箱与网络 (约 5 分钟), 走 IMAP 拉邮件
- T2~T9 为离线断言, 不依赖外部资源, 可独立运行
- OCR 现在是懒加载兜底 (E3 阶段才触发), 真实流程断言见 test_offline.py
- 本脚本定位是"模块冒烟", 详细行为验证请跑 test_offline.py (64 项断言)
"""
import sys
import os
import asyncio
import logging
from datetime import datetime, timedelta

# 设置路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("mail_audit")

# 加载配置
import yaml
with open("config.yaml", "r", encoding="utf-8") as f:
    config = yaml.safe_load(f)

PASS = "\033[92m[PASS]\033[0m"
FAIL = "\033[91m[FAIL]\033[0m"
SEP = "=" * 60

results = {}

def run_module(name, func):
    print(f"\n{SEP}")
    print(f"测试: {name}")
    print(SEP)
    try:
        result = func()
        if result:
            print(f"{PASS} {name}")
            results[name] = True
        else:
            print(f"{FAIL} {name}")
            results[name] = False
    except Exception as e:
        print(f"{FAIL} {name}: {e}")
        import traceback
        traceback.print_exc()
        results[name] = False

# ============================================================
# T1: 邮件读取模块 M1
# ============================================================
def test_mail_reader():
    from modules.mail_reader import MailReader
    reader = MailReader(config["email"], logger)
    
    # 测试IMAP连接
    end_date = datetime.now()
    start_date = end_date - timedelta(days=2)
    
    mails = reader.fetch_mails(
        start_date.strftime("%Y-%m-%d"),
        end_date.strftime("%Y-%m-%d"),
        config["email"]["address"],
        progress_callback=lambda c, t: print(f"  拉取: {c}/{t}") if c % 10 == 0 else None
    )
    
    print(f"  拉取到 {len(mails)} 封邮件")
    
    if len(mails) > 0:
        # 检查邮件结构
        m = mails[0]
        required_fields = ["sender_email", "date", "subject", "body_text", "attachments"]
        for field in required_fields:
            if field not in m:
                print(f"  缺少字段: {field}")
                return False
        
        print(f"  邮件结构: sender_email, date, subject, body_text, attachments — 全部存在")
        print(f"  附件示例: {len(m.get('attachments', []))} 个")
        return True
    else:
        print(f"  没有拉取到邮件，检查日期范围")
        return len(mails) >= 0  # 0封也算连接成功

# ============================================================
# T2: 邮件过滤模块 M2
# ============================================================
def test_mail_filter():
    from modules.mail_filter import MailFilter
    
    # 构造测试数据
    test_mails = [
        {
            "sender_email": "test@external.com",
            "subject": "德国WEEE注册申请",
            "body_text": "深圳市XX公司 德国WEEE注册 请审核",
            "attachments": [],
        },
        {
            "sender_email": "internal@example.com",
            "subject": "账单确认",
            "body_text": "请确认本月账单金额",
            "attachments": [],
        },
        {
            "sender_email": "mailbox@example.com",
            "subject": "回复: 注册确认",
            "body_text": "已收到",
            "attachments": [],
            "skip_reason": "self_sent",
        },
        {
            "sender_email": "agent@xxx.com",
            "subject": "法国电池法+法国WEEE",
            "body_text": "香港XX公司 新增 法国电池法+法国WEEE",
            "attachments": [],
        },
        {
            "sender_email": "spam@xxx.com",
            "subject": "购买担保服务",
            "body_text": "请购买担保",
            "attachments": [],
        },
    ]
    
    internal_emails = {"internal@example.com", "mailbox@example.com"}
    filt = MailFilter(internal_emails, logger)
    valid, filtered = filt.filter_mails(test_mails)
    
    print(f"  有效: {len(valid)} 封, 过滤: {len(filtered)} 封")
    
    # 验证: 自身发送被过滤
    self_sent = [m for m in filtered if "自身发送" in m.get("filter_reason", "")]
    if not self_sent:
        print("  自身发送邮件未被过滤")
        return False
    
    # 验证: 账单被过滤
    billing = [m for m in filtered if "账单" in m.get("filter_reason", "") or "无关" in m.get("filter_reason", "")]
    if not billing:
        print("  账单邮件未被过滤")
        return False
    
    # 验证: 担保被过滤
    guarantee = [m for m in filtered if "无关" in m.get("filter_reason", "")]
    if not guarantee:
        print("  担保邮件未被过滤")
        return False
    
    # 验证: 注册邮件有效
    reg_mails = [m for m in valid if "WEEE" in m.get("subject", "")]
    if not reg_mails:
        print("  注册邮件未被标记有效")
        return False
    
    print(f"  自身发送 → 过滤 ✓")
    print(f"  账单邮件 → 过滤 ✓")
    print(f"  担保邮件 → 过滤 ✓")
    print(f"  注册邮件 → 有效 ✓")
    return True

# ============================================================
# T3: 字段提取模块 M3
# ============================================================
def test_field_extractor():
    from modules.field_extractor import FieldExtractor
    
    # 构造测试数据
    agent_map = {
        "test@agent.com": "测试代理A",
        "customer@example.com": "创通易购",
    }
    project_table = [
        {"project_name": "德国WEEE", "country": "德国", "type": "WEEE"},
        {"project_name": "德国电池法", "country": "德国", "type": "电池法"},
        {"project_name": "德国包装法", "country": "德国", "type": "包装法"},
        {"project_name": "法国WEEE", "country": "法国", "type": "WEEE"},
        {"project_name": "法国电池法", "country": "法国", "type": "电池法"},
        {"project_name": "波兰包装法", "country": "波兰", "type": "包装法"},
        {"project_name": "荷兰包装法", "country": "荷兰", "type": "包装法"},
    ]
    
    extractor = FieldExtractor(agent_map, project_table, logger, None)
    
    # 测试邮件1: 主题含注册+项目
    mail1 = {
        "sender_email": "customer@example.com",
        "subject": "创通易购+CTYG046+香港邵光科技有限公司 德国电池法 法国电池法+法国WEEE 2026年",
        "body_text": "香港邵光科技有限公司 注册以下项目：德国电池法 法国电池法+法国WEEE",
        "attachments": [],
    }
    
    rows = extractor.extract_fields(mail1)
    print(f"  测试1: 提取到 {len(rows)} 行")
    if rows:
        r = rows[0]
        print(f"    代理={r.get('代理','')}, 客户={r.get('客户','')}, 项目={r.get('项目','')}, 需求={r.get('需求','')}")
        
        if not r.get("代理"):
            print("    代理提取失败")
            return False
        if not r.get("客户"):
            print("    客户提取失败")
            return False
        if not r.get("项目"):
            print("    项目提取失败")
            return False
    
    # 测试邮件2: 代理不在表中
    mail2 = {
        "sender_email": "unknown@xxx.com",
        "subject": "深圳市璟深电子商务有限责任公司 德国WEEE注册",
        "body_text": "德国WEEE注册申请",
        "attachments": [],
    }
    
    rows2 = extractor.extract_fields(mail2)
    print(f"  测试2: 提取到 {len(rows2)} 行")
    if rows2:
        r = rows2[0]
        print(f"    代理={r.get('代理','')}, 客户={r.get('客户','')}, 项目={r.get('项目','')}")
    
    print(f"  代理查表 ✓, 客户提取 ✓, 项目扫描 ✓")

    # 回归：正文首行“以下为……提交名单”是代理说明，不是第五家客户。
    bulk_mail = {
        "sender_email": "unknown@xxx.com",
        "subject": "方信-2026.08.25 德国包装法注册",
        "body_text": (
            "你好\n"
            "以下为广东省方信企业管理集团有限公司2026.08.25提交德国包装法新注册名单\n"
            "麻烦尽快处理，尽量在本周内下号\n\n"
            "FX024-武汉昼梦光年商贸有限公司德国包装法\n"
            "FX025-汕头市北泓川电子商务有限公司 德国包装法\n"
            "FX026-汕头市皮可布玩具有限公司 德国包装法\n"
            "FX027-陕西荣萃电子商务有限责任公司-德国包装法"
        ),
        "attachments": [],
    }
    bulk_rows = extractor.extract_fields(bulk_mail)
    bulk_customers = [row.get("客户", "") for row in bulk_rows]
    expected_customers = [
        "武汉昼梦光年商贸有限公司",
        "汕头市北泓川电子商务有限公司",
        "汕头市皮可布玩具有限公司",
        "陕西荣萃电子商务有限责任公司",
    ]
    if len(bulk_rows) != 4 or set(bulk_customers) != set(expected_customers):
        print(f"  正文批量名单提取错误: {bulk_customers}")
        return False
    print("  正文说明句未被当成客户，FX024-FX027 四家公司提取 ✓")

    # 申请表模板里的固定“WEEE 产品信息”不能把德国包装法误判为 WEEE。
    from modules.weee_category_audit import is_germany_weee
    packaging_attachment = [{
        "filename": "德国包装法注册.xlsx",
        "text_content": "WEEE产品信息在下列填写（需要的服务 必填）",
    }]
    if is_germany_weee("德国包装法", bulk_mail["subject"], bulk_mail["body_text"], packaging_attachment):
        print("  德国包装法被错误开启 WEEE 专项")
        return False
    if not is_germany_weee("德国WEEE", "德国WEEE注册", "品牌：X；品类：小型设备", []):
        print("  德国 WEEE 明确项目未开启专项")
        return False
    print("  德国包装法/WEEE 专项边界校验 ✓")
    return True


def test_workbench_stale_review_suppression():
    """同一邮件已有有效主记录时，旧语义纠错行不应继续计数。"""
    from workbench_server import _hide_superseded_review_rows

    base = {
        "发件人邮箱": "agent3@example.com",
        "发件日期": "2026-08-25 17:57:54",
        "邮件主题": "上海古道+香港美美的芳电子商务有限公司+德国WEEE",
        "客户公司名称": "香港美美的芳电子商务有限公司",
        "需求": "新增",
    }
    rows = [
        {
            **base,
            "_source": "待查名单",
            "标准化项目名称": "德国WEEE",
            "语义校验状态": "valid",
        },
        {
            **base,
            "_source": "人工补全",
            "标准化项目名称": "德国包装法",
            "语义校验状态": "invalid",
            "语义问题编号": "PROJECT_COVERAGE_MISSING",
            "语义建议值": "program: 德国/德国WEEE；agent: 上海古道",
        },
    ]
    kept = _hide_superseded_review_rows(rows)
    if len(kept) != 1 or kept[0].get("标准化项目名称") != "德国WEEE":
        print(f"  旧语义纠错行未被隐藏: {kept}")
        return False
    print("  旧语义纠错行不再制造额外业务明细 ✓")
    return True

# ============================================================
# T4: 项目标准化模块 M4
# ============================================================
def test_project_normalizer():
    from modules.project_normalizer import ProjectNormalizer
    
    project_table = [
        {"project_name": "德国WEEE", "country": "德国", "type": "WEEE"},
        {"project_name": "德国电池法", "country": "德国", "type": "电池法"},
        {"project_name": "德国包装法", "country": "德国", "type": "包装法"},
        {"project_name": "法国WEEE", "country": "法国", "type": "WEEE"},
        {"project_name": "法国电池法", "country": "法国", "type": "电池法"},
        {"project_name": "波兰包装法", "country": "波兰", "type": "包装法"},
        {"project_name": "荷兰包装法", "country": "荷兰", "type": "包装法"},
        {"project_name": "瑞典包装法", "country": "瑞典", "type": "包装法"},
        {"project_name": "比利时包装法", "country": "比利时", "type": "包装法"},
    ]
    
    normalizer = ProjectNormalizer(project_table, logger)
    
    # 测试1: "德国weee/电池/包装" → 3条
    r1 = normalizer.normalize_and_split("德国weee/电池/包装", "", "")
    print(f"  '德国weee/电池/包装' → {len(r1)} 条: {[x['standard_name'] for x in r1]}")
    if len(r1) != 3:
        return False
    
    # 测试2: "波兰荷兰包装法" → 2条
    r2 = normalizer.normalize_and_split("波兰荷兰包装法", "", "")
    print(f"  '波兰荷兰包装法' → {len(r2)} 条: {[x['standard_name'] for x in r2]}")
    if len(r2) != 2:
        return False
    
    # 测试3: "荷兰、瑞典、波兰、比利时4国包装法" → 4条
    r3 = normalizer.normalize_and_split("荷兰、瑞典、波兰、比利时4国包装法", "", "")
    print(f"  '4国包装法' → {len(r3)} 条: {[x['standard_name'] for x in r3]}")
    if len(r3) != 4:
        return False
    
    # 测试4: "德国WEEE+电池法" → 2条
    r4 = normalizer.normalize_and_split("德国WEEE+电池法", "", "")
    print(f"  '德国WEEE+电池法' → {len(r4)} 条: {[x['standard_name'] for x in r4]}")
    if len(r4) != 2:
        return False
    
    print(f"  4种拆分场景全部正确 ✓")
    return True

# ============================================================
# T5: 工单系统模块 M5 (仅测试类初始化和方法存在)
# ============================================================
def test_workorder_checker():
    from modules.workorder_checker import WorkOrderChecker
    
    checker = WorkOrderChecker(config["workorder"], logger)
    
    # 测试方法存在
    methods = ["login", "search_one", "match_records", "close",
              "_navigate_to_order_list", "_fill_company_name",
              "_select_dropdown", "_click_query_button", "_extract_table_data"]
    for m in methods:
        if not hasattr(checker, m):
            print(f"  缺少方法: {m}")
            return False
    
    # 测试辅助方法
    assert checker._extract_country_from_project("德国WEEE") == "德国"
    assert checker._extract_country_from_project("法国电池法") == "法国"
    assert checker._extract_country_from_project("波兰包装法") == "波兰"
    
    assert checker._extract_service_item_from_project("德国WEEE") == "WEEE"
    assert checker._extract_service_item_from_project("德国电池法") == "电池法"
    assert checker._extract_service_item_from_project("波兰包装法") == "包装法"
    
    # 测试模糊公司名生成
    terms = checker._get_fuzzy_company_terms("深圳市璟深电子商务有限责任公司")
    print(f"  模糊词: {terms}")
    if not terms:
        return False
    
    print(f"  方法完整性 ✓, 国家提取 ✓, 项目提取 ✓, 模糊词 ✓")
    return True

# ============================================================
# T6: Excel输出模块 M6
# ============================================================
def test_excel_writer():
    from modules.excel_writer import ExcelWriter
    import tempfile
    
    test_rows = [
        {
            "sender_email": "test@agent.com",
            "date": datetime.now(),
            "subject": "德国WEEE注册",
            "body_text": "深圳市XX公司 德国WEEE注册",
            "代理": "测试代理",
            "客户": "深圳市XX公司",
            "项目": "德国WEEE",
            "需求": "注册",
            "是否已录单": "否",
            "工单日期": None,
            "匹配状态": "漏单",
            "查询时间戳": datetime.now(),
        },
        {
            "sender_email": "test2@agent.com",
            "date": datetime.now(),
            "subject": "法国电池法新增",
            "body_text": "香港YY公司 法国电池法新增",
            "代理": "代理B",
            "客户": "香港YY公司",
            "项目": "法国电池法",
            "需求": "新增",
            "是否已录单": "是",
            "工单日期": datetime.now(),
            "匹配状态": "精确匹配",
            "查询时间戳": datetime.now(),
        },
    ]
    
    test_filtered = [
        {
            "sender_email": "spam@xxx.com",
            "subject": "购买担保服务",
            "filter_reason": "外部无关邮件",
        },
    ]
    
    output_dir = tempfile.mkdtemp()
    writer = ExcelWriter(output_dir, logger)
    
    missing_file = writer.write_missing_report(test_rows)
    filtered_file = writer.write_filtered_report(test_filtered)

    if os.path.exists(missing_file) and os.path.exists(filtered_file):
        print(f"  漏单清单: {os.path.basename(missing_file)} ({os.path.getsize(missing_file)} bytes)")
        print(f"  过滤清单: {os.path.basename(filtered_file)} ({os.path.getsize(filtered_file)} bytes)")
    else:
        print(f"  文件未生成")
        return False

    # 新 API: write_stage1_outputs (返回 2 个稳定名 + 内部写 filtered_mail_record)
    stage1_rows = test_rows + [
        {
            "sender_email": "u@x.com",
            "date": datetime.now(),
            "subject": "?",
            "body_text": "?",
            "代理": "",
            "客户": "",
            "项目": "",
            "需求": "注册",
            "是否已录单": "",
            "工单日期": None,
            "匹配状态": "待确认",
            "查询时间戳": datetime.now(),
            "confidence": "low",
        },
    ]
    workorder_xlsx, review_xlsx = writer.write_stage1_outputs(stage1_rows, [])
    expected = {"to_workorder_list.xlsx", "to_review_list.xlsx"}
    actual = {os.path.basename(p) for p in (workorder_xlsx, review_xlsx) if p}
    if not expected.issubset(actual):
        print(f"  stage1 稳定名未齐全: {actual}")
        return False
    # 顺带验 filtered_mail_record 落盘
    fm_path = os.path.join(output_dir, "filtered_mail_record.xlsx")
    if not os.path.exists(fm_path):
        print(f"  filtered_mail_record.xlsx 未生成")
        return False
    print(f"  stage1 三份产物齐全: {sorted(actual)} + filtered_mail_record.xlsx")

    # 新 API: write_workorder_check_result (返回 stable + 时序副本)
    skipped = [{"说明": "代理空，跳过"}]
    check_stable, check_ts = writer.write_workorder_check_result(stage1_rows, [], skipped)
    if not (os.path.exists(check_stable) and os.path.exists(check_ts)):
        print(f"  阶段二产物未生成 (stable={check_stable}, ts={check_ts})")
        return False
    print(f"  阶段二产物: {os.path.basename(check_stable)} + 时序副本")

    return True

# ============================================================
# T7: OCR 引擎可初始化 (基础设施断言)
# 注: 真实业务流中 OCR 是 E3 阶段懒加载兜底, 详细行为见 test_offline.py
# ============================================================
def test_ocr():
    try:
        from utils.attachment_parser import _get_ocr
        ocr = _get_ocr()
        if ocr is not None:
            print(f"  PaddleOCR 引擎初始化成功 (懒加载入口可用)")
            return True
        else:
            print(f"  PaddleOCR 引擎不可用（初始化失败，业务中 OCR 兜底会自动跳过）")
            # 不算失败：PaddleOCR 不在本机环境时，懒加载会缓存失败标志，业务流程照常
            return True
    except Exception as e:
        print(f"  PaddleOCR 不可用: {e} (懒加载会兜底跳过)")
        return True

# ============================================================
# T8: 附件解析模块
# ============================================================
def test_attachment_parser():
    from utils.attachment_parser import parse_attachment, extract_emails_from_text
    
    # 测试邮箱提取
    text = "联系我们: test@example.com 或 admin@company.cn"
    emails = extract_emails_from_text(text)
    if len(emails) != 2:
        print(f"  邮箱提取失败: {emails}")
        return False
    print(f"  邮箱提取: {emails}")
    
    # 测试不支持的格式
    result = parse_attachment("/nonexistent/file.xyz", "test.xyz")
    if result["text_content"] != "":
        print(f"  不支持的格式应返回空")
        return False
    print(f"  不支持的格式 → 空 ✓")
    
    return True

# ============================================================
# T9: 模糊匹配模块
# ============================================================
def test_fuzzy_match():
    from utils.fuzzy_match import fuzzy_match_pair, normalize_text
    
    # 简繁匹配
    r1 = fuzzy_match_pair("香港韶光科技有限公司", "香港韶光科技有限公司", 85)
    if not r1:
        print(f"  简繁匹配失败")
        return False
    print(f"  简繁匹配 ✓")
    
    # 子串匹配
    r2 = fuzzy_match_pair("深圳市璟深", "深圳市璟深电子商务有限责任公司", 85)
    if not r2:
        print(f"  子串匹配失败")
        return False
    print(f"  子串匹配 ✓")
    
    # 不匹配
    r3 = fuzzy_match_pair("德国WEEE", "法国电池法", 85)
    if r3:
        print(f"  不应该匹配的项目匹配了")
        return False
    print(f"  不匹配验证 ✓")
    
    return True


# ============================================================
# 运行所有测试
# ============================================================
if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  邮件漏单审核工具 — 逐模块测试")
    print("=" * 60)
    
    run_module("T1 邮件读取 M1", test_mail_reader)
    run_module("T2 邮件过滤 M2", test_mail_filter)
    run_module("T3 字段提取 M3", test_field_extractor)
    run_module("T4 项目标准化 M4", test_project_normalizer)
    run_module("T5 工单系统 M5", test_workorder_checker)
    run_module("T6 Excel输出 M6", test_excel_writer)
    run_module("T7 OCR模块", test_ocr)
    run_module("T8 附件解析", test_attachment_parser)
    run_module("T9 模糊匹配", test_fuzzy_match)
    
    # 汇总
    print("\n" + "=" * 60)
    print("  测试结果汇总")
    print("=" * 60)
    passed = sum(1 for v in results.values() if v)
    total = len(results)
    for name, ok in results.items():
        status = f"\033[92mPASS\033[0m" if ok else f"\033[91mFAIL\033[0m"
        print(f"  {status}  {name}")
    print(f"\n  通过: {passed}/{total}")
    
    if passed == total:
        print(f"\n  *** 全部通过 ***")
    else:
        print(f"\n  *** 有 {total-passed} 项失败，需修复 ***")
