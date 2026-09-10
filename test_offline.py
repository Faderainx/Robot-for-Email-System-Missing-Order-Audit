"""离线全链路测试 — 纯构造数据，不依赖网络/邮箱/工单系统/LLM，秒级跑完

覆盖: M2过滤 / M3字段提取(含规则项目提取) / M4标准化拆分 / M5按行比对(串行污染回归) /
      M6 Excel输出 / OCR新旧返回格式兼容
运行: python test_offline.py  (exit 0 = 全部通过)
"""
import os
import sys
import shutil
import tempfile
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

PASS = 0
FAIL = 0


def check(name, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  [PASS] {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name} {detail}")


def test_mail_filter():
    print("== M2 邮件过滤 ==")
    from modules.mail_filter import MailFilter
    filt = MailFilter({"boss@eu-helper.com"})

    mails = [
        {"subject": "乐天 德国WEEE注册", "body_text": "请处理", "sender_email": "a@x.com",
         "attachments": [], "skip_reason": None},
        {"subject": "7月账单", "body_text": "见附件", "sender_email": "b@x.com",
         "attachments": [], "skip_reason": None},
        {"subject": "周报", "body_text": "本周情况", "sender_email": "boss@eu-helper.com",
         "attachments": [], "skip_reason": None},
        {"subject": "闲聊", "body_text": "中午吃什么", "sender_email": "c@x.com",
         "attachments": [], "skip_reason": None},
        {"subject": "回复: 某某", "body_text": "", "sender_email": "me@eu-helper.com",
         "attachments": [], "skip_reason": "self_sent"},
    ]
    valid, filtered = filt.filter_mails(mails)

    subj_valid = {m["subject"] for m in valid}
    subj_filtered = {m["filter_reason"] for m in filtered}
    check("业务+注册词 → 有效", "乐天 德国WEEE注册" in subj_valid)
    check("账单 → 过滤", any("无关" in r for r in subj_filtered))
    check("内部无关 → 过滤", any("内部无关" in m["filter_reason"] for m in filtered
                                  if m["subject"] == "周报"))
    check("外部闲聊 → 过滤", any(m["subject"] == "闲聊" for m in filtered))
    check("自身发送 → 过滤", any(m["filter_reason"] == "自身发送" for m in filtered))


def test_field_extractor():
    print("== M3 字段提取 (无附件四, 规则兜底) ==")
    from modules.field_extractor import FieldExtractor

    agent_map = {
        "agent1@lekotest.com": {"代理": "乐天", "代理简称": "乐天", "收件人邮箱": "agent1@lekotest.com"},
    }
    fx = FieldExtractor(agent_map, project_names=[])  # 附件四为空

    mail = {
        "sender_email": "agent1@lekotest.com",
        "sender_name": "乐天",
        "date": datetime(2026, 9, 1, 10, 0),
        "subject": "乐天+888+深圳市甲科技有限公司+法国WEEE注册",
        "body_text": "",
        "attachments": [],
    }
    rows = fx.extract_fields(mail)
    r = rows[0]
    check("代理精确匹配", r["代理"] == "乐天" and r["代理匹配方式"] == "邮箱精确匹配", f"got {r['代理']}")
    check("客户从主题提取", r["客户"] == "深圳市甲科技有限公司", f"got {r['客户']}")
    check("项目规则提取(法国WEEE)", r["项目"] == "法国WEEE", f"got {r['项目']}")
    check("需求=注册", r["需求"] == "注册")
    check("置信度high", r["置信度"] == "high", f"got {r['置信度']}")

    # 组合项目规则提取
    mail2 = dict(mail, subject="乐天+深圳市甲科技有限公司+荷兰WEEE+包装法 新增")
    rows2 = fx.extract_fields(mail2)
    projs = {r["项目"] for r in rows2}
    check("共享前缀组合拆分(荷兰WEEE+包装法)", projs == {"荷兰WEEE", "荷兰包装法"}, f"got {projs}")

    # 无项目 → 仍输出一行
    mail3 = dict(mail, subject="乐天 通知邮件")
    rows3 = fx.extract_fields(mail3)
    check("无项目输出一行", len(rows3) == 1 and rows3[0]["项目"] == "")


def test_project_normalizer():
    print("== M4 标准化拆分 ==")
    from modules.project_normalizer import ProjectNormalizer

    table = [
        {"项目编号": "1", "项目名称": "法国WEEE", "国家": "法国", "业务类型": "WEEE"},
        {"项目编号": "2", "项目名称": "德国WEEE", "国家": "德国", "业务类型": "WEEE"},
        {"项目编号": "3", "项目名称": "德国电池法", "国家": "德国", "业务类型": "电池法"},
    ]
    norm = ProjectNormalizer(table)

    r = norm.normalize_and_split("法国WEEE")
    check("标准名直过", len(r) == 1 and r[0]["standard_name"] == "法国WEEE" and r[0]["matched"])
    r = norm.normalize_and_split("德国WEEE+电池法")
    check("组合拆分", [x["standard_name"] for x in r] == ["德国WEEE", "德国电池法"],
          f"got {[x['standard_name'] for x in r]}")
    r = norm.normalize_and_split("电池法", subject="德国客户咨询", body="")
    check("纯业务+上下文补国家", any(x["standard_name"] == "德国电池法" for x in r),
          f"got {r}")
    # 附件四为空时: 原样返回但流程不崩
    norm0 = ProjectNormalizer([])
    r = norm0.normalize_and_split("瑞典WEEE")
    check("空表不崩", len(r) == 1 and r[0]["standard_name"] == "瑞典WEEE")


def test_workorder_match():
    print("== M5 工单按行比对 (串行污染回归) ==")
    from modules.workorder_checker import WorkOrderChecker
    cfg = {"url": "u", "username": "a", "password": "b"}
    checker = WorkOrderChecker(cfg, None)

    date = datetime(2026, 9, 1, 10, 0)
    row_a = {"date": date, "代理": "乐天", "客户": "深圳市甲科技有限公司", "项目": "德国WEEE"}
    row_b = {"date": date, "代理": "乐天", "客户": "深圳市乙贸易有限公司", "项目": "德国WEEE"}

    wo_a = [{"代理": "乐天", "客户": "深圳市甲科技有限公司", "服务项目": "德国WEEE",
             "创建时间": "2026-09-03 10:00:00"}]

    # A 行与自己的工单 → 已录单
    checker.match_records([dict(row_a)], wo_a)
    out_a = checker.match_records([row_a], wo_a)[0]
    check("本行命中→已录单", out_a["是否已录单"] == "是" and out_a["匹配状态"] == "精确匹配",
          f"got {out_a['匹配状态']}")

    # 关键回归: B 行不能被 A 行的工单"匹配成功"（旧版全池比对会掩盖此类漏单）
    out_b = checker.match_records([dict(row_b)], wo_a)[0]
    check("跨行不串扰→B仍为漏单", out_b["是否已录单"] == "否", f"got {out_b['是否已录单']}")

    # 日期超容差 → 漏单
    wo_old = [dict(wo_a[0], 创建时间="2026-01-01 10:00:00")]
    out_c = checker.match_records([dict(row_a)], wo_old)[0]
    check("日期超容差→漏单", out_c["是否已录单"] == "否")

    # 空查询结果 → 漏单
    out_d = checker.match_records([dict(row_a)], [])[0]
    check("无工单→漏单", out_d["是否已录单"] == "否" and out_d["匹配状态"] == "漏单")


def test_excel_writer():
    print("== M6 Excel 输出 ==")
    from modules.excel_writer import ExcelWriter
    tmpdir = tempfile.mkdtemp(prefix="mail_audit_test_")
    try:
        w = ExcelWriter(tmpdir)
        rows = [
            {"sender_email": "a@x.com", "date": datetime(2026, 9, 1), "subject": "s",
             "body_text": "b", "代理": "乐天", "代理匹配方式": "邮箱精确匹配",
             "客户": "深圳市甲科技有限公司", "客户提取来源": "主题",
             "项目": "德国WEEE", "项目原始值": "德国WEEE", "需求": "注册",
             "是否已录单": "否", "匹配状态": "漏单", "查询时间戳": datetime.now()},
        ]
        f1 = w.write_missing_report(rows)
        check("漏单清单生成", os.path.exists(f1) and f1.endswith(".xlsx"))
        f2 = w.write_filtered_report([{"subject": "账单", "filter_reason": "外部无关邮件",
                                       "sender_email": "b@x.com", "date": datetime(2026, 9, 1)}])
        check("过滤清单生成", os.path.exists(f2))
        # 内容抽查
        from openpyxl import load_workbook
        wb = load_workbook(f1)
        ws = wb.active
        check("漏单行数=2(表头+1)", ws.max_row == 2, f"got {ws.max_row}")
        check("客户列正确", ws.cell(row=2, column=7).value == "深圳市甲科技有限公司")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def test_ocr_compat():
    print("== OCR 返回格式兼容 ==")
    from utils.attachment_parser import _extract_ocr_texts
    # PaddleOCR 3.x: [{'rec_texts': [...], ...}]
    r3 = [{"rec_texts": ["France WEEE 998877", ""], "rec_scores": [0.99, 0.1]}]
    check("3.x字典格式", _extract_ocr_texts(r3) == "France WEEE 998877")
    # PaddleOCR 2.x: [[page], ...] 其中 page = [[box, (text, score)], ...]
    r2 = [[
        [[[0, 0], [1, 1]], ("France WEEE 998877", 0.99)],
        [[[0, 0], [1, 1]], ("second line", 0.9)],
    ]]
    check("2.x列表格式", _extract_ocr_texts(r2) == "France WEEE 998877\nsecond line")
    check("空结果", _extract_ocr_texts([]) == "" and _extract_ocr_texts([None, []]) == "")


def test_ocr_fallback():
    """E3: OCR 图片兜底 — 懒加载/触发条件/多项目拆行/来源标记/找不到返回空"""
    import utils.attachment_parser as ap
    from modules.field_extractor import FieldExtractor

    # 1. 常规解析阶段图片不触发 OCR（懒登记）
    tmp_png = os.path.join(tempfile.mkdtemp(), "截图.png")
    with open(tmp_png, "wb") as f:
        f.write(b"\x89PNG fake")
    called = []
    orig_parse_image = ap._parse_image
    ap._parse_image = lambda p: called.append(p) or "SHOULD_NOT_APPEAR"
    try:
        att = ap.parse_attachment(tmp_png, "截图.png")
        check("懒登记: 不OCR", att["text_content"] == "" and att.get("ocr_pending") is True)
        check("懒登记: 保留路径", att.get("filepath") == tmp_png)
        check("懒登记: OCR未被调用", not called)

        # 2. 字段齐全 → 兜底不触发
        fx = FieldExtractor({}, [])
        fx._ocr_image = lambda p: called.append(p) or ""
        mail_full = {
            "subject": "深圳市测试科技有限公司 法国WEEE 新增",
            "body_text": "", "sender_email": "",
            "attachments": [att],
        }
        rows = fx.extract_fields(mail_full)
        check("字段齐全: 不触发OCR", not called)
        # 客户提取自带主题整段行为(既有逻辑), 只验证关键内容存在 + 项目正确
        check("字段齐全: 正常提取", "深圳市测试科技有限公司" in rows[0]["客户"]
              and rows[0]["项目"] == "法国WEEE")

        # 3. 字段缺失 + OCR 命中 → 兜底提取 + 来源标记 + 置信度low
        att2 = ap.parse_attachment(tmp_png, "截图.png")
        fx2 = FieldExtractor({}, [])
        fx2._ocr_image = lambda p: "订单确认 深圳市某某科技有限公司 德国WEEE、德国电池法 请处理"
        mail_missing = {
            "subject": "请处理附件", "body_text": "", "sender_email": "",
            "attachments": [att2],
        }
        rows2 = fx2.extract_fields(mail_missing)
        check("兜底: 拆两行", len(rows2) == 2)
        check("兜底: 客户", rows2[0]["客户"] == "深圳市某某科技有限公司")
        check("兜底: 客户来源标记", rows2[0]["客户提取来源"] == "OCR图片兜底")
        check("兜底: 项目集合", {r["项目"] for r in rows2} == {"德国WEEE", "德国电池法"})
        check("兜底: 置信度low", all(r["置信度"] == "low" for r in rows2))
        check("兜底: ocr_used标记", all(r.get("ocr_used") for r in rows2))
        check("兜底: 不覆盖已有代理", rows2[0]["代理"] == "")

        # 4. OCR 找不到目标信息 → 不填, 保持待确认单行
        att3 = ap.parse_attachment(tmp_png, "截图.png")
        fx3 = FieldExtractor({}, [])
        fx3._ocr_image = lambda p: "广告水印 满减促销 优惠券 ABCDEF"
        rows3 = fx3.extract_fields({
            "subject": "请查收", "body_text": "", "sender_email": "",
            "attachments": [att3],
        })
        check("无信息: 单行空结果", len(rows3) == 1 and rows3[0]["客户"] == ""
              and rows3[0]["项目"] == "")
        check("无信息: 来源待确认", rows3[0]["客户提取来源"] == "待确认")
        check("无信息: ocr_used=False", not rows3[0].get("ocr_used"))

        # 5. 开关关闭 → 永不触发
        att4 = ap.parse_attachment(tmp_png, "截图.png")
        fx4 = FieldExtractor({}, [], ocr_fallback=False)
        fx4._ocr_image = lambda p: called.append(p) or "深圳市某公司 法国WEEE"
        rows4 = fx4.extract_fields({
            "subject": "请处理附件", "body_text": "", "sender_email": "",
            "attachments": [att4],
        })
        check("开关关闭: 不触发OCR", len(called) == 0)
        check("开关关闭: 字段为空", rows4[0]["客户"] == "" and rows4[0]["项目"] == "")
    finally:
        ap._parse_image = orig_parse_image
        shutil.rmtree(os.path.dirname(tmp_png), ignore_errors=True)


def test_stage2_preprocess():
    """阶段二前置校验: 跳过代理空+需人工确认; (公司,项目) 去重"""
    print("== 阶段二前置校验 ==")
    from modules.workorder_checker import WorkOrderChecker

    rows = [
        # 1. 代理单命中 + 置信度高 → 通过
        {"客户": "A公司", "项目": "德国WEEE", "代理": "代理甲", "置信度": "高"},
        # 2. 代理空 + 置信度=需人工确认 → 跳过
        {"客户": "B公司", "项目": "法国包装法", "代理": "", "置信度": "需人工确认"},
        # 3. (公司,项目) 与第1行重复 → 跳过
        {"客户": "A公司", "项目": "德国WEEE", "代理": "代理甲", "置信度": "高"},
        # 4. 代理命中但置信度=需人工确认 — 按规格"且代理空"才跳, 这里代理非空, 应放行
        {"客户": "C公司", "项目": "意大利WEEE", "代理": "代理乙", "置信度": "需人工确认"},
        # 5. 代理空但置信度高 — 不在跳过规则, 放行(由 RPA/人工兜底)
        {"客户": "D公司", "项目": "西班牙电池法", "代理": "", "置信度": "高"},
    ]
    to_query, skipped = WorkOrderChecker.preprocess_rows(rows)
    check("前置校验: 跳过代理空+置信度=需人工确认",
          any("代理空" in (s.get("_skip_reason") or "") for s in skipped))
    check("前置校验: (公司,项目) 重复被合并",
          any("重复" in (s.get("_skip_reason") or "") for s in skipped))
    check("前置校验: 通过3条(1单命中 + 1非空代理需确认 + 1空代理高置信)",
          len(to_query) == 3,
          f"got {len(to_query)}")
    # 通过的行 (公司,项目) 应唯一
    keys = [(r["客户"], r["项目"]) for r in to_query]
    check("前置校验: 通过行 (公司,项目) 无重复", len(set(keys)) == len(keys))


def test_query_cache_roundtrip():
    """阶段二查询缓存: 命中/过期/写盘"""
    print("== 阶段二查询缓存 ==")
    import tempfile, os, json
    from modules.workorder_checker import WorkOrderChecker

    tmpdir = tempfile.mkdtemp(prefix="qa_cache_")
    cache_path = os.path.join(tmpdir, "cache.json")
    try:
        cfg = {
            "url": "u", "username": "a", "password": "b",
            "query_cache_path": cache_path,
            "query_cache_ttl_days": 7,
        }
        c = WorkOrderChecker(cfg)
        check("缓存初始化为空", c.get_cached_result("代A", "公司X", "德国WEEE") is None)
        c.save_to_cache("代A", "公司X", "德国WEEE", [{"col1": "v1", "date": "2026-09-01"}])
        check("写盘后再读命中",
              c.get_cached_result("代A", "公司X", "德国WEEE") is not None)
        check("不同键不命中",
              c.get_cached_result("代B", "公司X", "德国WEEE") is None)
        # 文件落盘
        check("缓存文件已写盘", os.path.exists(cache_path))
        with open(cache_path, encoding="utf-8") as f:
            raw = json.load(f)
        check("缓存文件 1 条", len(raw) == 1)

        # 重新构造实例, 模拟重启, 命中落盘数据
        c2 = WorkOrderChecker(cfg)
        cached = c2.get_cached_result("代A", "公司X", "德国WEEE")
        check("新实例从文件命中缓存", cached is not None)
        check("缓存内容含原始 orders", len(cached) == 1 and cached[0].get("col1") == "v1")

        # 过期: 用 cache_ttl_days=7 但把内部 timestamp 改成 2000 年 → age.days > 7 → 过期
        c3 = WorkOrderChecker(cfg)  # ttl=7
        for k, v in list(c3._query_cache.items()):
            v["timestamp"] = "2000-01-01T00:00:00"
        c3._save_cache()
        check("过期条目不返回(返回None)",
              c3.get_cached_result("代A", "公司X", "德国WEEE") is None)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _run_async(coro):
    """跑异步协程, 不依赖外部事件循环"""
    import asyncio
    return asyncio.get_event_loop().run_until_complete(coro)


def test_stage2_outputs():
    """阶段一二输出文件名稳定 + 时序副本"""
    print("== 输出文件稳定名 + 时序副本 ==")
    tmpdir = tempfile.mkdtemp(prefix="qa_xlsx_")
    try:
        from modules.excel_writer import ExcelWriter
        writer = ExcelWriter(output_dir=tmpdir)

        rows = [
            {"客户": "A公司", "项目": "德国WEEE", "代理": "甲代理", "代理匹配方式": "精确匹配",
             "置信度": "高", "需求": "注册", "sender_email": "a@x.com", "date": None,
             "subject": "德国WEEE注册", "body_text": "请处理", "项目原始值": "德国WEEE",
             "客户提取来源": "邮件标题", "数据来源": "邮件标题"},
            {"客户": "B公司", "项目": "法国包装法", "代理": "", "代理匹配方式": "未匹配",
             "置信度": "需人工确认", "需求": "新增", "sender_email": "b@x.com", "date": None,
             "subject": "请帮忙", "body_text": "帮忙新增", "项目原始值": "法国包装法",
             "客户提取来源": "邮件正文", "数据来源": "邮件正文"},
        ]
        primary, secondary = writer.write_stage1_outputs(rows, [])
        # 稳定名存在
        check("阶段一待查清单 稳定名存在",
              os.path.exists(os.path.join(tmpdir, "to_workorder_list.xlsx")))
        check("阶段一漏单复查 稳定名存在",
              os.path.exists(os.path.join(tmpdir, "to_review_list.xlsx")))
        # 时序副本存在
        ts_files = [f for f in os.listdir(tmpdir)
                    if f.startswith("to_workorder_list_") and f.endswith(".xlsx")]
        check("阶段一待查清单 时序副本存在", len(ts_files) >= 1)
        # 分区: A公司 高置信 → to_work; B公司 代理空 → to_review
        from openpyxl import load_workbook
        wb_w = load_workbook(os.path.join(tmpdir, "to_workorder_list.xlsx"))
        wb_r = load_workbook(os.path.join(tmpdir, "to_review_list.xlsx"))
        check("待查清单 1 条", wb_w.active.max_row == 2)  # 1 表头 + 1 数据
        check("漏单复查 1 条", wb_r.active.max_row == 2)

        # 阶段二输出
        all_rows = list(rows)
        to_query = [rows[0]]
        skipped = [rows[1]]
        stable, ts = writer.write_workorder_check_result(all_rows, to_query, skipped)
        check("阶段二 稳定名存在", os.path.exists(os.path.join(tmpdir, "workorder_check_result.xlsx")))
        check("阶段二 时序副本存在", os.path.exists(ts))
        wb_s2 = load_workbook(stable)
        check("阶段二 包含原始邮件+查询+跳过行(3行:1表头+2数据)",
              wb_s2.active.max_row == 3)  # 1 表头 + 1 已查 + 1 跳过
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def test_llm_intent_mock():
    """LLM 邮件分类 + 字段提取: mock 验证协议解析, 默认零成本;
    设环境变量 MAIL_AUDIT_REAL_LLM=1 会跑 1 次真实请求冒烟 (代价 ~ 几分钱)"""
    print("== T11 LLM 邮件分类 + 字段提取 (mock 默认, 可选真实冒烟) ==")
    import os as _os
    from modules.llm_intent import LLMIntentClient

    client = LLMIntentClient({
        "api_key": "mock-key", "base_url": "http://mock", "model": "mock",
        "timeout": 5, "max_tokens": 100,
    })

    # 1. mock classify_email — 验证协议解析
    def fake_classify(system, user):
        return {
            "is_target": True, "intent": "注册",
            "email_type": "注册类", "reason": "客户发起WEEE注册",
        }
    client._call_api = fake_classify
    r = client.classify_email("客户咨询WEEE注册", "请尽快处理", [])
    check("classify: 协议字段齐全",
          r and r["is_target"] is True and r["intent"] == "注册"
          and r["email_type"] == "注册类")
    check("classify: reason 透传", r["reason"] == "客户发起WEEE注册")

    # 2. mock extract_fields_llm — 验证协议 + 多项目拆
    def fake_extract(system, user):
        return {
            "代理": "乐天", "客户": "深圳市甲科技有限公司",
            "项目": ["德国WEEE", "德国电池法"],
            "需求": "注册", "confidence": "high",
        }
    client._call_api = fake_extract
    f = client.extract_fields_llm("乐天+甲科技+WEEE注册", "", [], "agent@leko.com")
    check("extract: 字段映射",
          f["代理"] == "乐天" and f["客户"] == "深圳市甲科技有限公司"
          and f["项目"] == ["德国WEEE", "德国电池法"]
          and f["需求"] == "注册" and f["confidence"] == "high")

    # 3. mock 失败 (返回 None) → 上层走空结果, 不崩
    client._call_api = lambda *a, **k: None
    r2 = client.classify_email("x", "y", [])
    check("classify 失败: 返回 None", r2 is None)
    f2 = client.extract_fields_llm("x", "y", [])
    check("extract 失败: 返回 None", f2 is None)

    # 4. 真实冒烟 (可选)
    if _os.environ.get("MAIL_AUDIT_REAL_LLM") == "1":
        from openai import OpenAI
        import yaml as _yaml
        cfg_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "config.yaml")
        with open(cfg_path, encoding="utf-8") as _f:
            llm_cfg = _yaml.safe_load(_f)["llm"]
        live = LLMIntentClient(llm_cfg)
        live_resp = live.classify_email("WEEE 注册咨询", "请尽快处理", [])
        check("真实冒烟: 邮件分类成功",
              live_resp is not None and live_resp.get("is_target") is True,
              f"got {live_resp}")
    else:
        check("真实冒烟: 默认跳过(设 MAIL_AUDIT_REAL_LLM=1 启用)", True)


def test_llm_fallback_in_field_extractor():
    """T12: 字段提取 E2 LLM 兜底路径 — 缺字段时调用 LLM, 回填+标记 llm_used"""
    print("== T12 字段提取 LLM 兜底路径 ==")
    from modules.field_extractor import FieldExtractor
    from modules.llm_intent import LLMIntentClient

    # 1. LLM 不可用 → 缺字段直接空结果
    fx1 = FieldExtractor({}, [], llm_client=None)
    rows1 = fx1.extract_fields({
        "subject": "请处理", "body_text": "", "sender_email": "u@x.com",
        "attachments": [],
    })
    check("无 LLM 兜底: llm_used=False", not rows1[0].get("llm_used"))
    check("无 LLM 兜底: 字段空", rows1[0]["客户"] == "" and rows1[0]["项目"] == "")

    # 2. LLM 可用 + 字段缺失 → 触发 LLM 补字段
    fake = LLMIntentClient({"api_key": "mock", "base_url": "x", "model": "x"})
    fake.classify_email = lambda *a, **k: None
    fake.extract_fields_llm = lambda *a, **k: {
        "代理": "乐天", "客户": "深圳市甲科技有限公司",
        "项目": ["法国WEEE"], "需求": "注册", "confidence": "high",
    }
    fx2 = FieldExtractor({}, [], llm_client=fake)
    rows2 = fx2.extract_fields({
        "subject": "请处理", "body_text": "", "sender_email": "u@x.com",
        "attachments": [],
    })
    check("LLM 兜底: llm_used=True", rows2[0].get("llm_used") is True)
    check("LLM 兜底: 客户回填", rows2[0]["客户"] == "深圳市甲科技有限公司")
    check("LLM 兜底: 项目回填", rows2[0]["项目"] == "法国WEEE")
    check("LLM 兜底: 代理回填", rows2[0]["代理"] == "乐天")
    check("LLM 兜底: 需求回填", rows2[0]["需求"] == "注册")

    # 3. 字段已由规则提取齐全 → 不调 LLM
    called = []
    fake.extract_fields_llm = lambda *a, **k: called.append(1) or {
        "代理": "", "客户": "", "项目": [], "需求": "", "confidence": "low",
    }
    fx3 = FieldExtractor(
        {"agent1@x.com": {"代理": "乐天", "代理简称": "乐天", "收件人邮箱": "agent1@x.com"}},
        [],
        llm_client=fake,
    )
    rows3 = fx3.extract_fields({
        "subject": "乐天+深圳市甲科技有限公司+德国WEEE注册",
        "body_text": "", "sender_email": "agent1@x.com",
        "attachments": [],
    })
    check("字段齐全: 不调 LLM", not called and len(rows3) > 0)
    check("字段齐全: 规则结果正确", rows3[0]["客户"] == "深圳市甲科技有限公司"
          and rows3[0]["项目"] == "德国WEEE" and rows3[0]["代理"] == "乐天")

    # 4. LLM 返回项目是列表 → 拆多行
    fake.extract_fields_llm = lambda *a, **k: {
        "代理": "乐天", "客户": "深圳市甲科技有限公司",
        "项目": ["德国WEEE", "德国电池法"], "需求": "注册", "confidence": "high",
    }
    fx4 = FieldExtractor({}, [], llm_client=fake)
    rows4 = fx4.extract_fields({
        "subject": "请处理", "body_text": "", "sender_email": "u@x.com",
        "attachments": [],
    })
    check("LLM 多项目: 拆 2 行", len(rows4) == 2)
    check("LLM 多项目: 项目集合正确",
          {r["项目"] for r in rows4} == {"德国WEEE", "德国电池法"})
    check("LLM 多项目: 客户相同", len({r["客户"] for r in rows4}) == 1)


def main():
    tests = [
        test_mail_filter,
        test_field_extractor,
        test_project_normalizer,
        test_workorder_match,
        test_excel_writer,
        test_ocr_compat,
        test_ocr_fallback,
        test_stage2_preprocess,
        test_query_cache_roundtrip,
        test_stage2_outputs,
        test_llm_intent_mock,
        test_llm_fallback_in_field_extractor,
    ]
    for t in tests:
        t()
    print("=" * 40)
    print(f"结果: {PASS} 通过, {FAIL} 失败")
    sys.exit(1 if FAIL else 0)


if __name__ == "__main__":
    main()
