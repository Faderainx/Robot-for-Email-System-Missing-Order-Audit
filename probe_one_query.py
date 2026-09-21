"""单条工单查询验证脚本

只查 1 条, 1~2 分钟出结果, 用来确认「结果表解析」是否已经修好,
不用等整批 20 分钟。

用法:
    python probe_one_query.py "公司名" ["代理"] ["国家"] ["服务项目"] [--headless]

不带参数时用内置样例。默认弹出浏览器（探针是给人看过程的，方便现场核对）；
加 --headless 则跟随后台无头方式运行，适合只想看日志的场景。
"""
import asyncio
import json
import os
import sys

import yaml

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from modules.workorder_checker import WorkOrderChecker


class ConsoleLogger:
    def info(self, m):    print(m)
    def warning(self, m): print("[WARNING]", m)
    def error(self, m):   print("[ERROR]", m)
    def debug(self, m):   pass


def main():
    cfg = yaml.safe_load(open(os.path.join(ROOT, "config.yaml"), encoding="utf-8"))
    wo = dict(cfg["workorder"])
    wo["force_live_query"] = True          # 忽略缓存, 真实点网页查询
    wo["max_consecutive_query_failures"] = 1

    argv = sys.argv[1:]
    # 探针默认可视化（调试时能直接看页面填得对不对）；--headless 才后台跑。
    if "--headless" in argv:
        argv.remove("--headless")
        wo["headless"] = True
    else:
        wo["headless"] = False

    company = argv[0] if len(argv) > 0 else "厦门汉印股份有限公司"
    agent   = argv[1] if len(argv) > 1 else ""
    country = argv[2] if len(argv) > 2 else ""
    item    = argv[3] if len(argv) > 3 else ""

    print(f"查询条件: 公司={company} 代理={agent} 国家={country} 服务项目={item}")
    print("-" * 60)

    checker = WorkOrderChecker(wo, ConsoleLogger())
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    async def run():
        ok = await checker.login()
        print("-" * 60)
        print("登录:", "成功" if ok else "失败")
        if not ok:
            print("(登录失败, 后面的查询没跑; 如需重试请重新运行本脚本)")
            return
        recs = await checker.search_one(
            company, agent=agent, country=country, service_item=item
        )
        print("-" * 60)
        print("=== 结论 ===")
        print("结果表解析状态:", getattr(checker, "_last_table_parse_status", "?") or "(未设置)")
        print("结果表解析说明:", getattr(checker, "_last_table_parse_note", "") or "(无)")
        print("RPA 查询状态  :", checker._last_query_status or "(无)")
        print("读到工单记录数:", len(recs))
        for i, r in enumerate(recs, 1):
            print(f"  [{i}] {json.dumps(r, ensure_ascii=False)}")
        if recs:
            print()
            print("下游取值自检:")
            for key in ("客户", "所属代理", "项目", "日期", "工单编号", "工单状态"):
                print(f"  {key} = {r.get(key, '')!r}")

    try:
        loop.run_until_complete(run())
    except Exception as exc:                     # noqa: BLE001
        print("[ERROR] 运行异常:", repr(exc))
    finally:
        try:
            loop.run_until_complete(checker.close())
        except Exception:
            pass


if __name__ == "__main__":
    main()
