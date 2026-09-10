"""工单系统页面结构探测脚本

用途: Playwright 的选择器是基于通用框架猜测的, 实际页面结构未确认。
本脚本半自动登录工单系统后, 把"工单管理"页面上的真实结构 dump 出来:
  - 左侧/顶部导航项文字
  - 所有输入框的 placeholder / name / id
  - 所有下拉组件的 label 与当前值
  - 所有按钮文字
  - 表格表头

产出: output/workorder_probe.json + 若干截图
拿到结果后, 把真实选择器填进 config.yaml 的 selectors 段即可, 不用改代码。

用法:
    python probe_workorder.py
"""
import os
import sys
import json
import asyncio
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import yaml


DUMP_JS = r"""
() => {
  const txt = el => (el.innerText || el.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 60);
  const vis = el => {
    try {
      if (typeof el.getBoundingClientRect !== 'function') return false;
      const r = el.getBoundingClientRect();
      if (!r || r.width <= 0 || r.height <= 0) return false;
      const s = getComputedStyle(el);
      return s.display !== 'none' && s.visibility !== 'hidden';
    } catch (e) { return false; }
  };

  const pick = el => {
    try {
      return {
        tag: el.tagName ? el.tagName.toLowerCase() : '',
        text: txt(el),
        placeholder: el.getAttribute ? (el.getAttribute('placeholder') || '') : '',
        name: el.getAttribute ? (el.getAttribute('name') || '') : '',
        id: el.id || '',
        cls: (el.className && typeof el.className === 'string') ? el.className.slice(0, 120) : '',
      };
    } catch (e) { return { tag: 'unknown' }; }
  };

  // 输入框
  const inputs = Array.from(document.querySelectorAll('input, textarea'))
    .filter(vis).map(el => ({
      ...pick(el),
      type: el.getAttribute('type') || '',
      readonly: el.hasAttribute('readonly'),
      value: (el.value || '').slice(0, 40),
    }));

  // 下拉组件 (ElementUI / AntD / 原生 select)
  const selects = Array.from(document.querySelectorAll('select, .el-select, .ant-select, [class*="select"]'))
    .filter(el => {
      try { return vis(el) && el.tagName && el.tagName.toLowerCase() !== 'option'; }
      catch (e) { return false; }
    })
    .map(el => ({ ...pick(el), role: el.getAttribute('role') || '' }));

  // 按钮
  const buttons = Array.from(document.querySelectorAll('button, [role="button"], .el-button, .ant-btn, a.btn'))
    .filter(vis).map(pick);

  // 导航
  const navs = Array.from(document.querySelectorAll('a, li, .el-menu-item, .ant-menu-item, [class*="menu-item"]'))
    .filter(vis).map(pick).filter(x => x.text);

  // 表单 label
  const labels = Array.from(document.querySelectorAll('label, .el-form-item__label, .ant-form-item-label'))
    .filter(vis).map(pick).filter(x => x.text);

  // 表格
  const tables = Array.from(document.querySelectorAll('table')).map(t => {
    const ths = Array.from(t.querySelectorAll('thead th, tr:first-child th, tr:first-child td'))
      .map(th => txt(th)).filter(Boolean);
    let parentVis = false;
    try { parentVis = vis(t.parentElement || t); } catch (e) {}
    return { headers: ths, row_count: t.querySelectorAll('tbody tr').length, cls: t.className };
  }).filter(t => t.headers.length && parentVis);

  return {
    url: location.href,
    title: document.title,
    inputs, selects, buttons, navs, labels, tables,
    body_head: (document.body && document.body.innerText || '').slice(0, 1200),
  };
}
"""


async def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default="https://mp.ecopv-epr.com/work/workOrderRegister",
                        help="工单系统页面 URL (默认: 注册工单页)")
    args = parser.parse_args()

    config_path = os.path.join(ROOT, "config.yaml")
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    from playwright.async_api import async_playwright
    import logging
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    logger = logging.getLogger("mail_audit")

    out_dir = config.get("output", {}).get("dir", "output")
    os.makedirs(out_dir, exist_ok=True)

    report = {"probed_at": datetime.now().isoformat(), "target_url": args.url, "pages": []}

    async def dump(page, tag):
        try:
            data = await page.evaluate(DUMP_JS)
        except Exception as e:
            logger.warning(f"{tag} 结构 dump 失败: {e}")
            return
        data["tag"] = tag
        report["pages"].append(data)
        shot = os.path.join(out_dir, f"probe_{tag}.png")
        try:
            await page.screenshot(path=shot, full_page=False)
            data["screenshot"] = shot
        except Exception:
            pass
        logger.info(f"[{tag}] 输入框 {len(data['inputs'])} / 下拉 {len(data['selects'])} "
                    f"/ 按钮 {len(data['buttons'])} / 表格 {len(data['tables'])}")

    async with async_playwright() as p:
        # playwright 内置 chromium (Defender 白名单已加, 应已生效)
        browser = await p.chromium.launch(
            headless=False,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
            ],
        )
        ctx = await browser.new_context(viewport={"width": 1440, "height": 900})
        page = await ctx.new_page()

        # 第一步: 打开登录页, 让用户输验证码
        login_url = "https://mp.ecopv-epr.com"
        logger.info(f"goto {login_url}")
        await page.goto(login_url, wait_until="domcontentloaded", timeout=20000)
        await asyncio.sleep(3)
        await dump(page, "login_page")

        print("\n" + "=" * 60)
        print(">>> 请在弹出的浏览器里输入验证码 + 点登录 <<<")
        print("登录后我会自动跳转到", args.url)
        print("=" * 60)

        # 等登录完成: 检测 url 不再是登录页 (最多等 3 分钟)
        for _ in range(36):  # 36 * 5s = 180s
            await asyncio.sleep(5)
            cur = page.url
            if "/login" not in cur.lower() and cur != login_url and cur != login_url + "/":
                logger.info(f"检测到跳转: {cur}")
                break
        else:
            logger.warning("3 分钟内未检测到登录跳转, 仍按当前页面 dump")

        # 第二步: 跳到目标页
        try:
            logger.info(f"goto {args.url}")
            await page.goto(args.url, wait_until="domcontentloaded", timeout=20000)
            await asyncio.sleep(3)
            await dump(page, f"goto_{args.url.split('//')[-1].replace('/', '_')[:50]}")
        except Exception as e:
            logger.warning(f"goto {args.url} 失败: {e}")

        await asyncio.sleep(2)
        await dump(page, "final_dom")

        out_json = os.path.join(out_dir, "workorder_probe.json")
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)

        print("\n" + "=" * 60)
        print(f"探测完成, 结果已写入: {out_json}")
        print("=" * 60)

        await ctx.close()
        await browser.close()
        return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
