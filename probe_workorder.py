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
    const r = el.getBoundingClientRect();
    const s = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
  };

  const pick = el => ({
    tag: el.tagName.toLowerCase(),
    text: txt(el),
    placeholder: el.getAttribute('placeholder') || '',
    name: el.getAttribute('name') || '',
    id: el.id || '',
    cls: (el.className && typeof el.className === 'string') ? el.className.slice(0, 120) : '',
  });

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
    .filter(el => vis(el) && el.tagName.toLowerCase() !== 'option')
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
    return { headers: ths, row_count: t.querySelectorAll('tbody tr').length, cls: t.className };
  }).filter(t => t.headers.length && vis(t.parentElement || t));

  return {
    url: location.href,
    title: document.title,
    inputs, selects, buttons, navs, labels, tables,
    body_head: document.body.innerText.slice(0, 1200),
  };
}
"""


async def main():
    config_path = os.path.join(ROOT, "config.yaml")
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    from playwright.async_api import async_playwright
    from modules.workorder_checker import WorkOrderChecker
    import logging
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    logger = logging.getLogger("mail_audit")

    out_dir = config.get("output", {}).get("dir", "output")
    os.makedirs(out_dir, exist_ok=True)

    checker = WorkOrderChecker(config["workorder"], logger)

    if not await checker.login():
        print("登录失败, 无法继续探测")
        await checker.close()
        return 1

    page = checker._page

    report = {"probed_at": datetime.now().isoformat(), "pages": []}

    async def dump(tag):
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

    # 1. 登录后页面
    await dump("after_login")

    # 2. 尝试逐个展开左侧菜单, 找到工单列表页
    menu_texts = []
    try:
        items = await page.query_selector_all(
            '.el-menu-item, .ant-menu-item, [class*="menu-item"], aside a, .sidebar a'
        )
        for it in items:
            t = (await it.inner_text() or "").strip().replace("\n", " ")
            if t and t not in menu_texts:
                menu_texts.append(t)
    except Exception:
        pass
    report["menu_texts"] = menu_texts
    logger.info(f"左侧菜单项: {menu_texts}")

    # 依次点击含"工单"的菜单
    for kw in ["工单管理", "工单", "注册工单", "服务管理"]:
        for mt in menu_texts:
            if kw in mt:
                try:
                    el = await page.query_selector(f'text="{mt}"')
                    if el:
                        await el.click()
                        await asyncio.sleep(2)
                        await page.wait_for_load_state("networkidle")
                        await asyncio.sleep(1)
                        await dump(f"click_{mt}".replace(" ", "_")[:40])
                except Exception as e:
                    logger.warning(f"点击菜单 {mt} 失败: {e}")
                break

    # 3. 保存汇总
    out_json = os.path.join(out_dir, "workorder_probe.json")
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    print("\n" + "=" * 60)
    print(f"探测完成, 结果已写入: {out_json}")
    print("请打开该文件, 按里面的 inputs/selects/tables 把选择器填到 config.yaml 的 selectors 段")
    print("=" * 60)

    await checker.close()
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
