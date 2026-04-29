"""
知虾/猛拉商品数据抓取脚本 v2
网站: https://shopee.menglar.com/workbench/home
流程: 首页输入链接 -> 跳转新标签页 -> 抓取三类数据 -> 关闭标签页 -> 循环

输出 Excel 结构:
  Sheet "商品基础信息"         - 每个产品一行
  Sheet "销售数据概览"         - 每个产品一行
  Sheet "销售数据明细_产品ID"  - 每个产品单独一个 Sheet（近30天）
  downloads_时间戳/            - 导出按钮下载的备用文件
"""

from playwright.sync_api import sync_playwright
import pandas as pd
import re
import time
import os
from datetime import datetime

# ==================== 配置区 ====================
CHROME_PATH = "C:/Program Files/Google/Chrome/Application/chrome.exe"
ZHIXIA_HOME = "https://shopee.menglar.com/workbench/home"

# 方式1：直接写链接
PRODUCT_URLS = [
    # "https://shopee.ph/xxx",
]

# 方式2：从 txt 读（优先级更高）
URL_FILE = "urls.txt"


def read_urls():
    if os.path.exists(URL_FILE):
        with open(URL_FILE, "r", encoding="utf-8") as f:
            urls = [line.strip() for line in f if line.strip() and not line.startswith("#")]
        print(f"[INFO] 从 {URL_FILE} 读取了 {len(urls)} 个链接")
        return urls
    urls = [u for u in PRODUCT_URLS if u.strip()]
    print(f"[INFO] 从脚本读取了 {len(urls)} 个链接")
    return urls


def search_and_get_new_tab(context, home_page, url):
    """在首页搜索，并等待/返回新打开的标签页对象"""
    search_input = home_page.locator("input[placeholder='请输入商品ID或者商品链接']")

    search_input.click()
    time.sleep(0.5)
    search_input.fill("")
    search_input.fill(url)
    time.sleep(1)
    home_page.keyboard.press("Escape")
    time.sleep(0.5)

    print("  -> 点击搜索，等待新标签页打开...")
    home_page.locator("button.zx-search-button").click()

    try:
        new_page = context.wait_for_event("page", timeout=15000)
        new_page.wait_for_selector(".metric-title", timeout=20000)
        print("  -> ✅ 新标签页已打开且数据加载完成")
        return new_page
    except Exception as e:
        print(f"  -> ❌ 等待新标签页失败: {e}")
        return None


def scrape_basic_info(page):
    """
    抓取商品基础信息（页面顶部区域）。
    策略：用 JS 逐字段查找——找到包含该标签文字的 span/div，
    取其紧跟的兄弟元素或父容器中标签后的值。
    """
    info = {
        "商品名称": "", "产品ID": "", "品牌": "", "店铺类型": "",
        "店铺名称": "", "店铺ID": "", "上架时间": "", "类目路径": ""
    }

    try:
        result = page.evaluate("""
            () => {
                // 工具函数：找包含 label 文字的最小叶节点，取其后兄弟或父级中的值
                function extractAfterLabel(label) {
                    const walker = document.createTreeWalker(
                        document.body, NodeFilter.SHOW_TEXT, null, false
                    );
                    let node;
                    while (node = walker.nextNode()) {
                        const txt = node.textContent.trim();
                        if (txt === label || txt === label + '：' || txt === label + ':') {
                            // 找父元素的下一个兄弟
                            let parent = node.parentElement;
                            // 尝试 nextSibling（文字节点或元素）
                            let sib = parent.nextSibling;
                            while (sib) {
                                const v = sib.textContent ? sib.textContent.trim() : '';
                                if (v && v !== label) return v;
                                sib = sib.nextSibling;
                            }
                            // 尝试父元素的下一个兄弟元素
                            let parentSib = parent.nextElementSibling;
                            if (parentSib) {
                                const v = parentSib.textContent.trim();
                                if (v && v !== label) return v;
                            }
                            // 尝试祖父元素的下一个兄弟元素
                            let gp = parent.parentElement;
                            if (gp) {
                                let gpSib = gp.nextElementSibling;
                                if (gpSib) {
                                    const v = gpSib.textContent.trim();
                                    if (v && v !== label) return v.split('\\n')[0].trim();
                                }
                            }
                        }
                    }
                    return '';
                }

                // 商品名称：找页面里最长的、不含标签关键词的标题文字
                function extractTitle() {
                    // 先找带商品链接图标的容器（知虾商品名旁边有跳转图标）
                    const titleCandidates = document.querySelectorAll(
                        '.goods-name, .product-name, .item-name, ' +
                        '.goods-title, .product-title, .detail-title, ' +
                        'h1, h2'
                    );
                    for (const el of titleCandidates) {
                        const t = el.innerText ? el.innerText.trim() : '';
                        if (t.length > 10 && !t.includes('产品ID') && !t.includes('店铺')) {
                            return t.split('\\n')[0].trim();
                        }
                    }
                    // 备用：找 .metric 区域上方足够长的文字块
                    const metricArea = document.querySelector('.metric-title');
                    if (metricArea) {
                        let el = metricArea.parentElement;
                        for (let i = 0; i < 6; i++) {
                            el = el ? el.parentElement : null;
                        }
                        if (el) {
                            const lines = (el.innerText || '').split('\\n').map(l => l.trim()).filter(Boolean);
                            for (const line of lines) {
                                if (line.length > 20 && !line.includes('产品ID') && 
                                    !line.includes('店铺') && !line.includes('上架')) {
                                    return line;
                                }
                            }
                        }
                    }
                    return '';
                }

                // 类目路径：找包含 > 和 Sports 的行
                function extractCategory() {
                    const walker = document.createTreeWalker(
                        document.body, NodeFilter.SHOW_TEXT, null, false
                    );
                    let node;
                    while (node = walker.nextNode()) {
                        const txt = node.textContent.trim();
                        if (txt.includes('>') && txt.length > 20 && 
                            (txt.includes('Sports') || txt.includes('运动') || txt.includes('Fashion'))) {
                            return txt;
                        }
                    }
                    return '';
                }

                return {
                    title:       extractTitle(),
                    productId:   extractAfterLabel('产品ID'),
                    brand:       extractAfterLabel('品牌'),
                    shopType:    extractAfterLabel('店铺类型'),
                    shopName:    extractAfterLabel('店铺名称'),
                    shopId:      extractAfterLabel('店铺ID'),
                    listingTime: extractAfterLabel('上架时间'),
                    category:    extractCategory(),
                };
            }
        """)

        info["商品名称"] = result.get("title", "").strip()
        info["产品ID"]   = result.get("productId", "").strip()
        info["品牌"]     = result.get("brand", "").strip()
        info["店铺类型"] = result.get("shopType", "").strip()
        info["店铺名称"] = result.get("shopName", "").strip()
        info["店铺ID"]   = result.get("shopId", "").strip()
        info["上架时间"] = result.get("listingTime", "").strip()
        info["类目路径"] = result.get("category", "").strip()

        # 清理品牌：如果抓到导航菜单的词（如"分析"），则置空
        nav_noise = {"分析", "首页", "管理", "导航", "监控", "选品", "工具"}
        if info["品牌"] in nav_noise:
            info["品牌"] = ""

        filled = sum(1 for v in info.values() if v)
        print(f"  -> ✅ 商品基础信息: {filled}/8 个字段有值")

    except Exception as e:
        print(f"  -> ⚠️ 商品基础信息抓取出错: {e}")

    return info



def scrape_overview(page):
    """
    抓取核心指标 + 销售指标（销售数据概览）。
    知虾结构: .metric-title + .metric-value 在同一父容器里。
    """
    overview = {}

    field_map = {
        "总销量": "总销量",
        "产品排名": "产品排名",
        "累计评论数": "累计评论数",
        "累计点赞数": "累计点赞数",
        "折扣价": "折扣价",
        "产品评分": "产品评分",
        "近一日销量": "近一日销量",
        "近7天销量": "近7天销量",
        "近30天销量": "近30天销量",
        "近一日销售额": "近一日销售额",
        "近7天销售额": "近7天销售额",
        "近30天销售额": "近30天销售额",
    }

    try:
        titles = page.locator(".metric-title")
        count = titles.count()

        for i in range(count):
            raw = titles.nth(i).inner_text().strip()
            title_text = raw.split("\n")[0].strip()

            matched_key = None
            for page_field, col_name in field_map.items():
                if page_field in title_text:
                    matched_key = col_name
                    break

            if not matched_key:
                continue

            val = ""
            # 方法1：找 following-sibling 里含 value/number 的元素
            try:
                value_el = titles.nth(i).locator(
                    "xpath=following-sibling::*[contains(@class,'metric-value') "
                    "or contains(@class,'value') or contains(@class,'number')]"
                ).first
                if value_el.count() > 0:
                    val = value_el.inner_text().strip().split("\n")[0].strip()
            except:
                pass

            # 方法2：取父容器文字，去掉 title 行后取第一个有值的行
            if not val:
                try:
                    parent_text = titles.nth(i).locator("xpath=..").inner_text().strip()
                    for line in parent_text.split("\n"):
                        line = line.strip()
                        if line and line != title_text and "环比" not in line:
                            val = line
                            break
                except:
                    pass

            if val:
                overview[matched_key] = val

        print(f"  -> ✅ 销售数据概览: 抓取到 {len(overview)} 个字段")

    except Exception as e:
        print(f"  -> ⚠️ 销售数据概览抓取出错: {e}")

    return overview


def scrape_detail_table(page):
    """
    抓取销售数据明细表格（近30天）。
    知虾表格是 el-table，部分 td 含多行数据（用换行分隔）：
      - 日折扣价：₱价格 换行 ¥价格
      - 日销量列：日销量 换行 周销量 换行 月销量
      - 日销售额列：日销售额 换行 周销售额 换行 月销售额
      - 点赞/评论列：累计点赞数 换行 累计评论数
    """
    try:
        # 先滚动页面让明细区域可见，再操作按钮
        page.evaluate("window.scrollTo(0, document.body.scrollHeight * 0.5)")
        time.sleep(1.5)

        # 确保近30天激活
        try:
            btn_30 = page.locator("button:has-text('近30天')").first
            if btn_30.count() > 0 and btn_30.is_visible():
                btn_30.click()
                time.sleep(2)
        except:
            pass

        # 滚动到销售数据明细 tab
        try:
            # 知虾的 tab 栏通常在图表上方，文字是"销售数据明细"
            tab = page.locator(".el-tabs__item, .tab-item, [role='tab']").filter(has_text="销售数据明细").first
            if tab.count() > 0 and tab.is_visible():
                tab.scroll_into_view_if_needed()
                tab.click()
                time.sleep(1.5)
        except:
            pass

        # 继续向下滚动确保表格渲染
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(1.5)

        # 找表格：取所有 el-table 里包含"日期"列的那个
        # 用 JS 拿到明细表格的所有行数据，避免 Playwright locator 超时
        raw_data = page.evaluate("""
            () => {
                // 找到包含"日期"列的表格（排除只有汇总数据的小表）
                const tables = document.querySelectorAll('.el-table');
                let targetTable = null;
                for (const t of tables) {
                    const headers = t.querySelectorAll('.el-table__header th .cell');
                    for (const h of headers) {
                        if (h.textContent.trim() === '日期') {
                            // 确认这是明细表（有多行数据）
                            const bodyRows = t.querySelectorAll('.el-table__body tbody tr');
                            if (bodyRows.length > 2) {
                                targetTable = t;
                                break;
                            }
                        }
                    }
                    if (targetTable) break;
                }

                if (!targetTable) return null;

                // 提取所有 tbody tr
                const rows = targetTable.querySelectorAll('.el-table__body tbody tr');
                const result = [];
                for (const row of rows) {
                    const cells = row.querySelectorAll('td');
                    const cellData = [];
                    for (const cell of cells) {
                        cellData.push(cell.innerText.trim());
                    }
                    if (cellData.some(c => c)) {
                        result.push(cellData);
                    }
                }
                return result;
            }
        """)

        if not raw_data:
            print("  -> ⚠️ 未找到销售数据明细表格（JS方式）")
            return [], []

        parsed_rows = []
        for cell_texts in raw_data:
            parsed = _parse_detail_row(cell_texts)
            if parsed:
                parsed_rows.append(parsed)

        headers = [
            "日期",
            "日折扣价₱", "日折扣价¥",
            "日销量", "周销量", "月销量",
            "日销售额₱", "周销售额₱", "月销售额₱",
            "累计点赞数", "累计评论数",
        ]

        print(f"  -> ✅ 销售数据明细: 抓取到 {len(parsed_rows)} 行数据")
        return headers, parsed_rows

    except Exception as e:
        print(f"  -> ⚠️ 销售数据明细抓取出错: {e}")
        return [], []


def _parse_detail_row(cell_texts):
    """
    把一行 td 的 innerText 列表解析成固定11列。
    过滤掉"环比"行，按换行拆分多值单元格。
    """
    def split_clean(text, n):
        """按换行拆分，过滤环比行，补足 n 个值"""
        parts = [p.strip() for p in text.split("\n") if p.strip() and "环比" not in p]
        while len(parts) < n:
            parts.append("")
        return parts[:n]

    if len(cell_texts) < 3:
        return None

    try:
        date_val   = cell_texts[0].strip()
        price_p    = split_clean(cell_texts[1] if len(cell_texts) > 1 else "", 2)
        vol_p      = split_clean(cell_texts[2] if len(cell_texts) > 2 else "", 3)
        sales_p    = split_clean(cell_texts[3] if len(cell_texts) > 3 else "", 3)
        like_p     = split_clean(cell_texts[4] if len(cell_texts) > 4 else "", 2)

        return [
            date_val,
            price_p[0], price_p[1],
            vol_p[0], vol_p[1], vol_p[2],
            sales_p[0], sales_p[1], sales_p[2],
            like_p[0], like_p[1],
        ]
    except:
        return None


def click_export(page, download_dir, product_id=""):
    """点击导出按钮并保存（备用）。只操作可见的按钮，避免超时。"""
    try:
        # 滚动到页面底部确保导出按钮可见
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(1)

        # 找所有"导出"按钮，只取可见的那个
        export_btns = page.locator("button:has-text('导出'), .el-button:has-text('导出')").all()
        export_btn = None
        for btn in reversed(export_btns):  # 从后往前，明细区的导出按钮通常在下方
            try:
                if btn.is_visible(timeout=1000):
                    export_btn = btn
                    break
            except:
                continue

        if not export_btn:
            print("  -> ℹ️ 未找到可见的导出按钮，跳过")
            return None

        export_btn.scroll_into_view_if_needed()
        time.sleep(0.5)

        with page.expect_download(timeout=30000) as download_info:
            export_btn.click()

        download = download_info.value
        suggested = download.suggested_filename or f"detail_{product_id}.xlsx"
        save_path = os.path.join(download_dir, f"备用导出_{product_id}_{suggested}")
        download.save_as(save_path)
        print(f"  -> ✅ 备用导出文件已保存")
        return save_path
    except Exception as e:
        print(f"  -> ℹ️ 导出跳过: {e}")
        return None


def extract_product_id_from_url(url):
    """从 Shopee URL 提取产品ID，格式通常是 i.店铺ID.产品ID"""
    match = re.search(r'i\.(\d+)\.(\d+)', url)
    if match:
        return match.group(2)
    parts = url.rstrip("/").split("/")
    for part in reversed(parts):
        clean = re.sub(r'[^0-9]', '', part.split(".")[-1])
        if clean and len(clean) > 5:
            return clean
    return str(int(time.time()))


def run():
    if not os.path.exists("auth.json"):
        print("❌ 找不到 auth.json！请先运行 save_auth.py 保存登录态。")
        return

    urls = read_urls()
    if not urls:
        print("❌ 没有要抓取的链接！请修改脚本或创建 urls.txt")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    download_dir = f"./downloads_{timestamp}"
    os.makedirs(download_dir, exist_ok=True)

    all_basic    = []
    all_overview = []
    all_details  = {}  # {product_id: (headers, rows)}

    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=CHROME_PATH, headless=False)
        context = browser.new_context(
            storage_state="auth.json",
            viewport={"width": 1920, "height": 1080},
            accept_downloads=True,
        )

        home_page = context.new_page()
        home_page.goto(ZHIXIA_HOME, wait_until="domcontentloaded", timeout=30000)
        time.sleep(3)
        print("✅ 首页已打开\n")

        for idx, url in enumerate(urls):
            print(f"\n{'='*60}")
            print(f"[{idx+1}/{len(urls)}] 处理: {url}")
            print(f"{'='*60}")

            url_product_id = extract_product_id_from_url(url)

            detail_page = search_and_get_new_tab(context, home_page, url)
            if not detail_page:
                print("  -> 跳过此链接")
                continue

            time.sleep(2)

            # 1. 商品基础信息
            print("  -> 抓取商品基础信息...")
            basic = scrape_basic_info(detail_page)
            basic["来源链接"] = url
            basic["序号"] = idx + 1
            product_id = basic.get("产品ID", "").strip() or url_product_id
            basic["产品ID"] = product_id
            all_basic.append(basic)
            print(f"  -> 产品ID: {product_id}")

            # 2. 销售数据概览
            print("  -> 抓取销售数据概览...")
            overview = scrape_overview(detail_page)
            overview["产品ID"] = product_id
            overview["来源链接"] = url
            overview["序号"] = idx + 1
            all_overview.append(overview)

            # 3. 销售数据明细
            print("  -> 抓取销售数据明细（近30天）...")
            headers, rows = scrape_detail_table(detail_page)
            if rows:
                all_details[product_id] = (headers, rows)

            # 4. 备用导出
            print("  -> 尝试导出备用文件...")
            click_export(detail_page, download_dir, product_id=product_id)

            # 5. 关闭标签页
            print("  -> 🚪 关闭详情页，返回首页...")
            detail_page.close()
            time.sleep(1)

        browser.close()

    # ==================== 写入 Excel ====================
    print(f"\n{'='*60}")
    print("🚀 正在写入最终 Excel...")

    final_excel = f"知虾抓取结果_{timestamp}.xlsx"

    basic_cols = ["序号", "来源链接", "商品名称", "产品ID", "品牌",
                  "店铺类型", "店铺名称", "店铺ID", "上架时间", "类目路径"]

    overview_cols = ["序号", "产品ID", "来源链接",
                     "总销量", "产品排名", "累计评论数", "累计点赞数",
                     "折扣价", "产品评分",
                     "近一日销量", "近7天销量", "近30天销量",
                     "近一日销售额", "近7天销售额", "近30天销售额"]

    with pd.ExcelWriter(final_excel, engine="openpyxl") as writer:

        if all_basic:
            df_basic = pd.DataFrame(all_basic)
            ordered = [c for c in basic_cols if c in df_basic.columns]
            extra   = [c for c in df_basic.columns if c not in basic_cols]
            df_basic[ordered + extra].to_excel(writer, sheet_name="商品基础信息", index=False)
            print(f"  -> ✅ Sheet '商品基础信息': {len(df_basic)} 行")

        if all_overview:
            df_overview = pd.DataFrame(all_overview)
            ordered = [c for c in overview_cols if c in df_overview.columns]
            extra   = [c for c in df_overview.columns if c not in overview_cols]
            df_overview[ordered + extra].to_excel(writer, sheet_name="销售数据概览", index=False)
            print(f"  -> ✅ Sheet '销售数据概览': {len(df_overview)} 行")

        for pid, (hdrs, rws) in all_details.items():
            sheet_name = f"销售数据明细_{pid}"[:31]
            pd.DataFrame(rws, columns=hdrs).to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  -> ✅ Sheet '{sheet_name}': {len(rws)} 行")

    print(f"\n🎉 全部完成！最终文件: {final_excel}")
    if os.path.exists(download_dir) and os.listdir(download_dir):
        print(f"📁 备用导出文件: {download_dir}/")


if __name__ == "__main__":
    run()
