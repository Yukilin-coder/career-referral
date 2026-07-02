---
name: career-referral
description: >
  内推助手。核心功能：根据用户提供的范围（学校、目标公司等），爬取LinkedIn人脉信息，
  输出Excel表格（人名、关系、对方履历）；并根据用户背景和目标岗位，输出通用求内推话术模板。
  属于 career-transition 转行助手套件的可选组件，在 Step 5（job-hunt）之后或并行触发。
  当用户在套件流程中说"帮我找内推""爬LinkedIn""找校友""找前同事""内推话术"时触发。
  独立使用场景：用户已有目标岗位，想爬取人脉并获取内推话术。
tags: career, referral, linkedin, networking, job-search
---

# Career Referral — 内推助手

你的任务很直接：**帮用户找到可内推的人，再给他一套能直接复制的话术**。

两步走：
1. **爬取**：用户提供范围（学校、目标公司、目标岗位），你在LinkedIn上爬人脉信息，输出Excel
2. **话术**：根据用户背景和目标岗位，输出几版通用内推话术模板

---

## 核心原则

1. **用户提供范围，AI执行爬取**。不问"你有没有人脉"，问"你想搜哪个学校的校友、哪家公司、什么岗位"
2. **输出Excel**。人名、关系、履历、优先级标注，一目了然
3. **话术是通用模板**。不是给每个人的定制版，而是分场景的通用模板，用户拿到自己改人名就行
4. **爬不到就直说**。LinkedIn风控严，爬不了就告诉用户，不硬撑
5. **爬取后必须截图验证**。爬取完成后截图当前页面，确认提取的数据和页面内容一致
6. **搜索策略：三级优先级**。不是随便搜，而是先抓HR、再抓岗位匹配的、最后补普通校友

---

## 信息收集

**已有材料（直接复用）**：
- 目标岗位和公司（来自 job-hunt Step 5 或用户直接提供）
- 简历内容（来自 resume-craft Step 4）：学校、前公司、工作经历

**需要用户提供的新信息**：
1. **爬取范围**：
   - 学校名（全称）：想搜哪个学校的校友？
   - 目标公司名列表（全称）：想搜哪些公司的人？建议3-6家
   - **目标岗位关键词**（如"算法工程师""产品经理""数据科学家"）：用于优先匹配岗位相近的人
2. **LinkedIn 登录态**（必须）：检查 `linkedin_auth.json` 是否存在，不存在则引导用户登录
3. **爬取数量上限**（推荐）：单次爬多少人？建议20-40人

---

## LinkedIn 爬取

### 技术方案

- **工具**：Playwright（Python sync API）
- **登录持久化**：使用 `storage_state` 保存/复用登录态
- **速率控制**：每次请求间隔 2-3 秒
- **风险提示**：爬取前必须告知用户——LinkedIn对异常流量有风控，建议小批量使用

### 搜索策略：三级优先级

不要只用 `"公司名 学校名"` 搜一次就完事。按优先级分三轮搜，确保先抓到最有价值的人：

**第一轮（最高优先级）：找HR**
- 搜索 `"公司名 HR 学校名"`
- 搜索 `"公司名 recruiter 学校名"`
- 搜索 `"公司名 talent 学校名"`
- 提取到的人，备注标注为 **"HR-优先联系"**
- HR是内推链路中最关键的一环，他们直接掌握招聘需求和流程

**第二轮（高优先级）：找岗位相近的人**
- 搜索 `"公司名 目标岗位关键词 学校名"`
- 如 `"ByteDance 算法工程师 中国人民大学"`
- 提取到的人，备注标注为 **"岗位匹配-优先"**
- 同岗位/相近岗位的校友最了解 team 情况，内推成功率最高

**第三轮（兜底）：普通校友**
- 搜索 `"公司名 学校名"`
- 提取到的人，备注标注为 **"校友"**
- 补充数量，扩大覆盖

**去重规则**：同一人在多轮中出现，保留最高优先级的标注。

### 可复用脚本模板

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LinkedIn 内推人脉爬取脚本
策略：三级优先级搜索（HR → 岗位匹配 → 普通校友）
"""

import time
import os
import sys
import re
from datetime import datetime

if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


def save_excel(data, filename):
    try:
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill
    except ImportError:
        print("Installing openpyxl...")
        os.system("pip install openpyxl -q")
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Referral Contacts"

    headers = ["序号", "姓名", "关系", "当前职位", "当前公司", "教育背景", "LinkedIn链接", "备注"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="E94560", end_color="E94560", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for i, item in enumerate(data, 1):
        ws.append([
            i, item.get("name", ""), item.get("relation", ""),
            item.get("title", ""), item.get("company", ""),
            item.get("education", ""), item.get("link", ""), item.get("note", "")
        ])

    for col, width in zip(['A','B','C','D','E','F','G','H'], [8,18,12,35,25,30,45,20]):
        ws.column_dimensions[col].width = width

    wb.save(filename)
    print(f"Saved: {os.path.abspath(filename)}")


def is_logged_in(page):
    try:
        url = page.url
        if "feed" in url or "/in/" in url:
            return True
        if page.query_selector('input[placeholder*="Search"]') or page.query_selector('input[aria-label*="Search"]'):
            return True
        if page.query_selector('text=Sign in') or page.query_selector('text=Join now') or page.query_selector('text=登录'):
            return False
    except:
        pass
    return False


def get_priority(person, target_role=""):
    """
    判断联系人的优先级
    返回: (priority_score, note)
    priority_score: 0=HR最高, 1=岗位匹配, 2=普通校友
    """
    title = person.get("title", "").lower()
    company = person.get("company", "").lower()
    full_text = f"{title} {company}"

    # HR相关关键词
    hr_keywords = ["hr", "recruiter", "talent", "招聘", "人力资源", "human resource",
                   "hiring", "staffing", "head of people", "people ops"]
    for kw in hr_keywords:
        if kw in full_text:
            return (0, "HR-优先联系")

    # 岗位匹配
    if target_role:
        target = target_role.lower()
        # 拆分成多个关键词（如"算法工程师"→["算法","工程师"]）
        parts = re.findall(r'[\u4e00-\u9fff]+|[a-zA-Z]+', target)
        match_count = sum(1 for p in parts if len(p) >= 2 and p in full_text)
        if match_count >= max(1, len(parts) // 2):
            return (1, "岗位匹配-优先")

    return (2, "校友")


def clean_person(p):
    """Python层面清理提取到的人员数据"""
    name = p.get("name", "")
    title = p.get("title", "")
    company = p.get("company", "")

    name = re.sub(r'\s*[·•]\s*\d+(?:nd|rd|st|th)?\+?\s*$', '', name, flags=re.I).strip()
    name = re.sub(r'\s*[·•]\s*\d+\s*度\+?\s*$', '', name).strip()
    name = re.sub(r'\s*[✓√]\s*$', '', name).strip()

    if re.match(r'^[·•]\s*\d+(?:nd|rd|st|th)?\+?\s*$', title, re.I):
        title = ""
    if re.match(r'^[·•]\s*\d+\s*度\+?\s*$', title):
        title = ""

    location_patterns = [
        r'^(beijing|shanghai|shenzhen|guangzhou|hangzhou|nanjing|chengdu|wuhan|xian|tianjin)',
        r'^(北京|上海|深圳|广州|杭州|南京|成都|武汉|西安|天津|中国|美国|英国|海淀区|朝阳区|东城区|西城区|浦东新区|南山区)',
        r'^(mountain view|new york|london|san francisco|seattle|boston|area|china|united states|california)',
        r'^(haidian|chaoyang|dongcheng|xicheng|pudong|nanshan|futian|luohu|minhang|songjiang)',
    ]
    for pat in location_patterns:
        if re.match(pat, title, re.I) and len(title) < 50:
            title = ""
            break

    for pat in location_patterns:
        if re.match(pat, company, re.I) and len(company) < 50:
            company = ""
            break

    if title.lower().startswith('current:'):
        title = title[8:].strip()

    if not title and company:
        title = company
        company = ""

    p["name"] = name
    p["title"] = title
    p["company"] = company
    return p


def extract_people(page):
    """从LinkedIn people搜索结果页面提取人员信息"""
    batch = page.evaluate("""
        () => {
            const results = [];
            const seenLinks = new Set();

            const allLinks = document.querySelectorAll('a[href*="/in/"]');
            const containers = new Set();
            allLinks.forEach(link => {
                if (link.href.includes('/company/')) return;
                let c = link.closest('li');
                if (!c) c = link.closest('div[class*="search-result"], div[class*="entity-result"], div[class*="lockup"]');
                if (!c) {
                    let p = link.parentElement;
                    for (let i = 0; i < 5 && p; i++) {
                        if (p.querySelector('a[href*="/in/"]') && p.innerText.length > 40) {
                            c = p; break;
                        }
                        p = p.parentElement;
                    }
                }
                if (c) containers.add(c);
            });

            Array.from(containers).forEach(item => {
                try {
                    const linkEl = item.querySelector('a[href*="/in/"]');
                    if (!linkEl) return;
                    const link = linkEl.href.split('?')[0];
                    if (seenLinks.has(link) || link.includes('/company/')) return;
                    seenLinks.add(link);

                    const lines = item.innerText.split(String.fromCharCode(10))
                        .map(t => t.trim())
                        .filter(t => t.length > 0 && t.length < 200);

                    const isNoise = t => {
                        if (/^Connect$/i.test(t)) return true;
                        if (/^Message$/i.test(t)) return true;
                        if (t === '\u5173\u6ce8' || t === 'Follow') return true;
                        if (t === '\u9884\u7ea6') return true;
                        if (/are mutual connections/.test(t)) return true;
                        if (/is a mutual connection/.test(t)) return true;
                        if (/\u548c\u5176\u4ed6.*\u4f4d\u5171\u540c\u597d\u53cb/.test(t)) return true;
                        return false;
                    };
                    const cleanLines = lines.filter(t => !isNoise(t));

                    let name = linkEl.innerText.trim();
                    if (!name || name.length > 50) {
                        name = cleanLines[0] || '';
                    }

                    let title = '';
                    let company = '';
                    let startIdx = 0;
                    for (let i = 0; i < cleanLines.length; i++) {
                        if (cleanLines[i] === name || cleanLines[i].includes(name)) {
                            startIdx = i + 1;
                            break;
                        }
                    }

                    for (let i = startIdx; i < cleanLines.length; i++) {
                        const t = cleanLines[i];
                        if (t.toLowerCase().startsWith('current:')) continue;

                        if (!title && t.length > 2 && t.length < 120) {
                            title = t;
                        } else if (title && !company && t.length > 1 && t.length < 80 && t !== title) {
                            company = t;
                            break;
                        }
                    }

                    if (title && !company) {
                        const atMatch = title.match(/(.+?)\s+(?:at|@|\u5728)\s+(.+)/i);
                        if (atMatch) {
                            title = atMatch[1].trim();
                            company = atMatch[2].trim();
                        }
                    }

                    if (name && name.length > 0 && name.length < 50) {
                        results.push({name, title, company, link});
                    }
                } catch (e) {}
            });

            return results;
        }
    """)
    return [clean_person(p) for p in batch]


def scrape_linkedin_referrals(school_name, company_list, target_role="", max_results=30):
    """
    爬取LinkedIn人脉信息（三级优先级）
    school_name: 学校名，如 "Renmin University of China"
    company_list: 目标公司列表，如 ["ByteDance", "Baidu"]
    target_role: 目标岗位关键词，如 "算法工程师"
    max_results: 最大提取人数
    """
    from playwright.sync_api import sync_playwright

    auth_file = "linkedin_auth.json"
    use_existing = os.path.exists(auth_file)
    all_people = {}  # name -> person dict，用于去重并保留最高优先级

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        if use_existing:
            context = browser.new_context(storage_state=auth_file, viewport={"width": 1920, "height": 1080})
        else:
            context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()

        if not use_existing:
            page.goto("https://www.linkedin.com/login")
            print("Please login to LinkedIn in the browser window...")
            for i in range(30):
                time.sleep(2)
                if is_logged_in(page):
                    print(f"Login detected at {i*2}s")
                    break
                if i % 5 == 0:
                    print(f"  Waiting... {i*2}s")
            else:
                print("Timeout. Login not detected.")
                browser.close()
                return []
            context.storage_state(path=auth_file)
            print("Login saved.")

        # 为每家公司生成三级搜索查询
        search_rounds = []
        for company in company_list:
            # Round 1: HR
            search_rounds.append((company, f"{company} HR {school_name}", "hr"))
            search_rounds.append((company, f"{company} recruiter {school_name}", "hr"))
            search_rounds.append((company, f"{company} talent {school_name}", "hr"))
            # Round 2: 岗位匹配
            if target_role:
                search_rounds.append((company, f"{company} {target_role} {school_name}", "role"))
            # Round 3: 普通校友
            search_rounds.append((company, f"{company} {school_name}", "alumni"))

        for company, query, round_type in search_rounds:
            if len(all_people) >= max_results:
                break

            print(f"\n--- [{round_type.upper()}] {query} ---")
            encoded = query.replace(' ', '%20')
            url = f"https://www.linkedin.com/search/results/people/?keywords={encoded}&origin=GLOBAL_SEARCH_HEADER"
            page.goto(url, timeout=60000)
            time.sleep(5)

            for i in range(8):
                batch = extract_people(page)
                new_count = 0
                for person in batch:
                    name = person.get("name", "").strip()
                    if not name or len(name) >= 50:
                        continue

                    # 判断优先级
                    priority, note = get_priority(person, target_role)

                    if name not in all_people:
                        all_people[name] = {
                            **person,
                            "relation": "校友",
                            "education": school_name,
                            "priority": priority,
                            "note": note,
                        }
                        new_count += 1
                    else:
                        # 同一人出现，保留更高优先级
                        existing = all_people[name]
                        if priority < existing["priority"]:
                            existing["priority"] = priority
                            existing["note"] = note
                            # 更新职位/公司信息（取更完整的）
                            if person.get("title") and len(person.get("title", "")) > len(existing.get("title", "")):
                                existing["title"] = person["title"]
                            if person.get("company") and len(person.get("company", "")) > len(existing.get("company", "")):
                                existing["company"] = person["company"]

                if new_count > 0:
                    print(f"  +{new_count}, total unique: {len(all_people)}")

                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                time.sleep(2)
                if len(all_people) >= max_results:
                    break

        browser.close()

    # 按优先级排序：HR(0) > 岗位匹配(1) > 校友(2)
    sorted_people = sorted(all_people.values(), key=lambda p: p["priority"])

    # 移除内部字段后保存
    output = []
    for p in sorted_people[:max_results]:
        output.append({
            "name": p["name"],
            "relation": p["relation"],
            "title": p["title"],
            "company": p["company"],
            "education": p["education"],
            "link": p["link"],
            "note": p["note"],
        })

    fname = f"referral-contacts-{datetime.now().strftime('%Y%m%d')}.xlsx"
    save_excel(output, fname)

    # 统计
    hr_count = sum(1 for p in output if "HR" in p["note"])
    role_count = sum(1 for p in output if "岗位匹配" in p["note"])
    alumni_count = sum(1 for p in output if p["note"] == "校友")
    print(f"\nDone! HR: {hr_count}, Role-match: {role_count}, Alumni: {alumni_count}")
    print(f"Saved: {fname}")
    return output


if __name__ == "__main__":
    scrape_linkedin_referrals(
        school_name="Renmin University of China",
        company_list=["ByteDance", "Baidu", "Tencent", "Alibaba"],
        target_role="算法工程师",
        max_results=30
    )
```

### 爬取流程要点

1. **登录态管理**
   - 首次运行：浏览器弹窗让用户手动登录，脚本检测登录成功后自动保存 `linkedin_auth.json`
   - 后续运行：直接加载 `linkedin_auth.json`，无需再次登录
   - 如果登录过期：删除 `linkedin_auth.json` 重新登录

2. **三级搜索优先级**
   - 每家公司先搜HR（3个查询词），再搜岗位匹配（1个查询词），最后搜普通校友（1个查询词）
   - 去重时保留最高优先级的标注
   - Excel最终按优先级排序输出

3. **优先级判断逻辑**
   - **HR**：title/company中包含 "HR"、"recruiter"、"talent"、"招聘"、"人力资源" 等关键词
   - **岗位匹配**：title/company中包含目标岗位关键词（支持中英文混合匹配，不要求完全匹配，匹配一半以上关键词即算）
   - **校友**：其余所有人

4. **数据清理**
   - 清理关系度标识、地点文本、title/company错位

5. **截图验证**
   - 每次搜索后截图保存，人工抽查验证数据质量

---

## 输出：Excel 表格

爬取完成后，生成Excel文件：`referral-contacts-{日期}.xlsx`

**表格列**：

| 列名 | 说明 |
|------|------|
| 序号 | 1, 2, 3... |
| 姓名 | LinkedIn上显示的姓名 |
| 关系 | 校友 |
| 当前职位 | 职位名称 |
| 当前公司 | 公司名称 |
| 教育背景 | 学校名 |
| LinkedIn链接 | 个人主页链接 |
| 备注 | **HR-优先联系** / **岗位匹配-优先** / **校友** |

**输出顺序**：按优先级排序，HR在最前面，其次是岗位匹配的，最后是普通校友。

**Excel生成方式**：脚本中使用 Python openpyxl 生成，保存到当前工作目录。

---

## 内推话术模板

根据用户背景和目标岗位，输出通用话术模板。用户拿到后只需要改人名和公司名就能发。

### 模板结构（所有版本通用）

一条内推请求，包含五个要素（控制在一屏以内）：
1. **开场**：简短问候或说明怎么认识对方
2. **意图**：一句话说明来意
3. **价值交换**：让对方知道帮这个忙不亏
4. **降低行动成本**：告诉对方具体需要做什么
5. **退路**：给对方拒绝的出口

### 三个版本

**版本A：熟人直推**（前同事、同学、很熟的朋友）

> 哈喽XX，最近怎么样？
> 看到你们公司在招[目标岗位]，我跟这个岗位匹配度挺高的，想问问你能不能帮忙内推一下？
> 我简历和JD分析都准备好了，你只需要走个内推流程，不占用你什么时间。如果不方便也完全理解～

**版本B：校友/弱关系**（校友、共同认识的人）

> 哈喽XX，我是[你的学校/前公司]的[你的名字]，注意到你在[目标公司]做[方向]。
> 最近看到你们公司在招[目标岗位]，我的背景和岗位要求挺匹配的，想问问能不能麻烦你帮忙内推一下？
> 简单说下我的背景：[2-3句话核心经历]。我已经把简历和目标岗位整理好了，你只需要转发一下。
> 如果你最近不方便也没关系，完全理解。先谢啦！

**版本C：冷接触**（完全不认识，LinkedIn刚加的好友）

> 哈喽XX，我是[你的名字]，[简单身份标签]。关注你有一段时间了，[具体指出对方哪条内容对你有启发]，很受用。
> 最近看到你们公司在招[目标岗位]，我的背景和岗位要求匹配度挺高的，想冒昧问一下，能不能请你帮忙内推？
> 我的核心背景是：[2-3句话，突出与岗位相关的亮点]。我已经准备好了简历和岗位分析，如果你觉得合适，内推流程我来配合。
> 如果不方便也完全理解，打扰了。不管怎样，继续支持你分享的内容。

**版本D：找HR**（专门给HR发的版本）

> 哈喽XX，我是[你的名字]，[学校/前公司]背景，目前在看[目标岗位]的机会。
> 看到贵司在招这个岗位，我的背景和JD要求匹配度挺高的。想冒昧问一下，能否帮我推荐到合适的团队？
> 简历和JD分析已准备好，如果方便的话我可以发你详细材料。打扰了，期待你的回复！

### 定制规则

- 转行者：在模板中加入一句话解释转行原因
- 有gap/空窗期：不提，内推阶段不是解释的时候
- 语气：短句、换行，像微信聊天
- **给HR发的话术要更直接**，不需要套近乎，突出匹配度和行动力即可

---

## 输出流程

**第一步：确认范围**
- 用户提供学校名、目标公司列表、目标岗位关键词
- 确认LinkedIn登录态

**第二步：执行爬取**
- 按三级优先级搜索并提取数据
- 去重，保留最高优先级标注
- 按优先级排序

**第三步：输出Excel**
- 生成Excel表格，HR和岗位匹配的排在前面
- 预览前10条确认数据质量

**第四步：输出话术模板**
- 根据用户背景和目标岗位，输出四个版本的通用话术（含HR专用版）
- 告诉用户：优先联系备注为"HR-优先联系"的人，其次是"岗位匹配-优先"

---

## 手动搜索指南（爬取不可用时）

如果用户没有LinkedIn账号或爬取被限制，提供手动搜索步骤：

**优先级搜索**：
1. 打开LinkedIn → 顶部搜索框输入 `"公司名 HR 学校名"`
2. 筛选 People，记录HR姓名和职位
3. 再搜 `"公司名 目标岗位 学校名"`，记录岗位匹配的人
4. 最后搜 `"公司名 学校名"`，补充普通校友

---

## 语言风格

**直接、工具化。不聊天，给结果。**

- ✅ "你在LinkedIn上搜到了23个校友在字节，其中3个HR、8个岗位匹配的，表格已按优先级排好"
- ✅ "四个版本的话术，你根据关系远近选，HR用D版"
- ❌ "我们先来聊聊你的人脉基础..."
- ❌ "本方案包含四个模块..."

**用「你」不用「用户」。**

---

## 注意事项

1. **爬取前必须告知风险**。LinkedIn可能触发风控，由用户决定是否执行
2. **不爬隐私数据**。只提取公开可见信息，不碰联系方式
3. **爬不了就停**。遇到验证码或限制立即停止，切换到手动搜索指南
4. **严禁编造人脉**。Excel里每条记录都必须来自真实爬取结果
5. **Excel是主要输出**。爬取结果必须生成Excel，不能只输出文字列表
6. **话术是模板**。不是给每个人的定制版，用户拿到自己改人名

---

## Skill 间协同

- **上游衔接**：从 job-hunt（Step 5）衔接而来，复用搜索到的目标岗位列表和公司信息
- **数据复用**：在同一 career-transition 套件对话中，直接复用 resume-craft（Step 4）的简历内容（学校、前公司）
- **独立使用**：用户直接提供目标公司、学校、岗位信息，也可以启动爬取和话术生成
- **衔接 interview-story-bank**：内推拿到面试后，自然衔接到故事银行准备面试
