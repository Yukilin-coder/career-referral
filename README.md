# Career Referral — 内推助手

基于LinkedIn的人脉爬取+内推话术生成工具。帮你快速找到可内推的人，并给你一套能直接复制的话术。

## 功能

1. **LinkedIn人脉爬取**：自动搜索目标公司+学校的校友，输出Excel表格
2. **三级优先级**：HR优先 → 岗位匹配优先 → 普通校友
3. **内推话术模板**：4个版本（熟人/校友/冷接触/HR专用）

## 安装

作为Claude Code Skill使用：

```bash
# 方式1：克隆到个人skills目录
mkdir -p ~/.claude/skills
git clone https://github.com/Yukilin-coder/career-referral.git ~/.claude/skills/career-referral

# 方式2：下载zip解压到 ~/.claude/skills/career-referral/
```

作为独立Python脚本使用：

```bash
pip install playwright openpyxl
playwright install chromium
python scripts/linkedin_referral_scraper.py
```

## 快速开始

```python
from scripts.linkedin_referral_scraper import scrape_linkedin_referrals

results = scrape_linkedin_referrals(
    school_name="Renmin University of China",
    company_list=["ByteDance", "Baidu", "Tencent", "Alibaba"],
    target_role="算法工程师",
    max_results=30
)
```

## 输出示例

Excel表格包含以下列：

| 列名 | 说明 |
|------|------|
| 姓名 | LinkedIn上显示的姓名 |
| 关系 | 校友 |
| 当前职位 | 职位名称 |
| 当前公司 | 公司名称 |
| LinkedIn链接 | 个人主页链接 |
| 备注 | HR-优先联系 / 岗位匹配-优先 / 校友 |

## 风险提示

LinkedIn对异常流量有风控，建议：
- 单次爬取不超过40人
- 每次搜索间隔2-3秒
- 遇到验证码立即停止，切换手动搜索

## License

MIT
