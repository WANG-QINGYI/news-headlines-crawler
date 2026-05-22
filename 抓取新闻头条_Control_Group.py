"""
新闻头条爬虫脚本 - 控制组版本
用于经济研究项目：采集 7 个国家从 2010 年到 2025 年的经济类新闻头条
支持自动重试、智能休眠、进度跟踪
"""

# ================= 自动安装依赖包 =================
import subprocess
import sys
import os

def install_required_packages():
    """自动安装所需的 Python 包"""
    required_packages = [
        'pandas',
        'requests',
        'python-dateutil',
        'openpyxl'
    ]
    
    print("📦 检查并安装依赖包...\n")
    for package in required_packages:
        try:
            __import__(package.replace('-', '_'))
            print(f"  ✅ {package} 已安装")
        except ImportError:
            print(f"  ⬇️  正在安装 {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"  ✅ {package} 安装完成")
    print()

install_required_packages()

# ================= 导入库 =================
import pandas as pd
import time
import datetime
import random
from dateutil.relativedelta import relativedelta
import urllib.parse
import requests
import xml.etree.ElementTree as ET
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ================= 爬虫配置区域 =================
# 1. 目标国家列表
countries = ["Peru", "Poland", "Czech Republic", "Thailand", "Malaysia", "Uruguay", "Norway"]

# 2. 搜索关键词的后缀 (英文搜索，国际化视野)
search_suffix = "economy OR exchange rate OR central bank OR GDP OR inflation OR monetary policy"

# 3. 设置要抓取的起始和结束日期 (2010-01 到 2025-12)
start_date = datetime.date(2010, 1, 1)
end_date = datetime.date(2025, 12, 31)

# 4. 输出文件路径 (使用相对路径，与脚本同目录)
script_dir = os.path.dirname(os.path.abspath(__file__))
output_file = os.path.join(script_dir, "News_Headlines_Control_Group.xlsx")
backup_file = os.path.join(script_dir, "News_Headlines_Control_Group_Backup.xlsx")

# 5. 爬虫参数配置
REQUEST_TIMEOUT = 15  # 请求超时时间(秒)
BASE_SLEEP_TIME = 8   # 基础休眠时间(秒)
MAX_RETRIES = 3       # 单个请求最多重试次数
RANDOM_DELAY = True   # 是否添加随机延迟 (防止规律性被检测)

# ============================================

# ================= 工具函数 =================

def setup_session_with_retries():
    """
    创建带重试机制的 requests Session
    当网络不稳定或临时被限流时自动重试
    """
    session = requests.Session()
    
    # 配置重试策略：HTTP 连接和 HTTPS 连接都启用
    retry_strategy = Retry(
        total=MAX_RETRIES,
        status_forcelist=[429, 500, 502, 503, 504],  # 这些状态码会触发重试
        method_whitelist=["GET"],  # 只重试 GET 请求
        backoff_factor=1  # 重试间隔: 1s, 2s, 4s...
    )
    
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    
    return session

def get_random_user_agent():
    """返回随机的 User-Agent，避免被识别为机器人"""
    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15"
    ]
    return random.choice(user_agents)

def adaptive_sleep(base_time=BASE_SLEEP_TIME):
    """
    自适应休眠函数
    基础时间 + 随机延迟，防止被检测为爬虫
    """
    if RANDOM_DELAY:
        random_delay = random.uniform(0.5, 2.5)
        actual_sleep = base_time + random_delay
    else:
        actual_sleep = base_time
    
    time.sleep(actual_sleep)

def fetch_news_from_google(query, retry_count=0):
    """
    从 Google News RSS 接口获取新闻
    包含自动重试和错误处理
    
    Args:
        query: 搜索查询字符串
        retry_count: 当前重试次数
    
    Returns:
        headlines (list): 提取到的新闻标题列表
        status (str): 状态信息 ('success', 'no_news', 'error')
    """
    try:
        # 构建 RSS URL
        encoded_query = urllib.parse.quote(query)
        rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=en-US&gl=US&ceid=US:en"
        
        # 随机 User-Agent
        headers = {
            "User-Agent": get_random_user_agent(),
            "Accept": "application/rss+xml,application/xml,application/atom+xml",
            "Accept-Language": "en-US,en;q=0.9",
            "Referer": "https://news.google.com/"
        }
        
        # 获取会话并发送请求
        session = setup_session_with_retries()
        response = session.get(rss_url, headers=headers, timeout=REQUEST_TIMEOUT)
        
        # 检查响应状态
        if response.status_code != 200:
            return [], f"error_status_{response.status_code}"
        
        # 解析 XML
        root = ET.fromstring(response.content)
        items = root.findall('.//item')
        
        if not items:
            return [], "no_news"
        
        # 提取前 10 条新闻的标题 (比原来的 5 条多一些，提高数据质量)
        headlines = []
        for item in items[:10]:
            title_elem = item.find('title')
            if title_elem is not None and title_elem.text:
                headlines.append(title_elem.text)
        
        if not headlines:
            return [], "no_news"
        
        return headlines, "success"
    
    except ET.ParseError as e:
        print(f"        ⚠️  XML 解析错误: {e}")
        return [], "parse_error"
    
    except requests.exceptions.Timeout:
        print(f"        ⚠️  请求超时 (>{REQUEST_TIMEOUT}s)")
        if retry_count < MAX_RETRIES:
            print(f"        🔄 自动重试 ({retry_count + 1}/{MAX_RETRIES})...")
            adaptive_sleep(base_time=5)  # 超时后等待更长时间再重试
            return fetch_news_from_google(query, retry_count + 1)
        return [], "timeout"
    
    except requests.exceptions.ConnectionError as e:
        print(f"        ⚠️  连接错误: {e}")
        if retry_count < MAX_RETRIES:
            print(f"        🔄 自动重试 ({retry_count + 1}/{MAX_RETRIES})...")
            adaptive_sleep(base_time=5)
            return fetch_news_from_google(query, retry_count + 1)
        return [], "connection_error"
    
    except Exception as e:
        print(f"        ⚠️  未知错误: {type(e).__name__}: {e}")
        return [], "unknown_error"

def format_headlines(headlines):
    """将多条标题合并为一个字符串"""
    if not headlines:
        return "No News Found"
    return " | ".join(headlines)

# ============================================
# ================= 主程序 =================
# ============================================

all_results = []
current_date = start_date
total_months = 0
processed_months = 0

print("=" * 70)
print("🚀 新闻头条爬虫 - 控制组版本")
print("=" * 70)
print(f"📍 目标国家: {', '.join(countries)}")
print(f"📅 时间范围: {start_date.strftime('%Y-%m')} 至 {end_date.strftime('%Y-%m')}")
print(f"⏱️  基础休眠时间: {BASE_SLEEP_TIME} 秒")
print(f"🔄 最大重试次数: {MAX_RETRIES}")
print(f"💾 输出文件: {output_file}")
print("=" * 70)
print()

# 计算总月份数 (用于进度条)
temp_date = start_date
while temp_date <= end_date:
    total_months += 1
    temp_date += relativedelta(months=1)

total_tasks = total_months * len(countries)

print(f"📊 总任务数: {total_tasks} (共 {len(countries)} 个国家 × {total_months} 个月份)")
print()

try:
    for country_idx, country in enumerate(countries, 1):
        print(f"\n{'=' * 70}")
        print(f"[{country_idx}/{len(countries)}] 正在处理国家: {country}")
        print(f"{'=' * 70}")
        
        current_date = start_date
        month_in_country = 0
        
        while current_date <= end_date:
            month_in_country += 1
            processed_months += 1
            
            # 计算该月的起始和结束日期
            next_month = current_date + relativedelta(months=1)
            last_day = next_month - datetime.timedelta(days=1)
            
            # 格式化日期
            start_str = current_date.strftime("%Y-%m-%d")
            end_str = last_day.strftime("%Y-%m-%d")
            time_id_str = current_date.strftime("%Y-%m-01")
            
            # 构建搜索查询
            query = f'"{country}" {search_suffix} after:{start_str} before:{end_str}'
            
            # 进度显示
            progress_pct = (processed_months / total_tasks) * 100
            print(f"\n  [{progress_pct:5.1f}%] 🔍 {time_id_str} | ", end="")
            
            # 获取新闻
            headlines, status = fetch_news_from_google(query)
            combined_text = format_headlines(headlines)
            
            # 状态输出
            if status == "success":
                print(f"✅ 抓到 {len(headlines)} 条新闻")
            elif status == "no_news":
                print(f"⚠️  未找到新闻")
            elif status.startswith("error"):
                print(f"❌ HTTP 错误: {status}")
            elif status in ["timeout", "connection_error"]:
                print(f"⚠️  {status} - 已自动跳过")
            else:
                print(f"⚠️  {status}")
            
            # 存储结果
            all_results.append({
                "Country_ID": country,
                "Time_ID": time_id_str,
                "Headlines": combined_text,
                "Status": status
            })
            
            # 休眠 (防止被 Google 限流或封 IP)
            if current_date < end_date:  # 最后一个请求后不需要休眠
                adaptive_sleep(base_time=BASE_SLEEP_TIME)
            
            # 移动到下一个月
            current_date = next_month
        
        # 每个国家处理完成后的进度提示
        print(f"\n  ✅ {country} 的 {month_in_country} 个月份已全部处理完成")
        
        # ���家间的额外休眠 (防止频繁切换国家导致被检测)
        if country_idx < len(countries):
            print(f"  ⏱️  国家间休息中... (15 秒)")
            adaptive_sleep(base_time=15)

    # ============================================
    # 保存结果到 Excel
    # ============================================
    print(f"\n{'=' * 70}")
    print("💾 开始保存结果...")
    
    df_news = pd.DataFrame(all_results)
    df_news.to_excel(output_file, index=False, engine='openpyxl')
    
    print(f"✅ 数据已保存到: {output_file}")
    print(f"   总行数: {len(df_news)}")
    print(f"   包含列: {', '.join(df_news.columns.tolist())}")
    
    # 输出统计信息
    print(f"\n{'=' * 70}")
    print("📊 数据统计信息:")
    print(f"{'=' * 70}")
    
    success_count = (df_news['Status'] == 'success').sum()
    no_news_count = (df_news['Status'] == 'no_news').sum()
    error_count = len(df_news) - success_count - no_news_count
    
    print(f"  • 成功获取新闻: {success_count} 条记录")
    print(f"  • 未找到新闻: {no_news_count} 条记录")
    print(f"  • 失败/错误: {error_count} 条记录")
    print(f"  • 成功率: {(success_count/len(df_news)*100):.1f}%")
    
    print(f"\n{'=' * 70}")
    print("🎉 爬虫任务全部完成！")
    print(f"{'=' * 70}")

except KeyboardInterrupt:
    print("\n\n❌ 用户中断了程序")
    if len(all_results) > 0:
        print("💾 正在保存已收集的数据...")
        pd.DataFrame(all_results).to_excel(backup_file, index=False, engine='openpyxl')
        print(f"✅ 备份已保存到: {backup_file}")

except Exception as e:
    print(f"\n\n❌ 程序异常: {type(e).__name__}: {e}")
    import traceback
    traceback.print_exc()
    
    # 紧急备份
    if len(all_results) > 0:
        print("\n💾 执行紧急备份...")
        try:
            pd.DataFrame(all_results).to_excel(backup_file, index=False, engine='openpyxl')
            print(f"✅ 备份已保存到: {backup_file}")
        except Exception as backup_err:
            print(f"⚠️  备份失败: {backup_err}")

print("\n程序结束。")
