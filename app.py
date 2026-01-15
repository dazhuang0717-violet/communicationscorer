import streamlit as st
import pandas as pd
import numpy as np
import google.generativeai as genai
import requests
from bs4 import BeautifulSoup
import plotly.express as px
from docx import Document
import io
import math
import time
import json
import re

# --- 配置页面 ---
st.set_page_config(
    page_title="传播价值 AI 评分系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 核心工具类与函数 (Backend Logic) ---

class ScorerEngine:
    """处理评分逻辑的核心引擎"""
    
    def __init__(self, api_key):
        self.api_key = api_key
        if self.api_key:
            genai.configure(api_key=self.api_key)
            self.model = genai.GenerativeModel('gemini-pro')

    def fetch_url_content(self, url):
        """爬虫模块：Jina Reader 优先，Requests 降级"""
        if not url or pd.isna(url):
            return ""
        
        # 1. 尝试 Jina Reader API (适合 LLM 的 Markdown)
        try:
            jina_url = f"https://r.jina.ai/{url}"
            response = requests.get(jina_url, timeout=8)
            if response.status_code == 200 and len(response.text) > 100:
                return response.text[:10000] # 截断以节省 Token
        except Exception as e:
            pass # Silent fail to fallback

        # 2. 降级方案: Requests + BS4
        try:
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')
                # 提取所有 P 标签
                text = " ".join([p.get_text() for p in soup.find_all('p')])
                return text[:10000]
        except Exception as e:
            return f"Error fetching content: {str(e)}"
        
        return ""

    def calculate_volume_quality(self, views, interactions):
        """计算传播质量 (对数归一化)"""
        try:
            # 清洗数据：移除 'k', ',', '+' 等非数字字符
            def clean_num(x):
                if isinstance(x, str):
                    x = re.sub(r'[kK]', '000', x) # 简易处理 10k -> 10000
                    x = re.sub(r'[^\d\.]', '', x)
                return float(x) if x else 0.0

            v = clean_num(views)
            i = clean_num(interactions)
            
            # Score = min(10, log10(浏览量 + 互动量 * 5 + 1) * 1.5)
            raw_score = math.log10(v + i * 5 + 1) * 1.5
            return min(10.0, round(raw_score, 1))
        except:
            return 0.0

    def get_media_tier_score(self, media_name, tiers_config):
        """计算媒体分级分数"""
        if not media_name or pd.isna(media_name):
            return 5 # 默认分
        
        m_name = str(media_name).lower().strip()
        
        for tier_name, tier_list in tiers_config.items():
            # 检查媒体名是否在配置列表中 (模糊匹配)
            for configured_media in tier_list:
                if configured_media and configured_media in m_name:
                    if tier_name == 'tier1': return 10
                    if tier_name == 'tier2': return 8
        return 5 # Tier 3 / Others

    def analyze_content_with_ai(self, content, key_message, project_desc, audience_mode, media_name):
        """集成 AI 调用：一次性请求获取 KM、获客、受众精准度"""
        if not self.api_key:
            return 0, 0, 0, "API Key Missing"

        prompt = f"""
        你是一个专业的公关传播分析师。请基于以下输入信息对一篇文章进行评分。
        
        【输入信息】
        1. 目标受众模式: {audience_mode}
        2. 媒体名称: {media_name}
        3. 核心传播信息 (Key Message): {key_message}
        4. 项目描述: {project_desc}
        5. 文章/网页内容: 
        {content[:3000]}... (内容截断)

        【任务】
        请分析并返回以下 3 个维度的分数（0-10分），并严格按照 JSON 格式返回：
        1. km_score: 文章是否有效传递了核心信息 '{key_message}'？(0=完全未提及, 10=深度且准确传递)
        2. acquisition_score: 基于项目描述，这篇文章对目标受众的获客吸引力如何？(0=无吸引力, 10=极强吸引力)
        3. audience_precision_score: 考虑到媒体 '{media_name}' 和目标受众 '{audience_mode}'，受众精准度如何？(0=完全错配, 10=非常精准)

        【输出格式】
        仅返回 JSON 字符串，不要包含 Markdown 格式（如 ```json）。格式如下：
        {{"km_score": 8, "acquisition_score": 7, "audience_precision_score": 9}}
        """

        try:
            response = self.model.generate_content(prompt)
            # 清洗返回的文本，确保它是纯 JSON
            clean_text = response.text.replace('```json', '').replace('```', '').strip()
            data = json.loads(clean_text)
            return (
                data.get('km_score', 0), 
                data.get('acquisition_score', 0), 
                data.get('audience_precision_score', 0), 
                "Success"
            )
        except Exception as e:
            # Fallback for errors
            return 0, 0, 0, f"AI Error: {str(e)}"

# --- 侧边栏配置 (Sidebar) ---
with st.sidebar:
    st.header("⚙️ 系统配置")
    
    # API 配置
    api_key = st.text_input("Gemini API Key", type="password", help="需从 Google AI Studio 获取")
    if not api_key:
        st.warning("请先输入 API Key 以启用 AI 功能")
        st.markdown("[点击这里免费获取 Gemini API Key](https://aistudio.google.com/app/apikey)")

    st.markdown("---")
    st.subheader("📋 项目基础信息")
    project_key_message = st.text_input("核心信息 (Key Message)", value="AI 赋能医疗创新")
    project_desc = st.text_area("项目描述 (用于评估获客)", value="这是一款革命性的 AI 诊断工具，旨在帮助医生提高效率。")
    audience_mode = st.radio("目标受众模式", ["大众 (General)", "患者 (Patient)", "医疗专业人士 (HCP)"])

    st.markdown("---")
    st.subheader("🏆 媒体分级配置")
    st.caption("输入媒体名称关键词，用逗号分隔")
    
    tier1_input = st.text_area("Tier 1 (10分)", value="人民日报, 新华社, 36Kr")
    tier2_input = st.text_area("Tier 2 (8分)", value="动脉网, 丁香园, 虎嗅")
    tier3_input = st.text_area("Tier 3 (5分 - 默认)", disabled=True, value="其他未列出媒体")

    # 处理分级列表
    def parse_tiers(text):
        return [x.strip().lower() for x in text.split(',') if x.strip()]
    
    tier_config = {
        'tier1': parse_tiers(tier1_input),
        'tier2': parse_tiers(tier2_input)
    }

# --- 主界面 (Main) ---

st.title("📡 传播价值 AI 评分系统")
st.markdown("##### Communication Value AI Scorer | Powered by Gemini & Streamlit")

# 1. 顶部公式展示
with st.expander("查看核心算法公式", expanded=False):
    st.latex(r'''
    Total = 0.5 \times TrueDemand + 0.2 \times Acquisition + 0.3 \times Volume
    ''')
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**真需求 (True Demand)**")
        st.latex(r'''= 0.6 \times KM\_Match + 0.4 \times Precision''')
    with c2:
        st.markdown("**获客效能 (Acquisition)**")
        st.latex(r'''= AI\_Evaluated\_Score (0-10)''')
    with c3:
        st.markdown("**声量 (Volume)**")
        st.latex(r'''= 0.6 \times Quality + 0.4 \times Media\_Tier''')

# 初始化引擎
engine = ScorerEngine(api_key)

# 2. 标签页：Word 预检 vs CSV 批量
tab1, tab2 = st.tabs(["📝 Word 单篇预检", "🚀 CSV 批量评分"])

# --- TAB 1: Word 预检 ---
with tab1:
    st.info("上传新闻稿 Word 文档，AI 将预判核心信息传递情况。")
    uploaded_word = st.file_uploader("上传 .docx 文件", type=['docx'])
    
    if uploaded_word and api_key:
        if st.button("开始预检分析"):
            with st.spinner("AI 正在阅读文档..."):
                try:
                    doc = Document(uploaded_word)
                    full_text = "\n".join([para.text for para in doc.paragraphs])
                    
                    # 复用 AI 逻辑，虽然 Word 没有 URL 和 媒体名，我们传入 Dummy 值
                    km, acq, prec, status = engine.analyze_content_with_ai(
                        full_text, project_key_message, project_desc, audience_mode, "Internal Draft"
                    )
                    
                    col_res1, col_res2 = st.columns(2)
                    with col_res1:
                        st.metric("核心信息匹配度 (KM)", f"{km}/10")
                        st.progress(km/10)
                    with col_res2:
                        st.metric("预期获客吸引力", f"{acq}/10")
                        st.progress(acq/10)
                    
                    st.success("分析完成！建议优化方向：如果 KM 分数低，请在首段强化核心关键词。")
                    
                except Exception as e:
                    st.error(f"解析错误: {e}")

# --- TAB 2: CSV 批量评分 ---
with tab2:
    st.markdown("**上传媒体监测报表 (CSV)**。必须包含列: `媒体名称`, `URL`, `互动量`, `浏览量` (列名可模糊匹配)")
    # 提供示例数据下载
    example_data = """媒体名称,URL,互动量,浏览量
36Kr,[https://36kr.com/p/244321,120,5000](https://36kr.com/p/244321,120,5000)
动脉网,[https://vcbeat.top/12345,50,2000](https://vcbeat.top/12345,50,2000)
新浪微博,[https://weibo.com/123,500,10000](https://weibo.com/123,500,10000)"""
    
    st.download_button(
        "📥 下载示例 CSV 模板",
        example_data,
        "template.csv",
        "text/csv",
        help="点击下载一个测试用的 CSV 文件"
    )

    uploaded_csv = st.file_uploader("上传 .csv 文件", type=['csv'])

    if uploaded_csv:
        try:
            df = pd.read_csv(uploaded_csv)
            # 列名标准化处理 (Strip spaces)
            df.columns = df.columns.str.strip()
            
            # 简单的列名映射检查
            required_cols = ['媒体名称', 'URL', '互动量', '浏览量']
            missing_cols = [col for col in required_cols if col not in df.columns]
            
            if missing_cols:
                st.error(f"CSV 缺少必要列: {missing_cols}")
            else:
                st.dataframe(df.head(3), use_container_width=True)
                
                if st.button("开始 AI 全量评分", type="primary"):
                    if not api_key:
                        st.error("请先在左侧配置 API Key")
                        st.stop()

                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    results = []
                    total_rows = len(df)

                    # 逐行处理
                    for index, row in df.iterrows():
                        status_text.text(f"正在处理第 {index + 1}/{total_rows} 行: {row['媒体名称']}...")
                        
                        # A. 基础计算
                        vol_quality = engine.calculate_volume_quality(row['浏览量'], row['互动量'])
                        tier_score = engine.get_media_tier_score(row['媒体名称'], tier_config)
                        volume_total = 0.6 * vol_quality + 0.4 * tier_score
                        
                        # B. 爬虫与 AI
                        content = engine.fetch_url_content(row['URL'])
                        
                        if content:
                            km_score, acq_score, prec_score, msg = engine.analyze_content_with_ai(
                                content, project_key_message, project_desc, audience_mode, row['媒体名称']
                            )
                        else:
                            km_score, acq_score, prec_score = 0, 0, 0
                            msg = "URL Fail"

                        # C. 聚合计算
                        true_demand = 0.6 * km_score + 0.4 * prec_score
                        # Total = 0.5 * Demand + 0.2 * Acquisition + 0.3 * Volume
                        total_score = (0.5 * true_demand) + (0.2 * acq_score) + (0.3 * volume_total)

                        # 保存结果
                        results.append({
                            "媒体名称": row['媒体名称'],
                            "Total Score": round(total_score, 2),
                            "真需求 (Demand)": round(true_demand, 2),
                            "获客 (Acq)": acq_score,
                            "声量 (Volume)": round(volume_total, 2),
                            "KM匹配": km_score,
                            "精准度": prec_score,
                            "传播质量": vol_quality,
                            "媒体分级": tier_score,
                            "状态": msg
                        })
                        
                        # 更新进度条
                        progress_bar.progress((index + 1) / total_rows)
                        # 为了演示效果，稍微 sleep 一下 (实际生产可去掉)
                        # time.sleep(0.1)

                    # --- 结果展示 ---
                    status_text.text("分析完成！")
                    res_df = pd.DataFrame(results)
                    
                    st.divider()
                    
                    # 1. Metrics
                    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                    col_m1.metric("媒体总数", len(res_df))
                    high_value_count = len(res_df[res_df['Total Score'] >= 8])
                    col_m2.metric("高价值 (≥8分)", high_value_count, delta_color="normal")
                    col_m3.metric("平均得分", round(res_df['Total Score'].mean(), 2))
                    col_m4.metric("中位数得分", round(res_df['Total Score'].median(), 2))

                    # 2. Charts
                    col_chart1, col_chart2 = st.columns([2, 1])
                    with col_chart1:
                        st.subheader("📊 媒体得分排行")
                        fig = px.bar(
                            res_df.sort_values('Total Score', ascending=True), 
                            x='Total Score', 
                            y='媒体名称', 
                            orientation='h',
                            color='Total Score',
                            color_continuous_scale='Bluered'
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    
                    with col_chart2:
                        st.subheader("因子贡献分析")
                        # 简单的散点图看 声量 vs 需求
                        fig2 = px.scatter(
                            res_df,
                            x='声量 (Volume)',
                            y='真需求 (Demand)',
                            hover_name='媒体名称',
                            size='Total Score',
                            color='获客 (Acq)'
                        )
                        st.plotly_chart(fig2, use_container_width=True)

                    # 3. Detail Data
                    st.subheader("📋 详细评分表")
                    st.dataframe(
                        res_df.style.background_gradient(subset=['Total Score'], cmap='Greens'),
                        use_container_width=True
                    )

                    # 4. Download
                    csv = res_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "📥 导出评分报告 (CSV)",
                        csv,
                        "ai_media_scoring_report.csv",
                        "text/csv",
                        key='download-csv'
                    )

        except Exception as e:
            st.error(f"处理 CSV 时发生错误: {e}")
