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
import json
import re
import time

# --- 1. 页面配置 ---
st.set_page_config(
    page_title="传播价值 AI 评分系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. UI 强制浅色模式 (深度修复版) ---
st.markdown("""
    <style>
        /* A. 全局容器强制白底黑字 */
        [data-testid="stAppViewContainer"] {
            background-color: #ffffff !important;
            color: #31333F !important;
        }
        [data-testid="stSidebar"] {
            background-color: #f8f9fa !important;
            border-right: 1px solid #e0e0e0;
        }
        
        /* B. 修复顶部导航栏和右上角按钮 */
        header[data-testid="stHeader"] {
            background-color: #ffffff !important;
            border-bottom: 1px solid #f0f2f6;
        }
        header[data-testid="stHeader"] button, 
        header[data-testid="stHeader"] a, 
        header[data-testid="stHeader"] svg {
            color: #31333F !important;
            fill: #31333F !important;
        }

        /* C. 修复文件上传组件 */
        [data-testid="stFileUploaderDropzone"] {
            background-color: #f8f9fa !important;
            border: 1px dashed #d1d5db !important;
        }
        [data-testid="stFileUploaderDropzone"] div, 
        [data-testid="stFileUploaderDropzone"] span, 
        [data-testid="stFileUploaderDropzone"] small,
        [data-testid="stFileUploaderDropzone"] p {
            color: #31333F !important;
        }
        [data-testid="stFileUploaderDropzone"] button {
            background-color: #ffffff !important;
            color: #31333F !important;
            border: 1px solid #d1d5db !important;
        }

        /* D. 通用文本和输入框修复 */
        h1, h2, h3, h4, h5, h6, p, span, div, label {
            color: #31333F !important;
        }
        .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
            color: #31333F !important;
            background-color: #ffffff !important;
            border: 1px solid #d1d5db;
        }
        .stTextInput input:focus, .stTextArea textarea:focus {
            border-color: #ff4b4b;
        }
        
        /* E. 修复 Metric 指标和表格颜色 */
        [data-testid="stMetricValue"], [data-testid="stMetricLabel"] {
            color: #31333F !important;
        }
        [data-testid="stDataFrame"] {
            color: #31333F !important;
        }
        [data-testid="stDataFrame"] svg {
             fill: #31333F !important;
        }

        /* F. 修复 Katex 公式颜色 (强制纯黑) */
        .katex {
            color: #000000 !important;
        }
        .katex-display {
            color: #000000 !important;
        }
        .katex-html {
            color: #000000 !important;
        }

        /* G. 隐藏元素 */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        
        .streamlit-expanderHeader {
            background-color: #f0f2f6 !important;
            color: #31333F !important;
        }
        .streamlit-expanderContent {
            background-color: #ffffff !important;
            color: #31333F !important;
        }
    </style>
""", unsafe_allow_html=True)

# 硬编码 API Key
INTERNAL_API_KEY = "AIzaSyCdz_GYYbJhSMtAL3vP_2_-TNTYX0bUt94"

# --- 3. 核心引擎 (Backend) ---

class ScorerEngine:
    def __init__(self):
        if INTERNAL_API_KEY:
            genai.configure(api_key=INTERNAL_API_KEY)
            # --- 修复点：将 gemini-pro 改为 gemini-1.5-flash ---
            self.model = genai.GenerativeModel('gemini-1.5-flash')

    def read_docx_content(self, file_obj):
        """增强版 Word 读取：同时读取段落和表格"""
        try:
            file_obj.seek(0)
            doc = Document(file_obj)
            full_text = []
            
            # 1. 读取段落
            for para in doc.paragraphs:
                if para.text.strip():
                    full_text.append(para.text.strip())
            
            # 2. 读取表格 (很多新闻稿在表格里)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if para.text.strip():
                                full_text.append(para.text.strip())
            
            return "\n".join(full_text)
        except Exception as e:
            return f"Error: {str(e)}"

    def fetch_url_content(self, url):
        """尝试获取 URL 内容，失败则返回空字符串"""
        if not url or pd.isna(url): return ""
        if not str(url).startswith('http'): return ""

        # 1. 尝试 Jina (效果最好)
        try:
            jina_url = f"https://r.jina.ai/{url}"
            response = requests.get(jina_url, timeout=5)
            if response.status_code == 200 and len(response.text) > 50:
                return response.text[:10000]
        except: pass 

        # 2. 尝试 Requests (兜底)
        try:
            headers = {'User-Agent': 'Mozilla/5.0'}
            response = requests.get(url, headers=headers, timeout=5)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')
                text = " ".join([p.get_text() for p in soup.find_all('p')])
                if len(text) > 50:
                    return text[:10000]
        except: pass
        
        return ""

    def calculate_volume_quality(self, views, interactions):
        try:
            def clean_num(x):
                if isinstance(x, str):
                    x = re.sub(r'[kK]', '000', x)
                    x = re.sub(r'[^\d\.]', '', x)
                return float(x) if x else 0.0
            v = clean_num(views)
            i = clean_num(interactions)
            # 简单的声量计算逻辑
            raw_score = math.log10(v + i * 5 + 1) * 1.5
            return min(10.0, round(raw_score, 1))
        except: return 0.0

    def get_media_tier_score(self, media_name, tiers_config):
        if not media_name or pd.isna(media_name): return 3
        m_name = str(media_name).lower().strip()
        for tier_name, tier_list in tiers_config.items():
            for configured_media in tier_list:
                if configured_media and configured_media in m_name:
                    if tier_name == 'tier1': return 10
                    if tier_name == 'tier2': return 8
                    if tier_name == 'tier3': return 5
        return 3

    def analyze_content_with_ai(self, content, key_message, project_desc, audience_mode, media_name):
        if not INTERNAL_API_KEY: return 0, 0, 0, "API Key Error: Key is missing"
        
        # 容错：默认值
        safe_km = key_message if key_message else "文章主题及核心观点"
        safe_desc = project_desc if project_desc else "一般性行业项目"

        prompt = f"""
        你是一个专业的公关传播分析师。请基于以下输入信息对一篇文章进行评分。
        
        【输入信息】
        1. 目标受众模式: {audience_mode}
        2. 媒体名称: {media_name}
        3. 核心传播信息 (Key Message): {safe_km}
        4. 项目描述: {safe_desc}
        5. 待分析文本: 
        {content[:3000]}... (内容截断)

        【任务】
        请分析并返回以下 3 个维度的分数（0-10分），并严格按照 JSON 格式返回：
        1. km_score: 文本是否有效传递了核心信息？如果是标题且包含关键词，也可给高分。(0=无, 10=深度)
        2. acquisition_score: 基于项目描述，这篇内容的获客吸引力如何？
        3. audience_precision_score: 考虑到媒体和受众模式，受众精准度如何？

        【输出格式】
        仅返回 JSON 字符串:
        {{"km_score": 8, "acquisition_score": 7, "audience_precision_score": 9}}
        """
        try:
            response = self.model.generate_content(prompt)
            clean_text = response.text.replace('```json', '').replace('```', '').strip()
            data = json.loads(clean_text)
            return (
                data.get('km_score', 0), 
                data.get('acquisition_score', 0), 
                data.get('audience_precision_score', 0), 
                "Success"
            )
        except Exception as e:
            return 0, 0, 0, f"AI Error: {str(e)}"

# --- 4. 侧边栏 (Sidebar) ---
with st.sidebar:
    st.header("⚙️ 系统配置")
    
    st.subheader("📋 项目基础信息")
    project_key_message = st.text_input("核心信息 (Key Message)", value="")
    project_desc = st.text_area("项目描述 (用于评估获客)", value="", height=100)
    audience_mode = st.radio("目标受众模式", ["大众 (General)", "患者 (Patient)", "医疗专业人士 (HCP)"])

    st.markdown("---")
    st.subheader("🏆 媒体分级配置")
    st.caption("输入媒体名称关键词，用逗号分隔")
    
    tier1_input = st.text_area("Tier 1 (10分)", value="", height=68)
    tier2_input = st.text_area("Tier 2 (8分)", value="", height=68)
    tier3_input = st.text_area("Tier 3 (5分)", value="", height=68)

    def parse_tiers(text):
        return [x.strip().lower() for x in text.split(',') if x.strip()]
    
    tier_config = {
        'tier1': parse_tiers(tier1_input),
        'tier2': parse_tiers(tier2_input),
        'tier3': parse_tiers(tier3_input)
    }

# --- 5. 主界面 (Main) ---

st.title("📡 传播价值 AI 评分系统")

# 顶部公式展示
with st.expander("查看核心算法公式", expanded=False):
    # 使用 \color{black} 确保公式颜色正确
    st.latex(r'''
    \color{black} \text{总分} = 0.5 \times \text{真需求} + 0.2 \times \text{获客效能} + 0.3 \times \text{声量}
    ''')
    st.latex(r'''
    \color{black} \text{真需求} = 0.6 \times \text{信息匹配} + 0.4 \times \text{受众精准度} 
    , \quad 
    \text{声量} = 0.6 \times \text{传播质量} + 0.4 \times \text{媒体分级}
    ''')

# 初始化引擎
engine = ScorerEngine()

# 标签页
tab1, tab2 = st.tabs(["📄 新闻稿评分", "📊 媒体报道评分"])

# --- TAB 1: 新闻稿评分 ---
with tab1:
    st.info("上传新闻稿 Word 文档，AI 将预判核心信息传递情况。")
    uploaded_word = st.file_uploader("上传 .docx 文件", type=['docx'])
    
    # 结果容器 (使用 session_state 防止刷新丢失)
    if 'word_analysis_result' not in st.session_state:
        st.session_state.word_analysis_result = None

    if uploaded_word:
        st.success(f"✅ 文档已就绪: {uploaded_word.name}")
        
        # 按钮改为浅色，文案改为“开始分析”
        if st.button("开始分析", key="btn_word_analyze"):
            if not project_key_message:
                st.warning("⚠️ 建议在左侧填写【核心信息】，否则 AI 评分可能不准确。")
            
            with st.spinner("AI 正在阅读文档..."):
                try:
                    # 读取内容 (增强版)
                    full_text = engine.read_docx_content(uploaded_word)
                    
                    if len(full_text.strip()) < 10:
                        st.error(f"文档内容过少 (提取到 {len(full_text)} 字)，无法进行分析。请检查文档是否加密或仅包含图片。")
                        st.session_state.word_analysis_result = None
                    else:
                        km, acq, prec, status = engine.analyze_content_with_ai(
                            full_text, project_key_message, project_desc, audience_mode, "内部稿件"
                        )
                        st.session_state.word_analysis_result = {
                            "km": km,
                            "status": status,
                            "text_len": len(full_text)
                        }
                except Exception as e:
                    st.error(f"解析错误: {e}")
    
    # 显示结果 (如果存在)
    if st.session_state.word_analysis_result:
        res = st.session_state.word_analysis_result
        st.divider()
        if res['km'] > 0:
            # 仅显示核心信息匹配度，移除了获客吸引力
            st.metric("核心信息匹配度", f"{res['km']}/10")
            st.progress(res['km']/10)
            st.success(f"分析成功！(基于 {res['text_len']} 字文本分析)")
        else:
            # 显性显示 0 分原因
            st.error(f"评分失败 (0分)。\n原因: {res['status']}")
            st.caption("提示: 请检查 API Key 额度，或文档是否包含有效文字。")

# --- TAB 2: 媒体报道评分 ---
with tab2:
    uploaded_csv = st.file_uploader("上传媒体监测报表 (.csv)", type=['csv'])

    if uploaded_csv:
        try:
            try:
                df = pd.read_csv(uploaded_csv)
            except UnicodeDecodeError:
                uploaded_csv.seek(0)
                df = pd.read_csv(uploaded_csv, encoding='gbk')
            except Exception as e:
                st.error(f"文件读取失败，请检查文件格式。错误信息: {e}")
                st.stop()

            # 清洗列名
            df.columns = df.columns.str.strip()

            # --- 核心适配逻辑 ---
            
            # 1. 映射媒体名称 (处理 "媒体" -> "媒体名称")
            if '媒体' in df.columns and '媒体名称' not in df.columns:
                df['媒体名称'] = df['媒体']
            
            # 2. 映射 URL (处理 "链接" -> "URL")
            if '链接' in df.columns and 'URL' not in df.columns:
                df['URL'] = df['链接']

            # 3. 处理浏览数据 (优先 PV，其次 浏览量)
            def to_num(x):
                try:
                    # 处理空值、逗号、加号、万字
                    if pd.isna(x) or x == '': return 0.0
                    s = str(x).replace(',', '').replace('+', '').strip()
                    if '万' in s:
                        return float(s.replace('万', '')) * 10000
                    return float(s)
                except:
                    return 0.0

            # 确保列存在
            if 'PV' not in df.columns: df['PV'] = 0
            if '浏览量' not in df.columns: df['浏览量'] = 0

            # 填充逻辑：创建 Clean_Views
            df['Clean_Views'] = df['PV'].apply(to_num)
            # 如果 PV 是 0，尝试用浏览量填充
            mask = df['Clean_Views'] == 0
            df.loc[mask, 'Clean_Views'] = df.loc[mask, '浏览量'].apply(to_num)
            
            # 最终用于显示的浏览量
            df['浏览量'] = df['Clean_Views']

            # 4. 计算互动量 (点赞+评论+转发)
            df['互动量'] = 0
            for col in ['点赞量', '评论量', '转发量']:
                if col in df.columns:
                    df['互动量'] += df[col].apply(to_num)

            # --- 检查列名 ---
            required_cols = ['媒体名称', 'URL', '互动量', '浏览量']
            missing_cols = [col for col in required_cols if col not in df.columns]
            
            if missing_cols:
                st.error(f"⚠️ CSV 缺少必要列，且无法自动映射。缺失: {missing_cols}")
                st.info(f"当前列: {list(df.columns)}")
                st.markdown("请确保 CSV 包含 `媒体`、`链接`、`PV`(或浏览量) 等列。")
            else:
                st.success(f"✅ 成功读取 {len(df)} 条数据，预览如下:")
                # 预览关键列
                preview_cols = ['媒体名称', '标题'] if '标题' in df.columns else ['媒体名称']
                preview_cols += ['URL', '浏览量', '互动量']
                st.dataframe(df[preview_cols].head(3), use_container_width=True)
                
                st.markdown("---")
                if st.button("🚀 点击开始 AI 全量评分", type="primary"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    results = []
                    total_rows = len(df)

                    for index, row in df.iterrows():
                        status_text.text(f"正在分析第 {index+1}/{total_rows} 条: {row['媒体名称']}...")
                        
                        # 1. 声量计算
                        vol_quality = engine.calculate_volume_quality(row['浏览量'], row['互动量'])
                        tier_score = engine.get_media_tier_score(row['媒体名称'], tier_config)
                        volume_total = 0.6 * vol_quality + 0.4 * tier_score
                        
                        # 2. 内容获取 (爬虫 + 标题兜底)
                        content = engine.fetch_url_content(row['URL'])
                        
                        # 兜底：如果爬不到内容，且 CSV 有标题，用标题分析
                        if not content and '标题' in df.columns and pd.notna(row['标题']):
                            content = f"文章标题：{row['标题']}"
                            msg_suffix = " (基于标题)"
                        else:
                            msg_suffix = ""

                        if content:
                            km_score, acq_score, prec_score, msg = engine.analyze_content_with_ai(
                                content, project_key_message, project_desc, audience_mode, row['媒体名称']
                            )
                            msg += msg_suffix
                        else:
                            km_score, acq_score, prec_score = 0, 0, 0
                            msg = "URL Fail & No Title"

                        # 3. 总分计算
                        true_demand = 0.6 * km_score + 0.4 * prec_score
                        total_score = (0.5 * true_demand) + (0.2 * acq_score) + (0.3 * volume_total)

                        results.append({
                            "媒体名称": row['媒体名称'],
                            "总分": round(total_score, 2),
                            "真需求": round(true_demand, 2),
                            "获客力": acq_score,
                            "声量": round(volume_total, 2),
                            "信息匹配": km_score,
                            "受众精准度": prec_score, 
                            "媒体分级": tier_score,
                            "状态": msg
                        })
                        progress_bar.progress((index + 1) / total_rows)

                    status_text.success("🎉 分析全部完成！")
                    res_df = pd.DataFrame(results)
                    
                    st.divider()
                    
                    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                    col_m1.metric("文章总数", len(res_df))
                    col_m2.metric("高价值 (≥8分)", len(res_df[res_df['总分'] >= 8]))
                    col_m3.metric("平均分", round(res_df['总分'].mean(), 2))
                    col_m4.metric("中位数", round(res_df['总分'].median(), 2))

                    col_chart1, col_chart2 = st.columns([2, 1])
                    with col_chart1:
                        st.subheader("📊 得分排行")
                        fig = px.bar(
                            res_df.sort_values('总分', ascending=True), 
                            x='总分', y='媒体名称', orientation='h',
                            color='总分', color_continuous_scale='Bluered'
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    
                    with col_chart2:
                        st.subheader("声量 vs 需求")
                        fig2 = px.scatter(
                            res_df, x='声量', y='真需求',
                            hover_name='媒体名称', size='总分', color='获客力'
                        )
                        st.plotly_chart(fig2, use_container_width=True)

                    st.subheader("📋 详细数据表")
                    st.dataframe(res_df.style.background_gradient(subset=['总分'], cmap='Greens'), use_container_width=True)

                    csv = res_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 导出结果 CSV", csv, "report.csv", "text/csv")

        except Exception as e:
            st.error(f"文件处理错误: {e}")
