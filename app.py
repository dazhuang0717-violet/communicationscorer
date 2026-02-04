import streamlit as st
import pandas as pd
import numpy as np
import requests
from bs4 import BeautifulSoup
import plotly.express as px
import plotly.graph_objects as go
from docx import Document
import io
import math
import json
import re
import time

st.set_page_config(
    page_title="肿瘤业务-传播价值 AI 评分系统",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
        [data-testid="stAppViewContainer"] { background-color: #ffffff !important; color: #31333F !important; }
        [data-testid="stSidebar"] { background-color: #f8f9fa !important; border-right: 1px solid #e0e0e0; }
        
        header[data-testid="stHeader"] { background-color: #ffffff !important; border-bottom: 1px solid #f0f2f6; }
        header[data-testid="stHeader"] button, header[data-testid="stHeader"] a, header[data-testid="stHeader"] svg { color: #31333F !important; fill: #31333F !important; }
        
        .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] { 
            color: #31333F !important; 
            background-color: #ffffff !important; 
            border: 1px solid #d1d5db; 
        }
        
        .stTextInput input:focus, .stTextArea textarea:focus { 
            border-color: #1E88E5 !important; 
            box-shadow: 0 0 0 1px #1E88E5 !important;
        }
        div[data-baseweb="select"] > div:focus-within {
            border-color: #1E88E5 !important;
        }

        button[kind="primary"] {
            background-color: #1E88E5 !important;
            border-color: #1E88E5 !important;
        }
        button[kind="secondary"] {
            border-color: #1E88E5 !important;
            color: #1E88E5 !important;
        }
        
        [data-testid="stFileUploaderDropzone"] { background-color: #f8f9fa !important; border: 1px dashed #d1d5db !important; }
        [data-testid="stFileUploaderDropzone"] div, [data-testid="stFileUploaderDropzone"] span, [data-testid="stFileUploaderDropzone"] p { color: #31333F !important; }
        
        [data-testid="stDataFrame"] { 
            color: #000000 !important; 
        }
        [data-testid="stDataFrame"] svg { fill: #31333F !important; }
        
        [data-testid="stDataFrame"] * {
            font-family: "Microsoft YaHei", "PingFang SC", "Source Sans Pro", sans-serif !important;
        }
        
        #MainMenu { visibility: hidden; }
        footer { visibility: hidden; }
        
        .stAlert { 
            background-color: #e3f2fd !important; 
            border: 1px solid #90caf9 !important; 
            color: #0d47a1 !important; 
        }

        .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] {
            border-bottom-color: #1E88E5 !important;
        }
        .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] p {
            color: #1E88E5 !important;
        }
        div[data-baseweb="tab-highlight"] {
            background-color: #1E88E5 !important;
        }
        
        .stTabs [data-baseweb="tab-list"] button:hover p {
            color: #1E88E5 !important;
        }
        .stTabs [data-baseweb="tab-list"] button:hover {
            color: #1E88E5 !important;
            border-bottom-color: #1E88E5 !important;
        }
    </style>
""", unsafe_allow_html=True)

class ScorerEngine:
    def __init__(self, key):
        self.api_key = key
        self.portkey_url = "https://api.portkey.ai/v1/chat/completions"

    def read_docx_content(self, file_obj):
        try:
            file_obj.seek(0)
            doc = Document(file_obj)
            full_text = []
            for para in doc.paragraphs:
                if para.text.strip(): full_text.append(para.text.strip())
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if para.text.strip(): full_text.append(para.text.strip())
            return "\n".join(full_text)
        except Exception as e:
            return f"Error: {str(e)}"

    def fetch_url_content(self, url):
        if not url or pd.isna(url): return ""
        if not str(url).startswith('http'): return ""
        try:
            jina_url = f"https://r.jina.ai/{url}"
            response = requests.get(jina_url, timeout=5)
            if response.status_code == 200 and len(response.text) > 50: return response.text[:10000]
        except: pass 
        try:
            headers = {'User-Agent': 'Mozilla/5.0'}
            response = requests.get(url, headers=headers, timeout=5)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')
                text = " ".join([p.get_text() for p in soup.find_all('p')])
                if len(text) > 50: return text[:10000]
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
        if not self.api_key or not str(self.api_key).strip(): 
            return 0, 0, 0, "API Key Missing", "请在侧边栏配置 Portkey Key"
        
        if not content or len(str(content).strip()) < 10:
             return 0, 0, 0, "内容过短/无效", "内容过短，无法生成评价"

        safe_km = key_message if key_message else "文章主题及核心观点"
        safe_desc = project_desc if project_desc else "一般性行业项目"

        prompt = f"""
        你是一个专业的公关传播分析师。请严格按照以下规则对内容进行评分（0-10分）：
        
        【输入信息】
        - 目标受众模式: {audience_mode}
        - 媒体名称: {media_name}
        - 核心传播信息 (Key Message): {safe_km}
        - 项目描述: {safe_desc}
        - 待分析文本: {content[:3000]}

        【输出任务】
        请直接返回一个标准的 JSON 对象，不要包含任何 Markdown 格式：
        {{
            "km_score": <数字>,
            "acquisition_score": <数字>,
            "audience_precision_score": <数字>,
            "comment": "100字以内的客观评价"
        }}
        """
        
        candidate_models = [
            'gemini-1.5-flash',
            'gemini-1.5-flash-latest',
            'gemini-2.0-flash',
            'gemini-1.5-pro',
            'gemini-pro'
        ]
        
        # 自动识别 Key 类型
        headers = {
            "x-portkey-api-key": self.api_key,
            "x-portkey-provider": "google",
            "Content-Type": "application/json"
        }
        
        # 如果用户输入的是虚拟密钥（通常不以 pk- 开头或具有特定特征）
        if not str(self.api_key).startswith("pk-"):
            headers["x-portkey-virtual-key"] = self.api_key

        def extract_json(text):
            try: return json.loads(text)
            except: pass
            try:
                clean = text.replace('```json', '').replace('```', '').strip()
                return json.loads(clean)
            except: pass
            try:
                match = re.search(r'\{.*\}', text, re.DOTALL)
                if match: return json.loads(match.group(0))
            except: pass
            return None

        last_error = ""
        for model_name in candidate_models:
            try:
                payload = {
                    "model": model_name,
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.2
                }
                response = requests.post(self.portkey_url, headers=headers, json=payload, timeout=20)
                
                if response.status_code == 200:
                    res_json = response.json()
                    res_text = res_json['choices'][0]['message']['content']
                    data = extract_json(res_text)
                    if data:
                        return (
                            float(data.get('km_score', 0)), 
                            float(data.get('acquisition_score', 0)), 
                            float(data.get('audience_precision_score', 0)), 
                            "Success",
                            data.get('comment', 'AI 已返回评分')
                        )
                else:
                    err_info = response.json().get('error', {}).get('message', response.text)
                    last_error = f"Model {model_name}: {err_info}"
                    if response.status_code == 412: # 权限限制，尝试下一个模型
                        continue
                    elif response.status_code == 401:
                        return 0, 0, 0, "Auth Failed", "API Key 错误或失效"
            except Exception as e:
                last_error = str(e)
                continue

        return 0, 0, 0, f"Error: {last_error}", "AI 评分失败，请检查 Portkey 后台权限设置"

def generate_html_report(project_name, metrics, charts, df_top):
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>{project_name} - 评分报告</title>
        <style>
            body {{ font-family: "Microsoft YaHei", sans-serif; padding: 40px; color: #333; }}
            h1 {{ color: #1E88E5; border-bottom: 2px solid #1E88E5; padding-bottom: 10px; }}
            h2 {{ color: #1E88E5; margin-top: 30px; }}
            .metrics-container {{ display: flex; justify-content: space-between; margin-bottom: 30px; background: #f8f9fa; padding: 20px; border-radius: 8px; }}
            .metric-box {{ text-align: center; }}
            .metric-val {{ font-size: 24px; font-weight: bold; color: #1E88E5; }}
            .metric-lbl {{ font-size: 14px; color: #666; }}
            .chart-container {{ margin-bottom: 40px; page-break-inside: avoid; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #1E88E5; color: white; }}
            tr:nth-child(even) {{ background-color: #f2f2f2; }}
            @media print {{
                .no-print {{ display: none; }}
                body {{ padding: 0; }}
            }}
        </style>
    </head>
    <body>
        <h1>📈 项目评分报告: {project_name}</h1>
        
        <div class="metrics-container">
            <div class="metric-box"><div class="metric-val">{metrics['total']:.2f}</div><div class="metric-lbl">项目总分</div></div>
            <div class="metric-box"><div class="metric-val">{metrics['demand']:.2f}</div><div class="metric-lbl">真需求</div></div>
            <div class="metric-box"><div class="metric-val">{metrics['acquisition']:.2f}</div><div class="metric-lbl">获客效能</div></div>
            <div class="metric-box"><div class="metric-val">{metrics['volume']:.2f}</div><div class="metric-lbl">声量</div></div>
        </div>

        <h2>📊 数据洞察</h2>
        <div style="display: flex; flex-wrap: wrap;">
            <div style="width: 50%; min-width: 300px;" class="chart-container">
                <h3>项目能力雷达</h3>
                {charts['radar']}
            </div>
            <div style="width: 50%; min-width: 300px;" class="chart-container">
                <h3>传播价值矩阵</h3>
                {charts['scatter']}
            </div>
        </div>
        <div class="chart-container">
            <h3>媒体贡献 TOP 榜单</h3>
            {charts['bar']}
        </div>

        <h2>🏆 详细数据 (Top 10)</h2>
        {df_top.to_html(index=False)}
    </body>
    </html>
    """
    return html_content

with st.sidebar:
    st.header("⚙️ 系统配置")
    
    api_key = st.text_input("🔑 Portkey Key", value="", type="password", help="可以是 Portkey API Key 或 Virtual Key")

    st.subheader("📋 项目基础信息")
    project_name = st.text_input("项目名称", placeholder="请输入项目名")
    project_key_message = st.text_input("核心传播信息", placeholder="Key Message")
    project_desc = st.text_area("项目描述", placeholder="用于分析获客效能", height=100)
    audience_mode = st.radio("目标受众模式", ["大众 (General)", "患者 (Patient)", "医疗专业人士 (HCP)"])

    st.markdown("---")
    st.subheader("🏆 媒体分级")
    tier1_input = st.text_area("Tier 1 (10分)", placeholder="用英文逗号分隔")
    tier2_input = st.text_area("Tier 2 (8分)", placeholder="用英文逗号分隔")
    tier3_input = st.text_area("Tier 3 (5分)", placeholder="用英文逗号分隔")

    def parse_tiers(text):
        return [x.strip().lower() for x in text.split(',') if x.strip()]
    
    tier_config = {
        'tier1': parse_tiers(tier1_input),
        'tier2': parse_tiers(tier2_input),
        'tier3': parse_tiers(tier3_input)
    }

engine = ScorerEngine(api_key)

st.title("📡 肿瘤业务-传播价值 AI 评分系统")

tab1, tab2, tab3 = st.tabs(["📄 新闻稿评分", "📊 媒体报道评分", "📈 项目总揽"])

with tab1:
    st.info("📄 上传新闻稿 .docx 文档，AI 将评估信息匹配度。")
    uploaded_word = st.file_uploader("上传文档", type=['docx'])
    
    if uploaded_word:
        if st.button("开始 AI 分析", key="btn_word"):
            with st.spinner("AI 正在分析..."):
                full_text = engine.read_docx_content(uploaded_word)
                km, acq, prec, status, comment = engine.analyze_content_with_ai(
                    full_text, project_key_message, project_desc, audience_mode, "内部新闻稿"
                )
                if status == "Success":
                    st.metric("核心信息匹配得分", f"{km}/10")
                    st.progress(km/10)
                    st.success(f"评价：{comment}")
                else:
                    st.error(status)

with tab2:
    st.info("📊 上传媒体监测 Excel/CSV。")
    uploaded_file = st.file_uploader("上传报表", type=['xlsx', 'csv'])

    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            st.dataframe(df.head(5), use_container_width=True)
            
            if st.button("批量评分", key="btn_batch"):
                progress_bar = st.progress(0)
                results = []
                for i, row in df.iterrows():
                    media = row.get('媒体名称', row.get('媒体', '未知媒体'))
                    url = row.get('URL', row.get('链接', ''))
                    views = row.get('浏览量', 0)
                    ints = row.get('互动量', 0)
                    
                    vol_score = engine.calculate_volume_quality(views, ints)
                    tier_score = engine.get_media_tier_score(media, tier_config)
                    
                    # 获取正文或标题进行 AI 分析
                    content = str(row.get('正文', row.get('标题', media)))
                    km, acq, prec, status, comment = engine.analyze_content_with_ai(
                        content, project_key_message, project_desc, audience_mode, media
                    )
                    
                    true_demand = 0.6 * km + 0.4 * prec
                    vol_total = 0.6 * vol_score + 0.4 * tier_score
                    total = 0.5 * true_demand + 0.2 * acq + 0.3 * vol_total
                    
                    results.append({
                        "媒体名称": media,
                        "项目总分": round(total, 2),
                        "真需求": round(true_demand, 2),
                        "获客效能": acq,
                        "声量表现": round(vol_total, 2),
                        "AI 状态": status
                    })
                    progress_bar.progress((i+1)/len(df))
                
                st.session_state.res_df = pd.DataFrame(results)
                st.success("批量分析完成")
                st.dataframe(st.session_state.res_df)
        except Exception as e:
            st.error(f"处理失败: {e}")

with tab3:
    if 'res_df' in st.session_state:
        rdf = st.session_state.res_df
        st.subheader("项目整体表现")
        c1, c2, c3 = st.columns(3)
        c1.metric("平均总分", round(rdf['项目总分'].mean(), 2))
        c2.metric("需求覆盖", round(rdf['真需求'].mean(), 2))
        c3.metric("声量表现", round(rdf['声量表现'].mean(), 2))
        
        fig = px.scatter(rdf, x="声量表现", y="真需求", size="项目总分", color="项目总分", hover_name="媒体名称")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("请先在 Tab 2 完成数据评分")
