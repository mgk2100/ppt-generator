#!/usr/bin/env python3
"""배터리 3사 AI 활용 현황 분석 보고서 생성기"""

import base64
import os

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
OUTPUT = os.path.join(os.path.dirname(__file__), "report.html")


def load_image_b64(filename):
    path = os.path.join(DATA_DIR, filename)
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


def generate():
    img_lgcns = load_image_b64("lg cns.png")
    img_samsung = load_image_b64("samsung sdi.png")
    img_skon = load_image_b64("sk on.png")

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>국내 배터리 3사 AI 활용 현황 분석</title>
<link rel="preconnect" href="https://cdn.jsdelivr.net">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<style>
/* ===== RESET & BASE ===== */
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
html {{ scroll-behavior: smooth; font-size: 16px; }}
body {{
    font-family: 'Pretendard', 'Noto Sans KR', -apple-system, BlinkMacSystemFont, sans-serif;
    color: #222222;
    background: #FFFFFF;
    line-height: 1.7;
}}

/* ===== NAV ===== */
nav {{
    position: sticky; top: 0; z-index: 100;
    background: rgba(255,255,255,0.95);
    backdrop-filter: blur(10px);
    border-bottom: 1px solid #E0E0E0;
    padding: 0 2rem;
}}
nav ul {{
    max-width: 1200px; margin: 0 auto;
    display: flex; gap: 0; list-style: none;
    overflow-x: auto;
}}
nav li a {{
    display: block; padding: 0.9rem 1.1rem;
    text-decoration: none; color: #555;
    font-size: 0.88rem; font-weight: 500;
    white-space: nowrap;
    border-bottom: 3px solid transparent;
    transition: all 0.2s;
}}
nav li a:hover, nav li a.active {{
    color: #003478;
    border-bottom-color: #003478;
}}

/* ===== LAYOUT ===== */
.container {{ max-width: 1200px; margin: 0 auto; padding: 0 1.5rem; }}
section {{ padding: 3rem 0; }}
.section-divider {{ border: none; border-top: 2px solid #E8E8E8; margin: 0; }}

/* ===== COVER ===== */
.cover {{
    background: linear-gradient(135deg, #003478 0%, #00509E 40%, #0066CC 100%);
    color: #fff;
    padding: 5rem 0;
    text-align: center;
    position: relative;
    overflow: hidden;
}}
.cover::before {{
    content: '';
    position: absolute; inset: 0;
    background: radial-gradient(circle at 20% 80%, rgba(0,52,120,0.3) 0%, transparent 50%),
                radial-gradient(circle at 80% 20%, rgba(0,100,200,0.2) 0%, transparent 50%);
}}
.cover * {{ position: relative; z-index: 1; }}
.cover h1 {{
    font-size: 2.6rem; font-weight: 800;
    margin-bottom: 1rem;
    letter-spacing: -0.02em;
}}
.cover .subtitle {{
    font-size: 1.2rem; font-weight: 400;
    color: rgba(255,255,255,0.8);
    margin-bottom: 0.5rem;
}}
.cover .meta {{
    font-size: 0.9rem;
    color: rgba(255,255,255,0.6);
    margin-top: 1.5rem;
}}
.cover .badge-row {{
    display: flex; justify-content: center; gap: 1rem;
    margin-top: 2rem; flex-wrap: wrap;
}}
.cover .badge {{
    display: inline-block;
    padding: 0.4rem 1.2rem;
    border-radius: 20px;
    font-size: 0.85rem; font-weight: 600;
}}
.badge-lg {{ background: rgba(165,0,52,0.7); }}
.badge-ss {{ background: rgba(20,40,160,0.8); }}
.badge-sk {{ background: rgba(247,181,0,0.7); color: #333; }}

/* ===== HEADINGS ===== */
h2 {{
    font-size: 1.8rem; font-weight: 800;
    margin-bottom: 0.5rem;
    color: #003478;
    letter-spacing: -0.01em;
}}
h2 .accent {{ color: #003478; }}
h3 {{
    font-size: 1.3rem; font-weight: 700;
    margin-bottom: 0.8rem;
    color: #333;
}}
.section-desc {{
    color: #666; font-size: 1rem;
    margin-bottom: 2rem;
    max-width: 700px;
}}

/* ===== CARDS ===== */
.card {{
    background: #fff;
    border-radius: 16px;
    padding: 2rem;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    transition: box-shadow 0.2s;
}}
.card:hover {{ box-shadow: 0 4px 24px rgba(0,0,0,0.1); }}

/* ===== SUMMARY CARDS ===== */
.summary-grid {{
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1.5rem;
    margin-bottom: 2.5rem;
}}
.summary-card {{
    border-radius: 16px;
    padding: 2rem;
    color: #fff;
    position: relative;
    overflow: hidden;
}}
.summary-card::after {{
    content: '';
    position: absolute; right: -20px; top: -20px;
    width: 100px; height: 100px;
    border-radius: 50%;
    background: rgba(255,255,255,0.1);
}}
.summary-card.lg {{ background: linear-gradient(135deg, #A50034, #c41048); }}
.summary-card.ss {{ background: linear-gradient(135deg, #1428A0, #1e3cc0); }}
.summary-card.sk {{ background: linear-gradient(135deg, #E5A800, #F7B500); color: #333; }}
.summary-card .co-name {{
    font-size: 0.85rem; font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    opacity: 0.85;
    margin-bottom: 0.5rem;
}}
.summary-card .product {{
    font-size: 1.3rem; font-weight: 800;
    margin-bottom: 0.8rem;
}}
.summary-card .desc {{
    font-size: 0.92rem;
    line-height: 1.6;
    opacity: 0.9;
}}

/* ===== WHY BOX ===== */
.why-box {{
    background: linear-gradient(135deg, #f5f8fc, #eaf0f8);
    border-left: 5px solid #003478;
    border-radius: 0 16px 16px 0;
    padding: 2rem 2.5rem;
}}
.why-box h3 {{ color: #003478; margin-bottom: 1rem; }}
.why-list {{
    list-style: none; display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 1rem;
}}
.why-list li {{
    display: flex; align-items: flex-start; gap: 0.7rem;
    font-size: 0.95rem;
}}
.why-icon {{
    flex-shrink: 0;
    width: 36px; height: 36px;
    border-radius: 10px;
    background: #003478;
    color: #fff;
    display: flex; align-items: center; justify-content: center;
    font-size: 1.1rem;
}}

/* ===== AI EXPLAINER ===== */
.ai-box {{
    background: #fff;
    border: 2px solid #E8E8E8;
    border-radius: 16px;
    padding: 2rem 2.5rem;
    margin-top: 2rem;
}}
.ai-box h3 {{ color: #003478; }}
.ai-analogy {{
    background: #f0f8ff;
    border-radius: 12px;
    padding: 1.5rem;
    margin: 1rem 0;
    font-size: 1.05rem;
    font-style: italic;
    color: #333;
    border-left: 4px solid #3498db;
}}
.term-grid {{
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1rem;
    margin-top: 1.5rem;
}}
.term-card {{
    background: #fafafa;
    border-radius: 12px;
    padding: 1.2rem;
    text-align: center;
}}
.term-card .term-icon {{ font-size: 2rem; margin-bottom: 0.5rem; }}
.term-card .term-name {{ font-weight: 700; color: #1a1a2e; margin-bottom: 0.3rem; }}
.term-card .term-desc {{ font-size: 0.85rem; color: #666; }}

/* ===== COMPANY DEEP DIVE ===== */
.company-section {{
    border-radius: 20px;
    padding: 3rem;
    margin-bottom: 0;
    background: #fff;
    box-shadow: 0 2px 12px rgba(0,0,0,0.05);
}}
.company-header {{
    display: flex;
    align-items: center;
    gap: 1.5rem;
    margin-bottom: 2rem;
    padding-bottom: 1.5rem;
    border-bottom: 2px solid #f0f0f0;
}}
.company-logo {{
    width: 60px; height: 60px;
    border-radius: 14px;
    display: flex; align-items: center; justify-content: center;
    font-weight: 900; font-size: 0.75rem; color: #fff;
    flex-shrink: 0;
}}
.company-logo.lg {{ background: #A50034; }}
.company-logo.ss {{ background: #1428A0; }}
.company-logo.sk {{ background: #F7B500; color: #333; }}
.company-header h3 {{
    font-size: 1.5rem; margin-bottom: 0.2rem;
}}
.company-header .tagline {{
    color: #888; font-size: 0.95rem;
}}

.info-grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 1.5rem;
    margin-bottom: 2rem;
}}
.info-block {{
    background: #fafafa;
    border-radius: 12px;
    padding: 1.5rem;
}}
.info-block h4 {{
    font-size: 0.85rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: #888;
    margin-bottom: 0.8rem;
}}
.info-block ul {{
    list-style: none;
    padding: 0;
}}
.info-block li {{
    padding: 0.3rem 0;
    font-size: 0.95rem;
    display: flex; align-items: flex-start; gap: 0.5rem;
}}
.info-block li::before {{
    content: '\\25B8';
    color: #aaa;
    flex-shrink: 0;
    margin-top: 2px;
}}

/* Before/After */
.ba-grid {{
    display: grid;
    grid-template-columns: 1fr auto 1fr;
    gap: 0;
    margin: 1.5rem 0;
    border-radius: 12px;
    overflow: hidden;
}}
.ba-col {{
    padding: 1.5rem;
}}
.ba-before {{
    background: #fff5f5;
}}
.ba-arrow {{
    background: #f0f0f0;
    display: flex; align-items: center; justify-content: center;
    font-size: 1.5rem; color: #888;
    padding: 0 1rem;
}}
.ba-after {{
    background: #f0fff4;
}}
.ba-col h4 {{
    font-size: 0.85rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 0.8rem;
}}
.ba-before h4 {{ color: #c0392b; }}
.ba-after h4 {{ color: #27ae60; }}
.ba-col ul {{ list-style: none; }}
.ba-col li {{ padding: 0.25rem 0; font-size: 0.9rem; }}

/* Flowchart */
.flow {{
    display: flex;
    align-items: center;
    gap: 0;
    flex-wrap: wrap;
    justify-content: center;
    margin: 1.5rem 0;
    padding: 1.5rem;
    background: #fafafa;
    border-radius: 12px;
}}
.flow-step {{
    padding: 0.6rem 1.1rem;
    border-radius: 10px;
    font-size: 0.82rem;
    font-weight: 600;
    text-align: center;
    min-width: 90px;
    white-space: nowrap;
}}
.flow-arrow {{
    font-size: 1.2rem;
    color: #aaa;
    padding: 0 0.3rem;
}}

.flow-human {{ background: #e8f4fd; color: #1a5276; border: 2px solid #85c1e9; }}
.flow-ai {{ border: 2px solid; }}
.flow-ai.lg {{ background: #fce4ec; color: #A50034; border-color: #A50034; }}
.flow-ai.ss {{ background: #e8eaf6; color: #1428A0; border-color: #1428A0; }}
.flow-ai.sk {{ background: #fff8e1; color: #7a6300; border-color: #F7B500; }}
.flow-result {{ background: #e8f5e9; color: #1b5e20; border: 2px solid #66bb6a; }}

/* KPI */
.kpi-row {{
    display: flex; gap: 1rem;
    margin: 1.5rem 0;
}}
.kpi {{
    flex: 1;
    text-align: center;
    background: #fafafa;
    border-radius: 12px;
    padding: 1.5rem 1rem;
}}
.kpi .number {{
    font-size: 2.2rem;
    font-weight: 900;
    line-height: 1;
    margin-bottom: 0.3rem;
}}
.kpi .unit {{ font-size: 0.85rem; color: #888; }}
.kpi.lg .number {{ color: #A50034; }}
.kpi.ss .number {{ color: #1428A0; }}
.kpi.sk .number {{ color: #c79100; }}

/* Booth photo */
.booth-photo {{
    margin-top: 1.5rem;
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid #eee;
}}
.booth-photo img {{
    width: 100%;
    display: block;
}}
.booth-photo .caption {{
    padding: 0.8rem 1rem;
    background: #fafafa;
    font-size: 0.85rem;
    color: #888;
    text-align: center;
}}

/* Customers */
.customer-tags {{
    display: flex; gap: 0.5rem; flex-wrap: wrap;
    margin-top: 0.5rem;
}}
.customer-tag {{
    padding: 0.3rem 0.8rem;
    border-radius: 20px;
    font-size: 0.82rem;
    font-weight: 500;
    background: #f0f0f0;
    color: #555;
}}

/* ===== COMPARISON ===== */
.comp-table {{
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    margin: 1.5rem 0;
    font-size: 0.9rem;
}}
.comp-table th {{
    background: #003478;
    color: #fff;
    padding: 1rem;
    font-weight: 600;
    text-align: center;
}}
.comp-table th:first-child {{ text-align: left; }}
.comp-table td {{
    padding: 0.8rem 1rem;
    border-bottom: 1px solid #f0f0f0;
    text-align: center;
    vertical-align: top;
}}
.comp-table td:first-child {{
    font-weight: 600;
    text-align: left;
    background: #fafafa;
    min-width: 140px;
}}
.comp-table tr:last-child td {{ border-bottom: none; }}
.comp-table tr:hover td {{ background: #f5f8fc; }}
.comp-table tr:hover td:first-child {{ background: #eaf0f8; }}
.check {{ color: #27ae60; font-weight: 700; }}
.partial {{ color: #f39c12; }}
.dash {{ color: #ccc; }}

/* Chart container */
.chart-container {{
    max-width: 500px;
    margin: 2rem auto;
}}

/* Venn */
.venn-container {{
    display: flex;
    justify-content: center;
    align-items: center;
    padding: 2rem;
}}
.venn {{
    position: relative;
    width: 450px;
    height: 320px;
}}
.venn-circle {{
    position: absolute;
    width: 200px; height: 200px;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 0.8rem;
    text-align: center;
    padding: 1rem;
}}
.venn-lg {{
    background: rgba(165,0,52,0.15);
    border: 2px solid rgba(165,0,52,0.4);
    left: 30px; top: 30px;
    color: #A50034;
}}
.venn-ss {{
    background: rgba(20,40,160,0.15);
    border: 2px solid rgba(20,40,160,0.4);
    left: 150px; top: 30px;
    color: #1428A0;
}}
.venn-sk {{
    background: rgba(247,181,0,0.15);
    border: 2px solid rgba(247,181,0,0.4);
    left: 90px; top: 120px;
    color: #7a6300;
}}
.venn-center {{
    position: absolute;
    left: 140px; top: 100px;
    font-size: 0.75rem;
    font-weight: 800;
    color: #333;
    text-align: center;
    z-index: 10;
    line-height: 1.3;
}}

/* ===== INSIGHTS ===== */
.insight-cards {{
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1.5rem;
    margin: 2rem 0;
}}
.insight-card {{
    background: #fff;
    border-radius: 16px;
    padding: 2rem;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    border-top: 4px solid;
}}
.insight-card:nth-child(1) {{ border-color: #A50034; }}
.insight-card:nth-child(2) {{ border-color: #1428A0; }}
.insight-card:nth-child(3) {{ border-color: #F7B500; }}
.insight-card .insight-num {{
    font-size: 2rem; font-weight: 900;
    color: #ddd;
    margin-bottom: 0.5rem;
}}
.insight-card h4 {{
    font-size: 1.05rem; font-weight: 700;
    margin-bottom: 0.8rem;
    color: #003478;
}}
.insight-card p {{
    font-size: 0.92rem;
    color: #555;
    line-height: 1.6;
}}

/* SL Implications */
.sl-box {{
    background: linear-gradient(135deg, #003478, #00509E);
    color: #fff;
    border-radius: 16px;
    padding: 2.5rem;
    margin-top: 2rem;
}}
.sl-box h3 {{ color: #fff; margin-bottom: 1rem; }}
.sl-list {{ list-style: none; }}
.sl-list li {{
    padding: 0.6rem 0;
    font-size: 0.95rem;
    display: flex; align-items: flex-start; gap: 0.8rem;
    line-height: 1.6;
}}
.sl-list .sl-num {{
    flex-shrink: 0;
    width: 28px; height: 28px;
    border-radius: 50%;
    background: rgba(255,255,255,0.15);
    display: flex; align-items: center; justify-content: center;
    font-size: 0.8rem; font-weight: 700;
}}

/* ===== GLOSSARY ===== */
.glossary {{
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 1rem;
    margin: 1.5rem 0;
}}
.glossary-item {{
    padding: 1rem;
    background: #fafafa;
    border-radius: 10px;
}}
.glossary-item dt {{
    font-weight: 700;
    color: #003478;
    margin-bottom: 0.3rem;
}}
.glossary-item dd {{
    font-size: 0.9rem;
    color: #666;
}}

/* ===== CONFERENCE BOX ===== */
.conf-box {{
    background: #fff;
    border-radius: 16px;
    padding: 2rem;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    margin-bottom: 2rem;
}}
.conf-stats {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 1rem;
    margin: 1.5rem 0;
}}
.conf-stat {{
    text-align: center;
    padding: 1rem;
    background: #f5f8fc;
    border-radius: 12px;
}}
.conf-stat .num {{
    font-size: 1.8rem;
    font-weight: 900;
    color: #003478;
}}
.conf-stat .label {{
    font-size: 0.82rem;
    color: #888;
    margin-top: 0.3rem;
}}

/* ===== FOOTER ===== */
footer {{
    background: #003478;
    color: rgba(255,255,255,0.6);
    padding: 2rem 0;
    text-align: center;
    font-size: 0.85rem;
}}
footer a {{ color: rgba(255,255,255,0.8); }}

/* ===== PRINT ===== */
@media print {{
    nav {{ display: none; }}
    body {{ background: #fff; }}
    .cover {{ padding: 2rem 0; }}
    section {{ padding: 1.5rem 0; }}
    .card, .company-section {{ box-shadow: none; border: 1px solid #ddd; page-break-inside: avoid; }}
    .summary-grid, .info-grid, .insight-cards {{ grid-template-columns: 1fr; }}
    .chart-container {{ max-width: 400px; }}
    .ba-grid {{ grid-template-columns: 1fr; }}
    .ba-arrow {{ display: none; }}
}}

/* ===== RESPONSIVE ===== */
@media (max-width: 900px) {{
    .summary-grid, .insight-cards {{ grid-template-columns: 1fr; }}
    .info-grid {{ grid-template-columns: 1fr; }}
    .why-list {{ grid-template-columns: 1fr; }}
    .term-grid {{ grid-template-columns: 1fr; }}
    .conf-stats {{ grid-template-columns: repeat(2, 1fr); }}
    .glossary {{ grid-template-columns: 1fr; }}
    .ba-grid {{ grid-template-columns: 1fr; }}
    .ba-arrow {{ transform: rotate(90deg); padding: 0.5rem 0; }}
    h2 {{ font-size: 1.4rem; }}
    .cover h1 {{ font-size: 1.8rem; }}
    .company-section {{ padding: 1.5rem; }}
}}
</style>
</head>
<body>

<!-- ===== NAVIGATION ===== -->
<nav>
<ul>
    <li><a href="#summary">요약</a></li>
    <li><a href="#context">산업 배경</a></li>
    <li><a href="#lgcns">LG CNS</a></li>
    <li><a href="#samsung">Samsung SDI</a></li>
    <li><a href="#skon">SK ON</a></li>
    <li><a href="#comparison">비교 분석</a></li>
    <li><a href="#insights">시사점</a></li>
    <li><a href="#appendix">부록</a></li>
</ul>
</nav>

<!-- ===== COVER ===== -->
<header class="cover">
<div class="container">
    <h1>국내 배터리 3사의 AI 활용 현황 분석</h1>
    <p class="subtitle">인터배터리 2026 전시회 기반 조사 보고서</p>
    <p class="subtitle">LG CNS &middot; Samsung SDI &middot; SK ON</p>
    <p class="meta">2026년 3월 &middot; SL</p>
    <div class="badge-row">
        <span class="badge badge-lg">LG CNS &mdash; 실험 자동화</span>
        <span class="badge badge-ss">Samsung SDI &mdash; 배터리 진단</span>
        <span class="badge badge-sk">SK ON &mdash; AI 연구원</span>
    </div>
</div>
</header>

<!-- ===== EXECUTIVE SUMMARY ===== -->
<section id="summary">
<div class="container">
    <h2><span class="accent">01</span> 핵심 요약</h2>
    <p class="section-desc">배터리 3사가 AI를 활용하여 만들어내는 핵심 결과물을 한눈에 정리했습니다.</p>

    <div class="summary-grid">
        <div class="summary-card lg">
            <div class="co-name">LG CNS</div>
            <div class="product">AI 탑재 장비 3종</div>
            <div class="desc">배터리 시험·활성화·폐배터리 방전 장비에 AI를 탑재하여 전 과정을 자동화합니다. 에이테크놀로지 공동개발, 장비 20대 납품.</div>
        </div>
        <div class="summary-card ss">
            <div class="co-name">Samsung SDI</div>
            <div class="product">SBI</div>
            <div class="desc">AI가 전 세계 1,400개 배터리 현장을 24시간 감시하며, 이상 징후를 미리 감지하여 사고를 예방합니다. 2026년 10월 상용화.</div>
        </div>
        <div class="summary-card sk">
            <div class="co-name">SK ON</div>
            <div class="product">ADAM</div>
            <div class="desc">AI 연구원이 배터리 설계부터 성능 예측, 원가 산출까지 자동 수행합니다. 설계 기간 1/3, 원가 분석 700배 향상.</div>
        </div>
    </div>

    <div class="why-box">
        <h3>왜 배터리 업계가 AI를 도입하고 있는가?</h3>
        <ul class="why-list">
            <li>
                <div class="why-icon">&#9201;</div>
                <div><strong>개발 기간 압박</strong><br>전기차 시장의 폭발적 성장으로 신제품 출시 주기가 크게 단축되고 있습니다.</div>
            </li>
            <li>
                <div class="why-icon">&#9888;</div>
                <div><strong>안전성 강화</strong><br>ESS 화재 등 안전 사고 예방에 대한 사회적 요구가 높아지고 있습니다.</div>
            </li>
            <li>
                <div class="why-icon">&#128202;</div>
                <div><strong>데이터 폭증</strong><br>제조 공정과 운영 현장에서 발생하는 데이터가 사람이 처리할 수 없는 수준으로 증가했습니다.</div>
            </li>
            <li>
                <div class="why-icon">&#127760;</div>
                <div><strong>글로벌 경쟁</strong><br>중국 CATL(세계 1위 배터리 기업, 점유율 39%)은 5,000만 건 이상의 데이터로 AI를 운영 중입니다.</div>
            </li>
        </ul>
    </div>
</div>
</section>

<hr class="section-divider">

<!-- ===== INDUSTRY CONTEXT ===== -->
<section id="context">
<div class="container">
    <h2><span class="accent">02</span> 산업 배경</h2>
    <p class="section-desc">배터리 산업에서 AI가 왜 필요해졌는지, 그리고 이번 조사의 배경을 설명합니다.</p>

    <div class="conf-box">
        <h3>인터배터리 2026 전시회</h3>
        <p style="color:#666; margin-bottom:1rem;">2026년 3월 11~13일, 서울 코엑스에서 개최된 국내 최대 배터리 산업 전시회입니다.<br>올해의 핵심 키워드는 <strong>AI, 로봇, ESS(에너지저장장치)</strong>였습니다.</p>
        <div class="conf-stats">
            <div class="conf-stat">
                <div class="num">667</div>
                <div class="label">참가 기업</div>
            </div>
            <div class="conf-stat">
                <div class="num">14</div>
                <div class="label">참가 국가</div>
            </div>
            <div class="conf-stat">
                <div class="num">7.7만</div>
                <div class="label">총 참관객</div>
            </div>
            <div class="conf-stat">
                <div class="num">3</div>
                <div class="label">AI 핵심 전시 기업</div>
            </div>
        </div>
    </div>

    <div class="ai-box">
        <h3>AI(인공지능)란 무엇인가?</h3>
        <div class="ai-analogy">
            "AI는 <strong>잠들지 않는 경험 많은 엔지니어</strong>와 같습니다.<br>
            수백만 개의 데이터를 동시에 읽고, 과거의 실수에서 배우며, 시간이 지날수록 더 정확해집니다."
        </div>
        <p style="color:#666; margin: 1rem 0;">이 보고서에서 자주 등장하는 AI 관련 용어 3가지를 쉽게 풀어드립니다.</p>
        <div class="term-grid">
            <div class="term-card">
                <div class="term-icon">&#129504;</div>
                <div class="term-name">머신러닝</div>
                <div class="term-desc">컴퓨터가 데이터에서 스스로 규칙과 패턴을 찾아내는 기술.<br>사람이 일일이 규칙을 알려주지 않아도 됩니다.</div>
            </div>
            <div class="term-card">
                <div class="term-icon">&#128218;</div>
                <div class="term-name">빅데이터</div>
                <div class="term-desc">사람이 수작업으로 처리하기 어려울 만큼 방대한 양의 데이터.<br>AI의 학습 재료가 됩니다.</div>
            </div>
            <div class="term-card">
                <div class="term-icon">&#128225;</div>
                <div class="term-name">실시간 모니터링</div>
                <div class="term-desc">데이터를 수집하자마자 즉시 분석하는 것.<br>이상이 생기면 바로 알려줍니다.</div>
            </div>
        </div>
    </div>
</div>
</section>

<hr class="section-divider">

<!-- ===== LG CNS ===== -->
<section id="lgcns">
<div class="container">
    <h2><span class="accent">03-1</span> LG CNS</h2>
    <p class="section-desc">AI로 배터리 실험을 자동화하다</p>

    <div class="company-section">
        <div class="company-header">
            <div class="company-logo lg">LG<br>CNS</div>
            <div>
                <h3>AI 배터리 장비 솔루션</h3>
                <div class="tagline">에이전틱 AI 기반 &mdash; R&amp;D 시험부터 폐배터리 처리까지 전 과정 자동화</div>
            </div>
        </div>

        <div class="info-grid">
            <div class="info-block">
                <h4>AI가 하는 일</h4>
                <ul>
                    <li>충방전 시험 장비에 AI를 탑재하여 실험 설계·분석·보고서를 자동화</li>
                    <li>배터리 활성화 장비에 AI를 탑재하여 공정 시간을 단축</li>
                    <li>폐배터리 방전 장비에 AI를 탑재하여 안전한 방전 조건을 자동 계산</li>
                    <li>1ms(0.001초) 단위의 정밀 데이터를 AI가 실시간 분석</li>
                    <li>LG에너지솔루션 오창 공장에 장비 20대 납품 완료</li>
                </ul>
            </div>
            <div class="info-block">
                <h4>사용되는 데이터</h4>
                <ul>
                    <li>배터리 충전/방전 사이클 데이터 (1ms 단위 고속 기록)</li>
                    <li>전압, 전류, 온도 실시간 측정값</li>
                    <li>과거 실험 결과 데이터베이스</li>
                    <li>폐배터리의 전압, 온도, 전류 실시간 데이터</li>
                </ul>
            </div>
        </div>

        <h4 style="margin-bottom:0.5rem; color:#888; font-size:0.85rem; text-transform:uppercase; letter-spacing:0.05em;">AI 작업 흐름</h4>
        <div class="flow">
            <div class="flow-step flow-human">&#128104;&#8205;&#128300; 연구원<br>명령 입력</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-ai lg">&#129302; AI<br>실험 설계</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-human">&#9881; 장비<br>자동 실행</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-ai lg">&#129302; AI<br>실시간 분석</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-result">&#128196; 보고서<br>자동 생성</div>
        </div>

        <div class="ba-grid">
            <div class="ba-col ba-before">
                <h4>Before AI (기존)</h4>
                <ul>
                    <li>&#10060; 연구원이 수작업으로 실험 조건을 계산</li>
                    <li>&#10060; 데이터를 엑셀로 수동 분석 (수 시간 소요)</li>
                    <li>&#10060; 폐배터리 방전 시 화재 위험에 노출</li>
                </ul>
            </div>
            <div class="ba-arrow">&#10132;</div>
            <div class="ba-col ba-after">
                <h4>After AI (현재)</h4>
                <ul>
                    <li>&#9989; AI가 최적의 실험 조건을 즉시 생성</li>
                    <li>&#9989; 실시간 자동 분석 및 보고서 작성</li>
                    <li>&#9989; AI가 안전한 방전 조건을 자동 계산</li>
                </ul>
            </div>
        </div>

        <div class="kpi-row">
            <div class="kpi lg">
                <div class="number">30%+</div>
                <div class="unit">생산성 향상</div>
            </div>
            <div class="kpi lg">
                <div class="number">20대</div>
                <div class="unit">장비 납품<br><small>LG에너지솔루션 오창</small></div>
            </div>
            <div class="kpi lg">
                <div class="number">3종</div>
                <div class="unit">AI 장비 라인업</div>
            </div>
        </div>

        <div style="display:flex; gap:1.5rem; align-items:flex-start; flex-wrap:wrap;">
            <div style="flex:1; min-width:200px;">
                <h4 style="margin-bottom:0.5rem; color:#888; font-size:0.85rem; text-transform:uppercase; letter-spacing:0.05em;">핵심 역할</h4>
                <p style="font-size:0.95rem;">IT 서비스 기업(SW)과 에이테크놀로지(HW) 공동 개발. <strong>생산성 30% 이상 향상</strong> 기대</p>
            </div>
        </div>

        <div class="booth-photo">
            <img src="data:image/png;base64,{img_lgcns}" alt="LG CNS 부스">
            <div class="caption">인터배터리 2026 LG CNS 전시 부스 &mdash; AI 배터리 장비 솔루션 전시 모습</div>
        </div>
    </div>
</div>
</section>

<hr class="section-divider">

<!-- ===== SAMSUNG SDI ===== -->
<section id="samsung">
<div class="container">
    <h2><span class="accent">03-2</span> Samsung SDI</h2>
    <p class="section-desc">AI로 배터리 건강을 진단하다</p>

    <div class="company-section">
        <div class="company-header">
            <div class="company-logo ss">SDI</div>
            <div>
                <h3>SBI (Samsung Battery Intelligence)</h3>
                <div class="tagline">AI가 전 세계 1,400개 배터리 현장을 24시간 진단하고 사고를 예방</div>
            </div>
        </div>

        <div class="info-grid">
            <div class="info-block">
                <h4>AI가 하는 일</h4>
                <ul>
                    <li>전 세계 1,400개 이상의 배터리 현장을 AI가 24시간 감시</li>
                    <li>배터리의 남은 수명을 예측하여 교체 시기를 알림</li>
                    <li>미세한 이상 징후를 사전에 감지하여 화재 등 사고 예방</li>
                    <li>AI가 스스로 판단하여 배터리 건강 상태 진단 리포트를 생성</li>
                    <li>위험 감지 시 자동으로 전력을 차단하여 안전 확보</li>
                </ul>
            </div>
            <div class="info-block">
                <h4>사용되는 데이터</h4>
                <ul>
                    <li>전 세계 1,400개 배터리 현장의 운영 데이터</li>
                    <li>전압, 온도, 배터리 건강 상태 실시간 센서 데이터</li>
                    <li>배터리 노화 속도 및 출력 변동 패턴 데이터</li>
                    <li>과거 이상 징후 및 사고 이력 데이터</li>
                </ul>
            </div>
        </div>

        <div class="info-grid">
            <div class="info-block">
                <h4>AI 제조 품질 검사 (Vision AI)</h4>
                <ul>
                    <li>AI 이미지 분석으로 배터리 불량을 자동 검출</li>
                    <li>X-ray + AI로 배터리 내부 이물질 탐지</li>
                    <li>전 제조 공정에서 500가지 품질 항목을 AI로 체크</li>
                </ul>
            </div>
            <div class="info-block">
                <h4>차세대 AI 배터리 관리 시스템</h4>
                <ul>
                    <li>AI 적용 시 배터리 수명 15% 향상</li>
                    <li>전기차 주행거리 6% 증가, 에너지 출력 10% 향상</li>
                    <li>기존: 충전량만 확인 &rarr; AI: 수명 및 결함까지 예측</li>
                </ul>
            </div>
        </div>

        <h4 style="margin-bottom:0.5rem; color:#888; font-size:0.85rem; text-transform:uppercase; letter-spacing:0.05em;">AI 작업 흐름</h4>
        <div class="flow">
            <div class="flow-step flow-human">&#128267; 배터리 센서<br>데이터 수집</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-ai ss">&#129302; SBI AI<br>실시간 분석</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-ai ss">&#129302; 배터리 상태<br>진단 및 예측</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-result">&#128196; 진단 리포트<br>자동 생성</div>
        </div>
        <div class="flow" style="margin-top:0.5rem;">
            <div class="flow-step flow-ai ss">&#129302; SBI AI<br>이상 감지</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step" style="background:#ffe0e0; color:#c0392b; border:2px solid #e74c3c; border-radius:10px; padding:0.6rem 1.1rem; font-size:0.82rem; font-weight:600;">&#9888; 위험 판단</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step" style="background:#c0392b; color:#fff; border-radius:10px; padding:0.6rem 1.1rem; font-size:0.82rem; font-weight:600;">&#9889; 자동<br>전력 차단</div>
        </div>

        <div class="ba-grid">
            <div class="ba-col ba-before">
                <h4>Before AI (기존)</h4>
                <ul>
                    <li>&#10060; 사후 대응: 문제가 발생한 후에야 감지</li>
                    <li>&#10060; 사람이 수동으로 데이터를 분석</li>
                    <li>&#10060; 배터리 교체 시기를 경험에 의존하여 판단</li>
                </ul>
            </div>
            <div class="ba-arrow">&#10132;</div>
            <div class="ba-col ba-after">
                <h4>After AI (현재)</h4>
                <ul>
                    <li>&#9989; 사전 예방: AI가 이상 징후를 미리 감지</li>
                    <li>&#9989; 전 세계 1,400개 현장을 AI가 24시간 자동 감시</li>
                    <li>&#9989; 정확한 수명 예측으로 최적의 교체 시기 결정</li>
                </ul>
            </div>
        </div>

        <div class="kpi-row">
            <div class="kpi ss">
                <div class="number">1,400+</div>
                <div class="unit">모니터링 현장</div>
            </div>
            <div class="kpi ss">
                <div class="number">+15%</div>
                <div class="unit">배터리 수명 향상</div>
            </div>
            <div class="kpi ss">
                <div class="number">+10%</div>
                <div class="unit">에너지 출력 향상</div>
            </div>
        </div>

        <div style="display:flex; gap:1.5rem; align-items:flex-start; flex-wrap:wrap;">
            <div style="flex:1; min-width:200px;">
                <h4 style="margin-bottom:0.5rem; color:#888; font-size:0.85rem; text-transform:uppercase; letter-spacing:0.05em;">핵심 역할</h4>
                <p style="font-size:0.95rem;">슬로건: <strong>'AI thinks, Battery enables'</strong> (AI의 상상, 배터리가 현실로)</p>
            </div>
        </div>

        <div class="booth-photo">
            <img src="data:image/png;base64,{img_samsung}" alt="Samsung SDI 부스">
            <div class="caption">인터배터리 2026 Samsung SDI 전시 부스 &mdash; SBI 배터리 진단 시스템 전시 모습</div>
        </div>
    </div>
</div>
</section>

<hr class="section-divider">

<!-- ===== SK ON ===== -->
<section id="skon">
<div class="container">
    <h2><span class="accent">03-3</span> SK ON</h2>
    <p class="section-desc">AI 연구원이 배터리를 설계하다</p>

    <div class="company-section">
        <div class="company-header">
            <div class="company-logo sk">SK<br>ON</div>
            <div>
                <h3>ADAM &amp; AI 연구원 체계</h3>
                <div class="tagline">AI 연구원이 사람 연구원과 함께 배터리를 설계하고 성능을 예측하는 시스템</div>
            </div>
        </div>

        <div class="info-grid">
            <div class="info-block">
                <h4>AI가 하는 일</h4>
                <ul>
                    <li><strong>셀 설계 AI</strong>: 고객 요청서를 받으면 다수의 설계안을 자동 생성</li>
                    <li><strong>성능 예측 AI</strong>: 실제 배터리를 만들지 않고도 성능을 미리 예측</li>
                    <li><strong>원가 산출 AI</strong>: 재료비, 공정비 등 제조 비용을 자동 계산</li>
                    <li><strong>비전 AI</strong>: 제조 공정에서 불량품을 자동으로 검출</li>
                    <li><strong>수명 예측 AI</strong>: 초기 데이터만으로 배터리 수명을 예측</li>
                </ul>
            </div>
            <div class="info-block">
                <h4>사용되는 데이터</h4>
                <ul>
                    <li>과거 배터리 셀 설계 데이터베이스</li>
                    <li>배터리 소재 특성 데이터</li>
                    <li>성능 시험 결과 데이터</li>
                    <li>원가 데이터 (재료비, 공정비 등)</li>
                    <li>제조 공정 이미지 데이터 (비전 AI용)</li>
                </ul>
            </div>
        </div>

        <div class="info-grid">
            <div class="info-block">
                <h4>AI 연구원 핵심 성과</h4>
                <ul>
                    <li>배터리 설계 기간: 기존 대비 1/3로 단축</li>
                    <li>원가 분석 속도: 기존 대비 700배 향상</li>
                    <li>검토 가능한 설계안: 기존 대비 15배 증가</li>
                </ul>
            </div>
            <div class="info-block">
                <h4>미래 계획</h4>
                <ul>
                    <li>2028년까지 '배터리 전용 AI' 자체 구축 목표</li>
                    <li>소재 개발 AI 연구원 개발 중 (개발 기간 50% 단축 기대)</li>
                    <li>제조, 품질, 영업까지 전사적 AI 확대 예정</li>
                </ul>
            </div>
        </div>

        <h4 style="margin-bottom:0.5rem; color:#888; font-size:0.85rem; text-transform:uppercase; letter-spacing:0.05em;">ADAM AI 파이프라인</h4>
        <div class="flow">
            <div class="flow-step flow-human">&#128104;&#8205;&#128300; 고객<br>요청서 입력</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-ai sk">&#129302; 셀 설계<br>AI</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-ai sk">&#129302; 성능<br>예측 AI</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-ai sk">&#129302; 원가<br>산출 AI</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-result">&#128196; 보고서<br>자동 작성</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-human">&#128104;&#8205;&#128300; 연구원<br>최종 판단</div>
            <div class="flow-arrow">&#10132;</div>
            <div class="flow-step flow-result">&#9989; 최종<br>설계안</div>
        </div>

        <div class="kpi-row">
            <div class="kpi sk">
                <div class="number">1/3</div>
                <div class="unit">설계 기간 단축<br><small>기존 대비 약 1/3 수준</small></div>
            </div>
            <div class="kpi sk">
                <div class="number">700x</div>
                <div class="unit">원가 분석 속도<br><small>기존 대비 700배 향상</small></div>
            </div>
            <div class="kpi sk">
                <div class="number">15x</div>
                <div class="unit">설계안 검토량<br><small>기존 대비 15배 증가</small></div>
            </div>
        </div>

        <div class="ba-grid">
            <div class="ba-col ba-before">
                <h4>Before AI (기존)</h4>
                <ul>
                    <li>&#10060; 연구원이 수작업으로 셀 설계 (시행착오 반복)</li>
                    <li>&#10060; 한 번에 하나의 설계안만 검토 가능</li>
                    <li>&#10060; 원가 계산에 별도 팀과 수 주 소요</li>
                </ul>
            </div>
            <div class="ba-arrow">&#10132;</div>
            <div class="ba-col ba-after">
                <h4>After AI (현재)</h4>
                <ul>
                    <li>&#9989; AI가 다수의 설계안을 동시에 자동 생성</li>
                    <li>&#9989; 실제 배터리 제작 없이 성능과 원가를 예측</li>
                    <li>&#9989; 연구원은 안전성 판단과 최종 결정에 집중</li>
                </ul>
            </div>
        </div>

        <div style="display:flex; gap:1.5rem; align-items:flex-start; flex-wrap:wrap;">
            <div style="flex:1; min-width:200px;">
                <h4 style="margin-bottom:0.5rem; color:#888; font-size:0.85rem; text-transform:uppercase; letter-spacing:0.05em;">핵심 역할</h4>
                <p style="font-size:0.95rem;">'AI 연구원'은 사람을 대체하는 것이 아닌, <strong>함께 일하는 AI 동료</strong> 개념</p>
            </div>
        </div>

        <div class="booth-photo">
            <img src="data:image/png;base64,{img_skon}" alt="SK ON 부스">
            <div class="caption">인터배터리 2026 SK ON 전시 부스 &mdash; AI 연구원, ADAM 시스템 소개</div>
        </div>
    </div>
</div>
</section>

<hr class="section-divider">

<!-- ===== COMPARISON ===== -->
<section id="comparison">
<div class="container">
    <h2><span class="accent">04</span> 비교 분석</h2>
    <p class="section-desc">3사의 AI 전략을 다양한 관점에서 비교합니다.</p>

    <!-- 4-1. AI Application Matrix -->
    <div class="card" style="margin-bottom:2rem;">
        <h3>AI 적용 영역 비교</h3>
        <p style="color:#666; font-size:0.9rem; margin-bottom:1rem;">배터리 밸류체인의 어느 단계에 AI를 적용했는지 보여줍니다.</p>
        <table class="comp-table">
            <thead>
                <tr>
                    <th>밸류체인 단계</th>
                    <th style="color:#f8a0b8;">LG CNS</th>
                    <th style="color:#8c9eff;">Samsung SDI</th>
                    <th style="color:#ffe082;">SK ON</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>&#9997; R&amp;D / 설계</td>
                    <td class="dash">&mdash;</td>
                    <td class="partial">&#9675; 부분</td>
                    <td class="check">&#9679; 핵심</td>
                </tr>
                <tr>
                    <td>&#128300; 시험 / 분석</td>
                    <td class="check">&#9679; 핵심</td>
                    <td class="dash">&mdash;</td>
                    <td class="partial">&#9675; 부분</td>
                </tr>
                <tr>
                    <td>&#9881; 생산 / 품질</td>
                    <td class="partial">&#9675; 부분</td>
                    <td class="check">&#9679; 핵심</td>
                    <td class="check">&#9679; 핵심</td>
                </tr>
                <tr>
                    <td>&#128225; 운영 / 모니터링</td>
                    <td class="dash">&mdash;</td>
                    <td class="check">&#9679; 핵심</td>
                    <td class="dash">&mdash;</td>
                </tr>
                <tr>
                    <td>&#9888; 안전 / 예방</td>
                    <td class="partial">&#9675; 부분</td>
                    <td class="check">&#9679; 핵심</td>
                    <td class="partial">&#9675; 부분</td>
                </tr>
            </tbody>
        </table>
        <p style="color:#888; font-size:0.82rem; margin-top:0.5rem;">
            &#9679; 핵심 = AI 적용의 주력 영역 &nbsp;|&nbsp; &#9675; 부분 = 보조적 활용 &nbsp;|&nbsp; &mdash; = 미적용 또는 미공개
        </p>
    </div>

    <!-- 4-2. Radar Chart -->
    <div class="card" style="margin-bottom:2rem;">
        <h3>AI 활용 수준 비교 (레이더 차트)</h3>
        <p style="color:#666; font-size:0.9rem; margin-bottom:1rem;">5가지 관점에서 각사의 AI 활용 수준을 시각적으로 비교합니다.</p>
        <div class="chart-container">
            <canvas id="radarChart"></canvas>
        </div>
        <p style="color:#888; font-size:0.82rem; text-align:center;">
            * 인터배터리 2026 전시 내용 및 공개 자료 기반의 상대적 평가
        </p>
    </div>

    <!-- 4-3. Summary Table -->
    <div class="card">
        <h3>종합 비교표</h3>
        <table class="comp-table">
            <thead>
                <tr>
                    <th>항목</th>
                    <th style="color:#f8a0b8;">LG CNS</th>
                    <th style="color:#8c9eff;">Samsung SDI</th>
                    <th style="color:#ffe082;">SK ON</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>회사 유형</td>
                    <td>IT 서비스</td>
                    <td>배터리 제조</td>
                    <td>배터리 제조</td>
                </tr>
                <tr>
                    <td>AI 핵심 제품</td>
                    <td>AX 사이클러<br>포메이션<br>Factova HED</td>
                    <td>SBI<br>AI BMS<br>Vision AI</td>
                    <td>ADAM<br>비전 AI<br>수명 예측 AI</td>
                </tr>
                <tr>
                    <td>AI가 하는 일</td>
                    <td>실험 자동 설계<br>결과 분석<br>보고서 생성</td>
                    <td>24시간 배터리 감시<br>수명 예측<br>사고 예방</td>
                    <td>셀 설계 자동화<br>성능 및 원가 예측</td>
                </tr>
                <tr>
                    <td>주요 성과</td>
                    <td>생산성 30%&#8593;<br>장비 20대 납품</td>
                    <td>수명 15%&#8593;<br>1,400개 현장 관리</td>
                    <td>설계 기간 1/3<br>원가분석 700배&#8593;</td>
                </tr>
                <tr>
                    <td>적용 단계</td>
                    <td>시험 및 분석<br>폐배터리 처리</td>
                    <td>운영 및 안전<br>생산 및 품질</td>
                    <td>R&amp;D 및 설계<br>생산 및 품질</td>
                </tr>
                <tr>
                    <td>비유적 역할</td>
                    <td>배터리 자동 실험실</td>
                    <td>24시간 배터리 전담의</td>
                    <td>AI 동료 연구원</td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
</section>

<hr class="section-divider">

<!-- ===== INSIGHTS ===== -->
<section id="insights">
<div class="container">
    <h2><span class="accent">05</span> 시사점</h2>
    <p class="section-desc">3사의 AI 활용 분석을 통해 도출한 핵심 시사점입니다.</p>

    <div class="insight-cards">
        <div class="insight-card">
            <div class="insight-num">01</div>
            <h4>밸류체인 전 단계에서 AI 도입 가속화</h4>
            <p>설계(SK ON) &rarr; 시험(LG CNS) &rarr; 생산(삼성SDI 및 SK ON) &rarr; 운영 및 안전(삼성SDI) &rarr; 폐배터리(LG CNS)까지 밸류체인의 모든 단계에서 AI 도입이 진행 중. 배터리 산업에서 AI는 선택이 아닌 <strong>필수</strong>.</p>
        </div>
        <div class="insight-card">
            <div class="insight-num">02</div>
            <h4>각사의 차별화된 AI 전략</h4>
            <p>LG CNS = 실험 자동화 (IT기업의 AI 소프트웨어 강점) / Samsung SDI = 배터리 진단 (1,400개 현장 데이터) / SK ON = AI 연구원 (설계 및 예측 자동화) / 같은 'AI'라도 접근 방식이 <strong>완전히 다름</strong>.</p>
        </div>
        <div class="insight-card">
            <div class="insight-num">03</div>
            <h4>AI는 사람을 '대체'하지 않는다</h4>
            <p>SK ON: AI 연구원은 '함께 일하는 동료'. AI가 반복 작업을 대신 수행하고, 사람은 최종 판단과 창의적 결정에 집중. 3사 모두 AI를 '도구'가 아닌 <strong>'파트너'</strong>로 활용.</p>
        </div>
    </div>

    <div class="sl-box">
        <h3>SL에 주는 시사점</h3>
        <ul class="sl-list">
            <li>
                <span class="sl-num">1</span>
                <span><strong>글로벌 AI 경쟁이 심화되고 있습니다</strong> &mdash; 중국 CATL(세계 1위 배터리 기업, 점유율 39%)은 이미 5,000만 건 이상의 데이터로 AI를 운영 중입니다. 한국 3사 합산 점유율은 약 16%이며, AI 도입이 늦을수록 격차가 벌어집니다.</span>
            </li>
            <li>
                <span class="sl-num">2</span>
                <span><strong>데이터 확보가 AI 성공의 핵심입니다</strong> &mdash; 삼성SDI는 1,400개 현장, SK ON은 설계 DB, LG CNS는 실험 데이터를 AI에 활용합니다. 체계적인 데이터 관리가 AI 도입의 첫걸음입니다.</span>
            </li>
            <li>
                <span class="sl-num">3</span>
                <span><strong>작은 것부터 시작하는 것이 현실적입니다</strong> &mdash; 처음부터 대규모 AI 시스템이 아닌, 특정 업무(예: 시험 자동화, 데이터 분석)부터 작은 규모로 시작하여 점진적으로 확대하는 것이 현실적입니다.</span>
            </li>
        </ul>
    </div>
</div>
</section>

<hr class="section-divider">

<!-- ===== APPENDIX ===== -->
<section id="appendix">
<div class="container">
    <h2><span class="accent">06</span> 부록</h2>

    <div class="card" style="margin-bottom:2rem;">
        <h3>AI 용어집</h3>
        <dl class="glossary">
            <div class="glossary-item">
                <dt>AI (인공지능)</dt>
                <dd>사람의 학습 및 판단 능력을 컴퓨터로 구현한 기술. 데이터에서 패턴을 찾고 예측합니다.</dd>
            </div>
            <div class="glossary-item">
                <dt>SBI</dt>
                <dd>Samsung Battery Intelligence. 삼성SDI의 AI 배터리 진단 소프트웨어입니다.</dd>
            </div>
            <div class="glossary-item">
                <dt>ADAM</dt>
                <dd>AI-Based Design &amp; Analysis Machine. SK ON의 AI 배터리 설계 시스템입니다.</dd>
            </div>
            <div class="glossary-item">
                <dt>BMS (Battery Management System)</dt>
                <dd>배터리 충방전과 온도를 관리하는 시스템입니다.</dd>
            </div>
            <div class="glossary-item">
                <dt>ESS (Energy Storage System)</dt>
                <dd>대용량 에너지를 저장하는 장치입니다.</dd>
            </div>
            <div class="glossary-item">
                <dt>에이전틱 AI</dt>
                <dd>사람의 지시 없이도 스스로 판단하고 실행하는 AI. LG CNS가 배터리 장비에 적용.</dd>
            </div>
            <div class="glossary-item">
                <dt>RFQ (Request for Quotation)</dt>
                <dd>고객이 원하는 배터리 사양을 적어 보내는 요청서.</dd>
            </div>
            <div class="glossary-item">
                <dt>비전 AI</dt>
                <dd>카메라 이미지를 AI가 분석하여 불량품을 자동으로 검출하는 기술.</dd>
            </div>
            <div class="glossary-item">
                <dt>밸류체인 (Value Chain)</dt>
                <dd>설계 &rarr; 시험 &rarr; 생산 &rarr; 운영의 전체 과정. 제품이 만들어지는 모든 단계입니다.</dd>
            </div>
        </dl>
    </div>

    <div class="card">
        <h3>출처</h3>
        <ul style="list-style:none;">
            <li style="padding:0.4rem 0; font-size:0.9rem; color:#555;">&#128196; 인터배터리 2026 전시회 현장 조사 (2026.3.11~13, 서울 코엑스)</li>
            <li style="padding:0.4rem 0; font-size:0.9rem; color:#555;">&#128196; LG CNS &mdash; AI로 배터리 공정 혁신 (스마트비즈, kidd, ZDNet, 2026.3)</li>
            <li style="padding:0.4rem 0; font-size:0.9rem; color:#555;">&#128196; 삼성SDI &mdash; AI 기반 ESS 화재 예방 SW 'SBI' (서울경제, 헤럴드경제, Korea Herald)</li>
            <li style="padding:0.4rem 0; font-size:0.9rem; color:#555;">&#128196; SK ON &mdash; AI 연구원이 배터리 개발 참여 (ASK inno, 오늘경제, 2026.3)</li>
            <li style="padding:0.4rem 0; font-size:0.9rem; color:#555;">&#128196; 배터리 3사 기술 수장 AI R&amp;D 혁신 발표 (헤럴드경제, 2026.3)</li>
            <li style="padding:0.4rem 0; font-size:0.9rem; color:#555;">&#128196; IDTechEx &mdash; AI-Driven Battery Technology 2025-2035</li>
            <li style="padding:0.4rem 0; font-size:0.9rem; color:#555;">&#128196; CATL AI Strategy Analysis (SCMP, Klover.ai, 2026.3)</li>
        </ul>
    </div>
</div>
</section>

<!-- ===== FOOTER ===== -->
<footer>
<div class="container">
    <p>국내 배터리 3사 AI 활용 현황 분석 보고서 &middot; 인터배터리 2026 기반 &middot; SL &middot; 2026년 3월</p>
</div>
</footer>

<!-- ===== CHART.JS SCRIPTS ===== -->
<script>
document.addEventListener('DOMContentLoaded', function() {{
    // Radar Chart
    const ctx = document.getElementById('radarChart').getContext('2d');
    new Chart(ctx, {{
        type: 'radar',
        data: {{
            labels: ['자동화 수준', '데이터 활용', '실시간성', 'AI 제품 다양성', '성과 입증도'],
            datasets: [
                {{
                    label: 'LG CNS',
                    data: [8, 7, 7, 8, 7],
                    borderColor: '#A50034',
                    backgroundColor: 'rgba(165, 0, 52, 0.1)',
                    borderWidth: 2,
                    pointBackgroundColor: '#A50034',
                }},
                {{
                    label: 'Samsung SDI',
                    data: [7, 9, 9, 8, 7],
                    borderColor: '#1428A0',
                    backgroundColor: 'rgba(20, 40, 160, 0.1)',
                    borderWidth: 2,
                    pointBackgroundColor: '#1428A0',
                }},
                {{
                    label: 'SK ON',
                    data: [9, 7, 5, 8, 9],
                    borderColor: '#F7B500',
                    backgroundColor: 'rgba(247, 181, 0, 0.1)',
                    borderWidth: 2,
                    pointBackgroundColor: '#F7B500',
                }}
            ]
        }},
        options: {{
            responsive: true,
            plugins: {{
                legend: {{
                    position: 'bottom',
                    labels: {{
                        padding: 20,
                        font: {{ family: "'Pretendard', sans-serif", size: 13 }}
                    }}
                }}
            }},
            scales: {{
                r: {{
                    beginAtZero: true,
                    max: 10,
                    ticks: {{
                        stepSize: 2,
                        font: {{ size: 11 }},
                        backdropColor: 'transparent'
                    }},
                    pointLabels: {{
                        font: {{ family: "'Pretendard', sans-serif", size: 12, weight: '600' }}
                    }},
                    grid: {{
                        color: 'rgba(0,0,0,0.08)'
                    }}
                }}
            }}
        }}
    }});

    // Active nav highlight on scroll
    const sections = document.querySelectorAll('section[id]');
    const navLinks = document.querySelectorAll('nav a');
    window.addEventListener('scroll', () => {{
        let current = '';
        sections.forEach(section => {{
            const top = section.offsetTop - 100;
            if (window.scrollY >= top) current = section.getAttribute('id');
        }});
        navLinks.forEach(link => {{
            link.classList.remove('active');
            if (link.getAttribute('href') === '#' + current) link.classList.add('active');
        }});
    }});
}});
</script>

</body>
</html>"""

    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)

    size_mb = os.path.getsize(OUTPUT) / (1024 * 1024)
    print(f"Report generated: {OUTPUT} ({size_mb:.1f} MB)")


if __name__ == "__main__":
    generate()
