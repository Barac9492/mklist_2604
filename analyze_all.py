#!/usr/bin/env python3
"""Analyze all 8 weeks of MK 신설법인 data for startup candidates."""

import xlrd
import re
import json
import os
from collections import defaultdict

CACHE_FILE = 'data/llm_cache.json'
try:
    with open(CACHE_FILE, 'r', encoding='utf-8') as f:
        llm_cache = json.load(f)
except FileNotFoundError:
    llm_cache = {}

is_openai_ready = False
if os.environ.get("OPENAI_API_KEY"):
    try:
        from openai import OpenAI
        client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
        is_openai_ready = True
    except ImportError:
        pass

def get_llm_tag(business_desc):
    if not is_openai_ready:
        return None
    if business_desc in llm_cache:
        return llm_cache[business_desc]
        
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a tech VC analyst. Read the Korean business description and respond with EXACTLY ONE short English tech keyword or phrase (max 20 characters, e.g. 'B2B SaaS', 'EdTech', 'Vision AI', 'Robotics', 'Digital Twin') that best categorizes their core model. Output only the English keyword, nothing else."},
                {"role": "user", "content": f"Description: {business_desc}"}
            ],
            temperature=0.1,
            max_tokens=10
        )
        tag = response.choices[0].message.content.strip()
        tag = tag.replace('"', '').replace("'", "")
        llm_cache[business_desc] = tag
        return tag
    except Exception as e:
        print(f"  [!] OpenAI API error: {e}")
        return None

def get_outreach_draft(name, business_desc, llm_tag):
    if not is_openai_ready: return None
    cache_key = f"draft_{name}_{business_desc}"
    if cache_key in llm_cache: return llm_cache[cache_key]
    
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a professional Korean top-tier VC analyst. Write a concise, polite, and persuasive cold-email draft (in Korean) to a newly incorporated startup CEO. Praise their specific deep-tech approach based on their 'business description' and 'LLM tag', and propose a brief coffee chat to discuss potential early-stage cooperation. Tone should be highly professional, avoiding cliches. Do not include subject line, just the body."},
                {"role": "user", "content": f"Company Name: {name}\nTag: {llm_tag}\nBusiness Description: {business_desc}"}
            ],
            temperature=0.3,
            max_tokens=250
        )
        draft = response.choices[0].message.content.strip()
        llm_cache[cache_key] = draft
        return draft
    except Exception as e:
        return None

def get_lp_teaser(name, business_desc, llm_tag):
    if not is_openai_ready: return None
    cache_key = f"lp_{name}_{business_desc}"
    if cache_key in llm_cache: return llm_cache[cache_key]
    
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a top-tier VC analyst. Write a highly professional 1-paragraph investment teaser in ENGLISH tailored for foreign LPs/investors. Summarize the Korean startup's 'business description' professionally, highlighting their tech. Keep it under 60 words. No intro/outro."},
                {"role": "user", "content": f"Company: {name}\nTag: {llm_tag}\nDesc: {business_desc}"}
            ],
            temperature=0.3,
            max_tokens=150
        )
        draft = response.choices[0].message.content.strip()
        llm_cache[cache_key] = draft
        return draft
    except Exception as e:
        return None

# ── Import parse + score logic from filter_v2 ──

def parse_xls(filepath):
    wb = xlrd.open_workbook(filepath, encoding_override='cp949')
    sh = wb.sheet_by_index(0)
    companies = []
    current_region = ""
    current_sector = ""

    for row_idx in range(sh.nrows):
        row = [str(sh.cell_value(row_idx, col)).strip() for col in range(sh.ncols)]
        marker = row[0]
        if not any(row):
            continue

        # Region detection
        if not row[2] and row[1]:
            cleaned = row[1].replace(' ', '')
            regions = ['서울','부산','대구','인천','광주','대전','울산',
                       '경기','강원','충북','충남','전북','전남','경북','경남','제주','세종']
            if cleaned in regions:
                current_region = cleaned
                continue

        if '단위' in row[1] or ('상호' in row[1] and '대표자' in row[2]):
            continue
        if marker == '◆':
            sector_text = row[1].replace(' ', '').strip()
            if sector_text:
                current_sector = sector_text
            continue
        if marker == '▷':
            name = row[1].strip()
            ceo = row[2].strip()
            try:
                capital = int(float(row[3].strip())) if row[3].strip() else 0
            except:
                capital = 0
            business = row[4].strip()
            address = row[5].strip()
            if not name:
                continue
            companies.append({
                'name': name, 'ceo': ceo, 'capital': capital,
                'business': business, 'address': address,
                'region': current_region, 'sector': current_sector,
            })
    return companies


TIER1_KEYWORDS = [
    '인공지능', 'AI', '머신러닝', '딥러닝', '자연어처리', 'LLM',
    '소프트웨어 개발', '소프트웨어 연구', '앱 개발', '애플리케이션 개발',
    '플랫폼 개발', '플랫폼 운영', 'SaaS', '클라우드', 'API',
    '블록체인', '핀테크', '가상자산', 'DeFi', 'Web3',
    '바이오', '헬스케어', '디지털헬스', '원격의료', '디지털 치료',
    '에듀테크', '이러닝', '온라인교육 플랫폼',
    '빅데이터', '데이터 분석', '데이터 플랫폼',
    'IoT', '사물인터넷', '스마트팜', '스마트시티',
    '자율주행', '로보틱스', '드론',
    '이커머스', '전자상거래 플랫폼', '마켓플레이스',
    '구독', '커스터마이징 플랫폼',
    '푸드테크', '배양육', '대체단백',
    '전기차', '배터리', '수소',
    '게임 개발', '콘텐츠 제작', '미디어 플랫폼', '스트리밍',
]

EXCLUDE_PATTERNS = [
    r'건축공사', r'토목공사', r'철거', r'인테리어\s*공사', r'도장공사',
    r'방수공사', r'전기공사', r'설비공사', r'소방공사', r'조경공사',
    r'포장공사', r'측량', r'철근', r'콘크리트', r'비계',
    r'부동산\s*(임대|매매|중개)', r'임대업$', r'건물\s*관리',
    r'화물\s*(운송|운수)', r'용달', r'퀵서비스', r'택배$',
    r'^음식점', r'^식당', r'^카페', r'커피\s*전문', r'베이커리', r'정육',
    r'반찬', r'프랜차이즈\s*가맹',
    r'미용실', r'네일', r'피부관리', r'세차', r'세탁', r'청소업',
    r'철물', r'문구', r'잡화', r'슈퍼', r'편의점',
    r'학원$', r'과외$', r'입시',
    r'^약국', r'^한의원', r'^치과', r'^병원',
    r'^농업', r'^축산', r'^어업', r'^임업',
    r'시멘트', r'골재', r'모래', r'자갈', r'석재',
    r'태양광\s*(발전|시공|설치|공사|컨설팅)$',
    r'^신재생에너지(업)?$',
    r'^(도|소매|도소매|무역|수출입)(업)?$',
    r'잡화\s*도소매',
    r'대부업', r'대부중개', r'담보대출',
]


def is_excluded(business):
    for pat in EXCLUDE_PATTERNS:
        if re.search(pat, business):
            return True
    return False


def categorize(biz):
    biz_lower = biz.lower()
    if any(k in biz_lower for k in ['인공지능', 'ai', '머신러닝', '딥러닝', 'llm']):
        return 'AI/ML'
    elif any(k in biz_lower for k in ['바이오', '헬스케어', '의료', '디지털헬스']):
        return 'Bio/Health'
    elif any(k in biz_lower for k in ['플랫폼', '이커머스', '전자상거래', '마켓플레이스']):
        return 'Platform'
    elif any(k in biz_lower for k in ['블록체인', '핀테크', '가상자산', 'defi', 'web3']):
        return 'Blockchain/Fintech'
    elif any(k in biz_lower for k in ['소프트웨어', '앱', '클라우드', 'saas', 'api']):
        return 'Software'
    elif any(k in biz_lower for k in ['게임', '콘텐츠', '미디어', '영상', '스트리밍']):
        return 'Content/Media'
    elif any(k in biz_lower for k in ['에듀', '교육', '이러닝']):
        return 'EdTech'
    elif any(k in biz_lower for k in ['에너지', '태양광', '배터리', '수소', '전기차']):
        return 'CleanTech'
    elif any(k in biz_lower for k in ['로봇', 'iot', '드론', '자율주행', '스마트']):
        return 'DeepTech'
    elif any(k in biz_lower for k in ['데이터', '빅데이터']):
        return 'Data'
    elif any(k in biz_lower for k in ['투자', '자산', '결제', '페이', '금융']):
        return 'Finance'
    else:
        return 'Other'


def detect_talent_signals(company, base_score):
    biz = company['business'].lower()
    address = company['address'].lower()
    capital = company['capital']
    
    signals = []
    
    # 1. The Ghost Radar (Incubator Geolocation)
    incubator_keywords = [
        '역삼로 180', '역삼로 172', '역삼로 165', '역삼로 169', # Maru 180, 360, TIPS Town
        '백범로31길 21', '대왕판교로815', # Seoul Startup Hub, Pangyo Startup Campus
        '관악로 1', '대학로 291', '문지로 193', # SNU, KAIST
        '청암로 77', '유니스트길 50', # Postech, UNIST
        '홍릉로', '고려대로', '연세로' # Major University Hubs
    ]
    if any(k in address for k in incubator_keywords):
        signals.append('국가 핵심 인큐베이터 / 대학 연구소 스핀오프')

    # Legacy Hub check for business purpse text
    hub_keywords = ['대학교', '산학협력단', '카이스트', 'kaist', '포스텍', '포항공대', 'unist', 'dgist', '팁스타운', '판교테크노밸리', '마루180', '마루360', '마곡', '홍릉', '연구소기업', '사내벤처', '기술지주']
    if any(k in biz for k in hub_keywords) and '국가 핵심 인큐베이터 / 대학 연구소 스핀오프' not in signals:
        signals.append('캠퍼스/연구소/테크허브')
        
    # 2. Deep-Tech / Over-Engineered Purpose
    deeptech_keywords = ['sllm', 'rag', '트랜스포머', '펩타이드', '재조합', '엑소좀', 'npu', '화합물 반도체', '양자컴퓨팅']
    if any(k in biz for k in deeptech_keywords):
        signals.append('딥테크 고도기술')
        
    # 3. Institutional Day-1 Seed (Raised threshold to 500M KRW)
    if capital >= 500 and base_score >= 25:
        signals.append('초기 자본금 5억 이상 (기관/빌더 참여 가능성)')
        
    return signals


def score_startup(company):
    biz = company['business']
    name = company['name']
    capital = company['capital']
    score = 0

    if is_excluded(biz):
        return 0, '', [], []

    t1_matches = [kw for kw in TIER1_KEYWORDS if kw.lower() in biz.lower() or kw.lower() in name.lower()]
    if t1_matches:
        score += 25
        if len(t1_matches) > 1:
            score += len(t1_matches) * 5

    if not t1_matches:
        return 0, '', [], []  # Only Tier1

    if 50 <= capital <= 3000:
        score += 10
    elif capital > 3000:
        score += 5

    if re.search(r'[a-zA-Z]', name):
        score += 5

    for pat in ['테크$', '랩$', '스튜디오', '이노', '벤처', '파트너스']:
        if re.search(pat, name):
            score += 3
            break

    if company['sector'] in ['건설', '건자재']:
        score -= 15

    signals = detect_talent_signals(company, score)
    if '국가 핵심 인큐베이터 / 대학 연구소 스핀오프' in signals:
        score += 15
    elif signals:
        score += 5

    cat = categorize(biz)
    return max(0, score), cat, t1_matches[:3], signals


def main():
    import glob
    files = glob.glob('data/*.xls')
    files.sort()
    
    weeks = []
    # Identify the week from the filename: week_0320_0326.xls -> '3/20~3/26'
    # Or just fallback to the filename
    for filepath in files:
        fname = filepath.split('/')[-1]
        m = re.search(r'week_(\d{2})(\d{2})_(\d{2})(\d{2})\.xls', fname)
        if m:
            m1, d1, m2, d2 = [int(x) for x in m.groups()]
            period = f"{m1}/{d1}~{m2}/{d2}"
        elif 'last_week' in fname:
            continue # We will just use the week_... files if they overlap, or period fallback
        elif 'this_week' in fname:
            continue
        else:
            m2 = re.search(r'week_(.+)\.xls', fname)
            period = m2.group(1) if m2 else fname
            
        weeks.append((filepath, period))

    all_weeks = []
    trend_data = []

    for filepath, period in weeks:
        companies = parse_xls(filepath)
        startups = []
        for c in companies:
            score, cat, kws, signals = score_startup(c)
            if score >= 25:
                c['score'] = score
                c['category'] = cat
                c['talent_signals'] = signals
                if score >= 35:
                    c['llm_tag'] = get_llm_tag(c['business'])
                    # Generate Outreach Draft for elite prospects
                    if score >= 40:
                        c['outreach_draft'] = get_outreach_draft(c['name'], c['business'], c['llm_tag'])
                        c['lp_teaser'] = get_lp_teaser(c['name'], c['business'], c['llm_tag'])
                else:
                    c['llm_tag'] = None
                    c['outreach_draft'] = None
                    c['lp_teaser'] = None
                startups.append(c)

        by_cat = defaultdict(list)
        for s in startups:
            by_cat[s['category']].append(s)

        all_weeks.append((period, len(companies), startups, by_cat))
        trend_data.append({
            'period': period,
            'total': len(companies),
            'startups': len(startups),
            'by_cat': {k: len(v) for k, v in by_cat.items()},
        })

    # ── Print Weekly Summary ──
    print("=" * 90)
    print("  매일경제 신설법인 스타트업 스캐너 — 8주 종합 리포트")
    print("=" * 90)

    print(f"\n{'기간':<14} | {'총법인':>6} | {'스타트업':>7} | {'AI':>4} | {'SW':>4} | {'Plat':>4} | {'Bio':>4} | {'Content':>7} | {'Clean':>5} | {'기타':>4}")
    print("-" * 90)
    for td in trend_data:
        bc = td['by_cat']
        print(f"  {td['period']:<12} | {td['total']:>6} | {td['startups']:>7} | "
              f"{bc.get('AI/ML',0):>4} | {bc.get('Software',0):>4} | {bc.get('Platform',0):>4} | "
              f"{bc.get('Bio/Health',0):>4} | {bc.get('Content/Media',0):>7} | "
              f"{bc.get('CleanTech',0):>5} | "
              f"{bc.get('DeepTech',0)+bc.get('Data',0)+bc.get('Finance',0)+bc.get('Other',0)+bc.get('Blockchain/Fintech',0)+bc.get('EdTech',0):>4}")

    # ── Per-week detail: only score >= 35 (top picks) ──
    print(f"\n\n{'='*90}")
    print("  주차별 TOP 스타트업 후보 (점수 35+)")
    print(f"{'='*90}")

    for period, total, startups, by_cat in all_weeks:
        top = [s for s in startups if s['score'] >= 35]
        top.sort(key=lambda x: (-x['score'], x['category']))
        print(f"\n  ── {period} ({total}개 중 {len(top)}개 주목) ──")
        for s in top:
            cap = f"{s['capital']}백만" if s['capital'] else "미상"
            biz_short = s['business'][:55]
            print(f"    [{s['score']:2d}] {s['category']:<18} {s['name']:<22} | {cap:>8} | {s['region']:<4} | {biz_short}")

    # ── Category deep dive: AI/ML across all weeks ──
    print(f"\n\n{'='*90}")
    print("  AI/ML 신설법인 전체 리스트 (8주)")
    print(f"{'='*90}")

    for period, total, startups, by_cat in all_weeks:
        ai_list = by_cat.get('AI/ML', [])
        if ai_list:
            print(f"\n  [{period}] — {len(ai_list)}개")
            for s in sorted(ai_list, key=lambda x: -x['score']):
                cap = f"{s['capital']}백만" if s['capital'] else "미상"
                print(f"    {s['name']:<24} | {cap:>8} | {s['region']:<4} | {s['business'][:60]}")

    # ── Platform companies ──
    print(f"\n\n{'='*90}")
    print("  Platform 신설법인 전체 리스트 (8주)")
    print(f"{'='*90}")

    for period, total, startups, by_cat in all_weeks:
        plat_list = by_cat.get('Platform', [])
        if plat_list:
            print(f"\n  [{period}] — {len(plat_list)}개")
            for s in sorted(plat_list, key=lambda x: -x['score']):
                cap = f"{s['capital']}백만" if s['capital'] else "미상"
                print(f"    {s['name']:<24} | {cap:>8} | {s['region']:<4} | {s['business'][:60]}")

    # ── JSON export ──
    export = []
    for period, total, startups, by_cat in all_weeks:
        for s in startups:
            score = s['score']
            if score >= 45:
                grade = 'S'
            elif score >= 35:
                grade = 'A'
            elif score >= 25:
                grade = 'B'
            else:
                grade = 'C'
                
            export.append({
                'period': period,
                'name': s['name'],
                'ceo': s['ceo'],
                'capital_million_krw': s['capital'],
                'business': s['business'],
                'address': s['address'],
                'region': s['region'],
                'category': s['category'],
                'score': score,
                'investment_grade': grade,
                'talent_signals': s.get('talent_signals', []),
                'llm_tag': s.get('llm_tag', None),
                'outreach_draft': s.get('outreach_draft', None),
                'lp_teaser': s.get('lp_teaser', None)
            })

    with open('data/startups_all_weeks.json', 'w', encoding='utf-8') as f:
        json.dump(export, f, ensure_ascii=False, indent=2)
    print(f"\n\n  ✓ JSON 내보내기 완료: data/startups_all_weeks.json ({len(export)}개 항목)")
    
    # Save cache
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(llm_cache, f, ensure_ascii=False, indent=2)


if __name__ == '__main__':
    main()
