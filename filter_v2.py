#!/usr/bin/env python3
"""V2: More precise startup filtering from MK 신설법인 data."""

import xlrd
import re
import json


def parse_xls(filepath):
    """Parse the MK 신설법인 XLS file."""
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
            regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산',
                       '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주', '세종']
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


# ── Tier 1: High-signal startup keywords (tech product companies) ──
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

# ── Tier 2: Medium-signal keywords ──
TIER2_KEYWORDS = [
    '솔루션', '컨설팅', '마케팅', '광고', '디자인',
    '스튜디오', '콘텐츠', '미디어', '영상', '브랜딩',
    '커머스', '온라인', '모바일', '디지털',
    '신재생에너지', '태양광', '에너지',
    '투자', '자산관리', '결제', '페이',
    '테크', '이노베이션', '벤처',
    '중개', '매칭', '공유',
    '로봇', '의료기기',
]

# ── Hard exclude: traditional/non-startup businesses ──
EXCLUDE_PATTERNS = [
    # Construction
    r'건축공사', r'토목공사', r'철거', r'인테리어\s*공사', r'도장공사',
    r'방수공사', r'전기공사', r'설비공사', r'소방공사', r'조경공사',
    r'포장공사', r'측량', r'철근', r'콘크리트', r'비계',
    # Real estate
    r'부동산\s*(임대|매매|중개)', r'임대업$', r'건물\s*관리',
    # Transport (traditional)
    r'화물\s*(운송|운수)', r'용달', r'퀵서비스', r'택배$',
    # Food service
    r'^음식점', r'^식당', r'^카페', r'커피\s*전문', r'베이커리', r'정육',
    r'반찬', r'프랜차이즈\s*가맹',
    # Personal services
    r'미용실', r'네일', r'피부관리', r'세차', r'세탁', r'청소업',
    # Retail
    r'철물', r'문구', r'잡화', r'슈퍼', r'편의점',
    # Traditional education
    r'학원$', r'과외$', r'입시',
    # Medical (traditional)
    r'^약국', r'^한의원', r'^치과', r'^병원',
    # Primary industry
    r'^농업', r'^축산', r'^어업', r'^임업',
    # Materials
    r'시멘트', r'골재', r'모래', r'자갈', r'석재',
    # Solar install only (not tech)
    r'태양광\s*(발전|시공|설치|공사|컨설팅)$',
    r'^신재생에너지(업)?$',
    # Pure trading
    r'^(도|소매|도소매|무역|수출입)(업)?$',
    r'잡화\s*도소매',
    # Boring finance
    r'대부업', r'대부중개', r'담보대출',
]


def is_excluded(business):
    for pat in EXCLUDE_PATTERNS:
        if re.search(pat, business):
            return True
    return False


def score_startup(company):
    """Score startup potential. Returns (score, category, reasons)."""
    biz = company['business']
    name = company['name']
    capital = company['capital']
    score = 0
    reasons = []
    category = ''

    if is_excluded(biz):
        return 0, '', ['제외 업종']

    # Tier 1 keywords
    t1_matches = [kw for kw in TIER1_KEYWORDS if kw.lower() in biz.lower() or kw.lower() in name.lower()]
    if t1_matches:
        score += 25
        reasons.append(f"핵심키워드: {t1_matches[0]}")
        if len(t1_matches) > 1:
            score += len(t1_matches) * 5
            reasons.append(f"+{len(t1_matches)-1}개 추가")

    # Tier 2 keywords
    t2_matches = [kw for kw in TIER2_KEYWORDS if kw.lower() in biz.lower() or kw.lower() in name.lower()]
    if t2_matches and not t1_matches:
        score += 10
        reasons.append(f"부키워드: {t2_matches[0]}")

    # Capital: sweet spot for startups is 50M~3B KRW
    if 50 <= capital <= 3000:
        score += 10
        reasons.append(f"자본금 {capital}백만")
    elif capital > 3000:
        score += 5
        reasons.append(f"고자본 {capital}백만")

    # English/creative name
    if re.search(r'[a-zA-Z]', name):
        score += 5
        reasons.append("영문 상호")

    # Name patterns
    for pat, label in [('테크$', '테크'), ('랩$', '랩'), ('스튜디오', '스튜디오'),
                        ('이노', '이노'), ('벤처', '벤처'), ('파트너스', '파트너스')]:
        if re.search(pat, name):
            score += 3
            reasons.append(f"상호: {label}")
            break

    # Sector penalty
    if company['sector'] in ['건설', '건자재']:
        score -= 15

    # Categorize
    biz_lower = biz.lower()
    if any(k in biz_lower for k in ['인공지능', 'ai', '머신러닝', '딥러닝', 'llm']):
        category = 'AI/ML'
    elif any(k in biz_lower for k in ['바이오', '헬스케어', '의료', '디지털헬스']):
        category = 'Bio/Health'
    elif any(k in biz_lower for k in ['플랫폼', '이커머스', '전자상거래', '마켓플레이스']):
        category = 'Platform'
    elif any(k in biz_lower for k in ['블록체인', '핀테크', '가상자산', 'defi', 'web3']):
        category = 'Blockchain/Fintech'
    elif any(k in biz_lower for k in ['소프트웨어', '앱', '클라우드', 'saas', 'api']):
        category = 'Software'
    elif any(k in biz_lower for k in ['게임', '콘텐츠', '미디어', '영상', '스트리밍']):
        category = 'Content/Media'
    elif any(k in biz_lower for k in ['에듀', '교육', '이러닝']):
        category = 'EdTech'
    elif any(k in biz_lower for k in ['에너지', '태양광', '배터리', '수소', '전기차']):
        category = 'CleanTech'
    elif any(k in biz_lower for k in ['로봇', 'iot', '드론', '자율주행', '스마트']):
        category = 'DeepTech'
    elif any(k in biz_lower for k in ['데이터', '빅데이터']):
        category = 'Data'
    elif any(k in biz_lower for k in ['투자', '자산', '결제', '페이', '금융']):
        category = 'Finance'
    elif any(k in biz_lower for k in ['솔루션', '컨설팅', '마케팅', '광고']):
        category = 'Service'
    else:
        category = 'Other'

    return max(0, score), category, reasons


def main():
    files = [
        ('data/last_week.xls', '3/20~3/26'),
        ('data/this_week.xls', '3/27~4/2'),
    ]

    for filepath, period in files:
        companies = parse_xls(filepath)

        scored = []
        for c in companies:
            score, cat, reasons = score_startup(c)
            if score >= 25:  # Only Tier 1 matches
                c['score'] = score
                c['category'] = cat
                c['reasons'] = reasons
                scored.append(c)

        scored.sort(key=lambda x: (-x['score'], x['category']))

        print(f"\n{'='*80}")
        print(f"  {period} | 총 {len(companies)}개 신설법인 → 스타트업 후보 {len(scored)}개")
        print(f"{'='*80}")

        # Group by category
        from collections import defaultdict
        by_cat = defaultdict(list)
        for s in scored:
            by_cat[s['category']].append(s)

        for cat in sorted(by_cat.keys()):
            items = by_cat[cat]
            print(f"\n  ── {cat} ({len(items)}개) ──")
            for s in items:
                cap_str = f"{s['capital']}백만" if s['capital'] else "미상"
                biz_short = s['business'][:55]
                print(f"    [{s['score']:2d}] {s['name']:<22} | {cap_str:>8} | {s['region']:<4} | {biz_short}")

        print()


if __name__ == '__main__':
    main()
