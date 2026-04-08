#!/usr/bin/env python3
"""Parse MK 신설법인 XLS files and filter for potential startups."""

import xlrd
import json
import re
import sys

# Startup-likely business keywords (Korean)
STARTUP_KEYWORDS = [
    # Tech / Software
    '소프트웨어', '플랫폼', '앱', '어플', '인공지능', 'AI', '데이터', '클라우드',
    '블록체인', '핀테크', '가상자산', '암호화폐', 'NFT', '메타버스', 'SaaS',
    '빅데이터', '머신러닝', '딥러닝', 'IoT', '사물인터넷', '로봇', '자율주행',
    '드론', 'IT', '정보기술', '디지털', '온라인', '이커머스', '전자상거래',
    '모바일', '인터넷', '웹', '게임', '콘텐츠', '미디어', '영상', '스트리밍',
    # Bio / Health
    '바이오', '헬스케어', '의료기기', '진단', '신약', '제약', '유전자', '세포',
    '줄기세포', '헬스', '웰니스', '디지털헬스', '원격의료', '테라피',
    # Clean / Energy
    '신재생에너지', '태양광', '태양열', '풍력', '수소', '전기차', '배터리',
    '에너지', '탄소', 'ESG', '친환경', '그린', '리사이클', '업사이클',
    # Commerce / Service
    '커머스', '구독', '큐레이션', '커스터마이징', '개인화', 'D2C', 'B2B',
    '마켓플레이스', '공유', '중개', '매칭', 'O2O',
    # Food / Agri
    '푸드테크', '식품기술', '대체육', '배양육', '스마트팜', '어그테크',
    # Education
    '에듀테크', '교육기술', '이러닝', '온라인교육',
    # Finance
    '자산관리', '투자', '보험기술', '인슈어테크', '결제', '페이',
    # Others
    '솔루션', '컨설팅', '마케팅', '광고', '브랜딩', '디자인',
    '스튜디오', '랩', '연구소', '테크', '이노베이션', '벤처',
]

# Keywords that suggest NOT a startup (traditional businesses)
NON_STARTUP_KEYWORDS = [
    '건축공사', '토목공사', '철거', '인테리어공사', '도장공사', '방수공사',
    '전기공사', '설비공사', '소방공사', '조경공사', '포장공사', '측량',
    '부동산임대', '부동산매매', '부동산중개', '임대업', '건물관리',
    '화물운송', '화물운수', '용달', '퀵서비스', '택배',
    '음식점', '식당', '카페', '커피', '베이커리', '정육', '반찬',
    '미용실', '네일', '피부관리', '세차', '세탁', '청소',
    '철물점', '문구', '잡화', '슈퍼마켓', '편의점',
    '노래방', '당구장', '목욕탕', '찜질방',
    '학원', '과외', '입시',  # traditional education, not edtech
    '약국', '한의원', '치과', '병원',
    '농업', '축산', '어업', '임업',
    '시멘트', '골재', '모래', '자갈', '석재',
]

# Name patterns suggesting startup (English/creative names)
STARTUP_NAME_PATTERNS = [
    r'[a-zA-Z]',          # Contains English characters
    r'테크$', r'랩$', r'스튜디오$',
    r'이노', r'벤처', r'파트너스',
]


def parse_xls(filepath):
    """Parse the MK 신설법인 XLS file and return list of companies.

    Column layout:
    A (0): marker (▷ for company, ◆ for sector header)
    B (1): 상호 (company name) or sector name or region
    C (2): 대표자 (CEO)
    D (3): 자본금 (capital, 백만원)
    E (4): 주요사업 (business description)
    F (5): 주소 (address)
    """
    wb = xlrd.open_workbook(filepath, encoding_override='cp949')
    sh = wb.sheet_by_index(0)

    companies = []
    current_region = ""
    current_sector = ""

    for row_idx in range(sh.nrows):
        row = [str(sh.cell_value(row_idx, col)).strip() for col in range(sh.ncols)]
        marker = row[0]  # ▷ or ◆ or empty

        # Skip empty rows
        if not any(row):
            continue

        # Row 0: Region header (e.g., "서 울")
        if '서' in row[1] and '울' in row[1] and not row[2]:
            current_region = row[1].replace(' ', '')
            continue
        # Other regions
        region_map = {'부 산': '부산', '대 구': '대구', '인 천': '인천',
                      '광 주': '광주', '대 전': '대전', '울 산': '울산',
                      '경 기': '경기', '강 원': '강원', '충 북': '충북',
                      '충 남': '충남', '전 북': '전북', '전 남': '전남',
                      '경 북': '경북', '경 남': '경남', '제 주': '제주',
                      '세 종': '세종'}
        for k, v in region_map.items():
            if k in row[1] and not row[2]:
                current_region = v
                break
        else:
            pass  # not a region row

        # Header / unit row
        if '단위' in row[1] or ('상호' in row[1] and '대표자' in row[2]):
            continue

        # Sector header: marker is ◆
        if marker == '◆':
            sector_text = row[1].replace(' ', '').strip()
            if sector_text:
                current_sector = sector_text
            continue

        # Company row: marker is ▷
        if marker == '▷':
            name = row[1].strip()
            ceo = row[2].strip()

            capital_raw = row[3].strip()
            try:
                capital = int(float(capital_raw)) if capital_raw else 0
            except:
                capital = 0

            business = row[4].strip()
            address = row[5].strip()

            if not name:
                continue

            companies.append({
                'name': name,
                'ceo': ceo,
                'capital': capital,  # in 백만원
                'business': business,
                'address': address,
                'region': current_region,
                'sector': current_sector,
            })

    return companies


def score_startup_potential(company):
    """Score a company's startup potential (0-100). Higher = more likely startup."""
    score = 0
    reasons = []

    biz = company['business']
    name = company['name']
    capital = company['capital']

    # 1. Business keyword matching
    biz_lower = biz.lower()
    name_lower = name.lower()

    for kw in STARTUP_KEYWORDS:
        kw_lower = kw.lower()
        if kw_lower in biz_lower or kw_lower in name_lower:
            score += 15
            reasons.append(f"키워드: {kw}")
            break  # Count once for keyword match

    # Additional keyword matches (bonus)
    keyword_count = sum(1 for kw in STARTUP_KEYWORDS if kw.lower() in biz_lower or kw.lower() in name_lower)
    if keyword_count > 1:
        score += min(keyword_count * 5, 20)
        reasons.append(f"복수 키워드 {keyword_count}개")

    # 2. Non-startup penalty
    for kw in NON_STARTUP_KEYWORDS:
        if kw in biz:
            score -= 30
            reasons.append(f"전통업종: {kw}")
            break

    # 3. Capital analysis
    # Startups typically: 1-500 백만원 (1억~5억)
    # Very small (1-5 백만원) could be shell company or sole proprietor
    if 50 <= capital <= 5000:
        score += 15
        reasons.append(f"적정 자본금: {capital}백만")
    elif capital > 5000:
        score += 5
        reasons.append(f"고자본금: {capital}백만")
    elif capital < 10 and capital > 0:
        score -= 5

    # 4. Name analysis
    has_english = bool(re.search(r'[a-zA-Z]', name))
    if has_english:
        score += 10
        reasons.append("영문 포함 상호")

    for pattern in STARTUP_NAME_PATTERNS[1:]:  # Skip the English one
        if re.search(pattern, name):
            score += 5
            reasons.append(f"상호 패턴: {pattern}")
            break

    # 5. Sector bonus
    tech_sectors = ['기타', '기계금속']
    if company['sector'] in tech_sectors:
        score += 5

    # Construction/건자재 penalty
    if company['sector'] in ['건설', '건자재']:
        score -= 10

    return max(0, min(100, score)), reasons


def filter_startups(companies, threshold=20):
    """Filter and rank companies by startup potential."""
    results = []
    for c in companies:
        score, reasons = score_startup_potential(c)
        if score >= threshold:
            c['startup_score'] = score
            c['reasons'] = reasons
            results.append(c)

    results.sort(key=lambda x: x['startup_score'], reverse=True)
    return results


def main():
    files = [
        ('data/last_week.xls', '3/20~3/26'),
        ('data/this_week.xls', '3/27~4/2'),
    ]

    all_results = {}

    for filepath, period in files:
        print(f"\n{'='*70}")
        print(f"  기간: {period} ({filepath})")
        print(f"{'='*70}")

        companies = parse_xls(filepath)
        print(f"  총 신설법인: {len(companies)}개")

        startups = filter_startups(companies, threshold=20)
        print(f"  스타트업 후보: {len(startups)}개")

        all_results[period] = {
            'total': len(companies),
            'startups': startups,
        }

        print(f"\n  {'순위':>4} | {'점수':>4} | {'상호':<20} | {'자본금':>8} | {'주요사업':<40} | 이유")
        print(f"  {'-'*4} | {'-'*4} | {'-'*20} | {'-'*8} | {'-'*40} | {'-'*20}")

        for i, s in enumerate(startups, 1):
            name_display = s['name'][:18]
            biz_display = s['business'][:38]
            reasons_str = ', '.join(s['reasons'][:3])
            capital_str = f"{s['capital']}백만"
            print(f"  {i:>4} | {s['startup_score']:>4} | {name_display:<20} | {capital_str:>8} | {biz_display:<40} | {reasons_str}")

    # Summary
    print(f"\n\n{'='*70}")
    print(f"  종합 요약")
    print(f"{'='*70}")
    for period, data in all_results.items():
        print(f"\n  [{period}]")
        print(f"  총 {data['total']}개 신설법인 중 스타트업 후보 {len(data['startups'])}개")
        if data['startups']:
            top = data['startups'][:5]
            print(f"  Top 5:")
            for i, s in enumerate(top, 1):
                print(f"    {i}. {s['name']} (점수: {s['startup_score']}) - {s['business'][:50]}")


if __name__ == '__main__':
    main()
