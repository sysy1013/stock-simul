import os
import csv
import pandas as pd
import random

# 재실행해도 항상 같은 시나리오 데이터가 나오도록 시드 고정
random.seed(42)

# 2. 40개 실제 기업 풀
company_pool = {
    '삼성전자': {'industry': '반도체/전자', 'price': 80000}, 'SK하이닉스': {'industry': '반도체/전자', 'price': 170000}, 'LG전자': {'industry': '반도체/전자', 'price': 100000}, '삼성전기': {'industry': '반도체/전자', 'price': 150000}, 'DB하이텍': {'industry': '반도체/전자', 'price': 50000},
    '네이버': {'industry': 'IT/소프트웨어', 'price': 190000}, '카카오': {'industry': 'IT/소프트웨어', 'price': 50000}, '엔씨소프트': {'industry': 'IT/소프트웨어', 'price': 200000}, '크래프톤': {'industry': 'IT/소프트웨어', 'price': 250000}, '삼성SDS': {'industry': 'IT/소프트웨어', 'price': 160000},
    '셀트리온': {'industry': '바이오/제약', 'price': 180000}, '삼성바이오로직스': {'industry': '바이오/제약', 'price': 800000}, '유한양행': {'industry': '바이오/제약', 'price': 70000}, '한미약품': {'industry': '바이오/제약', 'price': 300000}, 'SK바이오팜': {'industry': '바이오/제약', 'price': 90000},
    '하이브': {'industry': '엔터/미디어', 'price': 200000}, 'JYP Ent.': {'industry': '엔터/미디어', 'price': 70000}, '에스엠': {'industry': '엔터/미디어', 'price': 80000}, '와이지엔터테인먼트': {'industry': '엔터/미디어', 'price': 40000}, '스튜디오드래곤': {'industry': '엔터/미디어', 'price': 45000},
    '현대차': {'industry': '자동차/모빌리티', 'price': 240000}, '기아': {'industry': '자동차/모빌리티', 'price': 110000}, '현대모비스': {'industry': '자동차/모빌리티', 'price': 250000}, 'HL만도': {'industry': '자동차/모빌리티', 'price': 35000}, '한온시스템': {'industry': '자동차/모빌리티', 'price': 6000},
    'KB금융': {'industry': '금융/은행', 'price': 70000}, '신한지주': {'industry': '금융/은행', 'price': 50000}, '하나금융지주': {'industry': '금융/은행', 'price': 60000}, '우리금융지주': {'industry': '금융/은행', 'price': 15000}, '카카오뱅크': {'industry': '금융/은행', 'price': 25000},
    'CJ제일제당': {'industry': '식음료/소비재', 'price': 300000}, '아모레퍼시픽': {'industry': '식음료/소비재', 'price': 130000}, '농심': {'industry': '식음료/소비재', 'price': 400000}, '오리온': {'industry': '식음료/소비재', 'price': 90000}, 'LG생활건강': {'industry': '식음료/소비재', 'price': 350000},
    'LG에너지솔루션': {'industry': '에너지/화학', 'price': 400000}, 'SK이노베이션': {'industry': '에너지/화학', 'price': 120000}, 'S-Oil': {'industry': '에너지/화학', 'price': 75000}, '한화솔루션': {'industry': '에너지/화학', 'price': 30000}, '포스코퓨처엠': {'industry': '에너지/화학', 'price': 300000}
}

# 3. 산업군별 주력 타겟층 매핑
target_demographics = {
    '반도체/전자': ['B2B 기업고객', '20~40대 남성', '30~50대 직장인'],
    'IT/소프트웨어': ['10~20대 학생', '20~30대 직장인', '전 연령층'],
    '바이오/제약': ['50~70대 노년층', '40~60대 장년층', '의료기관'],
    '엔터/미디어': ['10~20대 여성', '10~30대 글로벌 팬덤', '20~30대 직장인'],
    '자동차/모빌리티': ['30~50대 남성', '20~30대 사회초년생', '40~60대 가족단위'],
    '금융/은행': ['30~50대 직장인', '20~30대 청년층', '고자산가 그룹'],
    '식음료/소비재': ['20~30대 여성', '1인 가구', '30~40대 주부'],
    '에너지/화학': ['B2B 기업고객', '정부 및 공공기관', '해외 수출시장']
}

# 4. 산업군별 주력 제품군 매핑 (새로 추가됨)
products_by_industry = {
    '반도체/전자': ['스마트폰', '프리미엄 가전', '메모리 반도체', '웨어러블 기기'],
    'IT/소프트웨어': ['클라우드 서비스', '모바일 게임', '구독형 솔루션', '광고 플랫폼'],
    '바이오/제약': ['전문의약품', '건강기능식품', '의료기기', '백신/치료제'],
    '엔터/미디어': ['앨범/굿즈', '콘서트 티켓', '영상 콘텐츠(드라마)', '팬클럽 구독'],
    '자동차/모빌리티': ['친환경 전기차', 'SUV/레저용', '세단', '차량용 부품'],
    '금융/은행': ['예적금 상품', '주택담보대출', '신용카드', '투자/자산관리'],
    '식음료/소비재': ['가정간편식(HMR)', '스낵/제과', '음료/주류', '뷰티/화장품'],
    '에너지/화학': ['2차전지 배터리', '석유/정유', '친환경 플라스틱', '태양광 패널']
}

sales_channels = ['오프라인 매장', '자사몰(온라인)', '종합 오픈마켓', 'B2B/기관 영업']
quarters = [f"20{y}_{q}Q" for y in range(23, 26) for q in range(1, 5)] # 과거 3개년 12개 분기 (2023 1분기 ~ 2025 4분기) - 시뮬레이션은 2026 1분기부터 시작

# 5. Orange3 호환 헤더 정의 (1행: 컬럼명 / 2행: 자료형 / 3행: 역할)
#    - 자료형: continuous(연속형 숫자), discrete(범주형), string(문자열)
#    - 역할: 빈칸(feature=입력 변수), meta(참고 정보), class(분류 분석의 타겟 변수)
ORANGE3_TYPES = {
    'Year_Quarter': 'string',
    'Year_Month': 'string',
    '매출액(억원)': 'continuous',
    '영업이익(억원)': 'continuous',
    '시장점유율(%)': 'continuous',
    '주요 타겟층': 'discrete',
    '주력 판매 제품': 'discrete',
    '주요 판매 채널': 'discrete',
    '재고율(%)': 'continuous',
    'ROAS(%)': 'continuous',
    '고객 이탈률(%)': 'continuous',
    '해외매출비중(%)': 'continuous',
    '고객 만족도(점)': 'continuous',
    'ESG 평판(점)': 'continuous',
    '마케팅_위기_등급': 'discrete',
    '기업명': 'string',
    '산업군': 'discrete',
    '주가(원)': 'continuous',
    '현재주가(원)': 'continuous',
    '주가등락률(%)': 'continuous',
    'AI평가배수': 'continuous',
    '시장이슈': 'discrete',
    '마케팅전략': 'string',
    'AI코멘트': 'string',
}
ORANGE3_ROLES = {
    'Year_Quarter': 'meta',
    'Year_Month': 'meta',
    '기업명': 'meta',
    '마케팅전략': 'meta',
    'AI코멘트': 'meta',
    '마케팅_위기_등급': 'class',
}


def grade_marketing_risk(inventory_rate, churn_rate, roas, brand_score):
    """의사결정나무 학습의 정답지가 될 '마케팅_위기_등급'을 규칙 기반으로 산출.
    학생들이 분석 후 '왜 위험인지'를 지표로 역추적할 수 있도록 기준을 단순하게 유지."""
    risk = 0
    if inventory_rate >= 30:  # 재고 과잉
        risk += 1
    if churn_rate >= 15:      # 고객 이탈 심각
        risk += 1
    if roas < 120:            # 광고비 대비 수익 저조
        risk += 1
    if brand_score < 70:      # 브랜드 만족도 부진
        risk += 1
    if risk == 0:
        return '안전'
    if risk == 1:
        return '주의'
    return '위험'


def save_orange3_csv(df, file_path):
    """Orange3 전용 3줄 헤더(컬럼명/자료형/역할)를 붙여 CSV로 저장"""
    cols = list(df.columns)
    with open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(cols)
        writer.writerow([ORANGE3_TYPES.get(c, 'string') for c in cols])
        writer.writerow([ORANGE3_ROLES.get(c, '') for c in cols])
        for row in df.itertuples(index=False):
            writer.writerow(row)


def generate_all_data():
    if not os.path.exists('data'):
        os.makedirs('data')

    print("40개 기업의 과거 3개년(2023~2025, 12분기) 초기 분석 데이터 생성을 시작합니다...")

    for company, info in company_pool.items():
        industry = info['industry']
        base_price = info['price']

        # 기업 규모(주가)에 비례하는 기본 매출액 세팅 (단위: 억 원)
        base_sales = int(base_price * random.uniform(0.8, 1.5))

        data_list = []

        for quarter in quarters:
            # 분기별 무작위 변동성 부여
            sales = int(base_sales * random.uniform(0.85, 1.2))
            operating_profit = int(sales * random.uniform(0.05, 0.25)) # 영업이익은 매출의 5~25%
            market_share = round(random.uniform(5.0, 45.0), 1)

            # 속성 무작위 추출
            target = random.choice(target_demographics[industry])
            top_product = random.choice(products_by_industry[industry])
            channel = random.choice(sales_channels)

            # 심화 마케팅 지표 생성
            inventory_rate = round(random.uniform(5.0, 40.0), 1)  # 재고율 5~40%
            brand_score = random.randint(60, 98)                  # 고객 만족도(브랜드) 60~98점
            roas = random.randint(80, 350)                        # ROAS 80~350%
            churn_rate = round(random.uniform(2.0, 25.0), 1)      # 이탈률 2~25%
            overseas_sales = random.randint(5, 85)                # 해외 매출 비중 5~85%
            esg_score = random.randint(40, 95)                    # ESG 평판 40~95점

            data_list.append({
                'Year_Quarter': quarter,
                '매출액(억원)': sales,
                '영업이익(억원)': operating_profit,
                '시장점유율(%)': market_share,
                '주요 타겟층': target,
                '주력 판매 제품': top_product,
                '주요 판매 채널': channel,
                '재고율(%)': inventory_rate,
                'ROAS(%)': roas,
                '고객 이탈률(%)': churn_rate,
                '해외매출비중(%)': overseas_sales,
                '고객 만족도(점)': brand_score,
                'ESG 평판(점)': esg_score,
                '마케팅_위기_등급': grade_marketing_risk(inventory_rate, churn_rate, roas, brand_score)
            })

            # 다음 분기를 위해 기본 매출액 미세 조정 (성장 또는 하락 추세 반영)
            base_sales = int(base_sales * random.uniform(0.95, 1.08))

        df = pd.DataFrame(data_list)
        file_path = f"data/{company}_초기분석데이터.csv"
        save_orange3_csv(df, file_path)

    print("총 40개 기업의 CSV 파일이 'data' 폴더에 성공적으로 생성되었습니다!")
    print("이제 터미널에서 'streamlit run app.py'를 실행하여 대시보드를 켜주세요.")


if __name__ == "__main__":
    generate_all_data()