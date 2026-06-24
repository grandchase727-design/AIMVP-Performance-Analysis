# CLAUDE.md — AIMVP 포트폴리오 분석 프로젝트

이 파일은 Claude Code가 본 작업 디렉토리에서 작업할 때 사용할 컨텍스트입니다.

---

## 프로젝트 개요

AIMVP (Asset Investment Multi-Vehicle Portfolio) 펀드의 시그널 룰(50% SPY→QQQ 스왑) 유효성을
정량 검증하고 포트폴리오 전반 개선 사항을 분석. 현재 **Streamlit 인터랙티브 대시보드** 형태로
지속 확장 중.

- **분석 대상**: AIMVP 적극형(89.55% 주식) + 중립형(59.70% 주식) 2개 risk profile
- **분석 기간**: 2022-03-01 ~ 2026-06-01 (53개월+, 월간 리밸런싱, 6월 데이터까지 갱신됨)
- **시그널 분포**: Bull 21회 / Base 28회 / Bear 2회(적극) 또는 1회(중립)

---

## 핵심 데이터 소스

| 파일 | 역할 |
|---|---|
| `AIMVP 포트폴리오 비중추이.xlsx` | **원본 데이터** — 적극·중립 53개월+ 자산 비중·시그널 시계열 |
| Yahoo Finance (yfinance) | SPY, QQQ, ACWI, USDKRW + 40개 ETF + 팩터 proxy(VLUE/IWM/QUAL/MTUM/USMV) + 벤치 proxy(BNDW) — `auto_adjust=True` |

### Excel 시트 구조 (중요)
- 시트 1: `AIMVP포트폴리오현황(국가별자산배분이전)` (2022-03 ~ 2023-09)
  - 적극: 시그널 row 3, 날짜 row 5, 자산 rows 6-21
  - 중립: 시그널 row 29, 날짜 row 31, 자산 rows 33-48
- 시트 2: `AIMVP포트폴리오현황(국가별자산배분이후)` (2023-10 ~ 2026-06)
  - 적극: 시그널 row 4, 날짜 row 6, 자산 rows 8-48 (PHYL 추가 후 8-48)
  - 중립: 시그널 row 54, 날짜 row 56, 자산 rows 58-98
- 시트 3: `투자유니버스` — 40개 ETF 분류

---

## 🎛️ Streamlit 대시보드 (`aimvp_streamlit_app.py`)

### 페이지 구조
1. **📊 Portfolio** — 보유 종목·성과 + 카테고리 break-down + 리밸 효과 + 시그널 win/lose
2. **📈 Performance** — KPI / 누적성과 추이 / 배분별 추이 / 위험 지표
3. **🎯 Signal Validation** — Bull 시그널 통계 검정 + 슬리브 분해 + 하방 완충 분석

### 사이드바
- **🔄 Update 버튼** — 캐시 비우고 Report Date를 오늘로 재설정
- **Report Date** — 기준일 선택, 모든 분석 데이터가 이 날짜 이전으로 자동 필터링
- **Profile multiselect** (적극형/중립형), **Swap ratio slider**, **Benchmark toggle**

### 핵심 컴포넌트

#### Portfolio 페이지 (profile별 탭)
- 종목별 성과 테이블 (1D/1W/MTD return + 기여 bps)
- **구분별 성과** breakdown (광역주식/국가주식/채권/현금)
- **현재 vs 전월 비교** 테이블 (counter-factual 비중 효과 분리)
- **🎯 시그널 Win/Lose 추이 분석** (방법 1: alpha vs profile-specific 벤치, 방법 2: 시그널 의도 vs ACWI regime, 4-Quadrant expander)

#### Portfolio 페이지 (공통 섹션, 탭 하단)
- 📋 **Commentary** — Weekly(500자) / Monthly(1000자, 월말 영업일만, 완전 data-driven)
- 🌍 **국가별 배분** (비중/성과 sub-tab, ETF look-through, Top 10)
- 🏭 **섹터별 배분** (비중/성과 sub-tab, GICS 11 섹터)
- 🎯 **팩터별 배분** (MSCI Style 큐레이션 + Fama-French 회귀 2가지 방법)
- 💵 **채권 듀레이션·YTM 분석** (공통 metrics + 적극/중립 슬리브 비중만 분리)
- 📈 **리밸런싱 효과 추이** — 4가지 뷰 모드 (📅 달력 히트맵 / 🏆 시그널 성적표 / 🎨 시계열 막대 / 📈 누적 시계열+필터)
- 🔬 **자산 배분 효과 심층 분석** — 3가지 분석 (시그널×자산 / 광역주식 vs ACWI 100% / 전체 주식 vs ACWI 100%)

#### Performance 페이지
- KPI 카드 (WTD/MTD/YTD/ITD per profile)
- **📊 누적성과 추이** — 10개 기간 탭 (📅 MTD / 🔄 1M / 📊 QTD / 🔄 3M / 🔄 6M / 📈 YTD / 🔄 1Y / 📐 3Y/a / 📐 5Y/a / 🏁 ITD), 3Y/5Y/ITD는 연환산 CAGR 자동 표시
- **🧭 배분별 성과 추이** (국가/섹터/팩터, 같은 10개 기간 탭) + 상단 🔥 유니버스 비중 분포 Heatmap (3개 카테고리)
- **방법론 토글:** 💯 절대 수익률 (비중 무관) vs 💰 historical 기여 (시점별 실제 비중)
- 위험 지표 비교 + Alerts

#### Signal Validation 페이지
- Bull QQQ vs SPY 적중률 통계 검정
- 🧩 **슬리브별 리밸런싱 정합 분석** (주식/채권 분해 + 누적 알파 + 적중률)
- 🛡️ **주식 하방 완충 분석** — 방어 슬리브(채권+현금)의 헷지 효과, 산점도 + 막대 + 요약 표

### Profile-specific 벤치마크 (시그널 Win/Lose 평가용)
- **적극형:** 90% ACWI + 10% Cash (4%/y proxy)
- **중립형:** 60% ACWI + 15% BNDW (Bloomberg Global Aggregate proxy) + 15% KIS 종합채권 (3.5%/y proxy) + 10% Cash

---

## 방법론 핵심 원칙

### ⚠️ Forward 1-month 기준만 사용
AIMVP는 **월간 리밸런싱**이므로 보유 기간이 1개월. 3m 등 보유 기간 초과 호라이즌은
실제 전략이 경험하지 않는 가상 시나리오이므로 **검증에서 제외**.

### 시뮬레이션 방식
1. 원본 Excel에서 각 리밸런싱 시점 자산 비중 그대로 읽기
2. Bull 시그널일 때만 SPY의 r%를 QQQ로 이동 (예: 50%면 SPY 42% → 21%, QQQ 11% → 32%)
3. 나머지 자산은 그대로 — 전체 합 100% 유지
4. 일별 포트폴리오 수익률 = Σ(자산_i 비중 × 자산_i 일간수익률)
5. **연속 체인 (중요):** 리밸 경계에서 수익률 누락 없이 일별 시계열을 연속 연결.
   리밸 당일 수익률(전일→리밸일)은 **직전 비중**으로 산출(rebalance-at-close), 다음
   거래일부터 새 비중 적용. (구버전은 윈도우별 `pct_change().dropna()` 후 concat 하여
   매 경계 1일을 누락 → 누적 과소계상. 2026-06-23 정정.)

### KR 채권·현금 가격 처리
yfinance에 없는 KR 자산은 proxy 상수 수익률로 처리:
```python
CASH_DAILY   = (1.04)**(1/252) - 1   # 연 4%
KRBOND_DAILY = (1.035)**(1/252) - 1  # 연 3.5%
```

### Counter-factual 비중 효과 분리 (compute_comparison 등)
**리밸 효과 = (현재 비중 - 전월 비중) × 현재 MTD return** — 같은 return으로 양쪽 평가해 **비중 의사결정 효과만** 분리.

### 시그널 독립성 존중 (월별 분석)
시그널은 매월 독립적으로 산출되므로 누적 시계열은 누적 의존성을 시각적으로 강조함. 월 독립성을 존중하려면:
- 📅 달력 히트맵, 🏆 시그널 성적표(KPI 카드), 🎨 시계열 막대(누적선 제거) 권장
- 통계는 mean/median/distribution 기반 (Wilcoxon, bootstrap CI)

### 통계 검정
**비모수 검정 우선** (수익률 분포 비대칭으로 t-test는 부적합):
- Wilcoxon signed-rank test (단측/양측)
- Block bootstrap CI (블록=3, 10,000회)
- Permutation test (10,000회)

### Look-through 매핑 정확도
| 차원 | 출처 | 정확도 |
|---|---|---|
| 국가 (단일국가 ETF) | ETF 정의 | ✅ 100% 정확 |
| 국가 (ACWI/EFA/VWO) | iShares/Vanguard fact sheet 2025 후반 스냅샷 + 수작업 라운딩 | ⚠️ ±1~2%p |
| 섹터 (GICS 11) | iShares/Vanguard/Invesco fact sheet 참고 큐레이션 | ⚠️ ±2~5%p |
| 팩터 (MSCI Style) | MSCI Style Box / Morningstar 일반 분석 기반 추정 | 🚨 가장 주관적 |
| 팩터 (Fama-French) | 일별 수익률을 long-short proxy 회귀 (VLUE/IWM/QUAL/MTUM/USMV-SPY) | ✅ 데이터 기반 |

---

## 핵심 결과 (재현 가능한 수치)

### Bull 시그널 QQQ vs SPY 적중률 (1m, n=21)
- 적중률 **61.9% (13/21)**, 평균 초과 **+0.75%p**
- W (적중 시) +1.69%p, L (미적중 시) -0.77%p, W/L ratio **2.19**
- Wilcoxon p = **0.032** ✓ 5% 유의
- Bootstrap 95% CI [+0.082, +1.273] — 0 미포함

### Base 시그널 동일 분석 (n=28)
- 적중률 **42.9% (12/28)** — 코인플립 미만
- Wilcoxon p = **0.469** ✗ 유의 미달
- ★ Base 확장 부적합

### Portfolio-level 누적 알파 (50% 룰) — 53개월, 2026-06-23 종가 기준
| Risk Profile | 실제(0%) | 50% 룰 | 알파 | Sharpe 변화 |
|---|---|---|---|---|
| **적극형** | +60.14% | +63.19% | **+3.04%p** | 0.92 → 0.94 |
| **중립형** | +33.96% | +35.67% | **+1.71%p** | 0.84 → 0.85 |

→ 중립이 적극의 약 56% 알파 (SPY/QQQ 비중 비례)

> ⚠️ **2026-06-23 엔진 정정:** `simulate_portfolio`가 월 구간을 배타적으로 잘라
> `pct_change().dropna()` 후 concat 하여 **매 리밸 경계의 turn-of-month 1일 수익률(~52일)을
> 누락**, 누적을 체계적으로 과소계상하던 버그를 수정(연속 일별 체인, rebalance-at-close).
> 단일자산 100% 보유 시 **직접 가격비율과 정확히 일치(누락 0일)** 검증 완료.
> 위 표는 **정정 후** 수치이며, 구버전 값(적극 +46.43%/+3.13%p, 중립 +27.87%/+1.82%p)을 대체함.
> 알파의 방향·상대크기 결론(스왑 룰 (+)알파, 중립≈적극 절반)은 유지됨.
> ※ 수치는 종가 기준 일별 변동 — 본 표는 2026-06-23 종가(직전일 −2% 급락 반영) 스냅샷.
> 재현: `py generate_report.py` (정정 차트 + 요약 docx 동시 생성).

### SPY vs QQQ 장기 (1999-2026)
- SPY +820% / CAGR 8.50% / MaxDD -56%
- QQQ +1,539% / CAGR 10.83% / MaxDD -83%
- QQQ 우세 46%, SPY 우세 19%, 중립 35%

---

## 생성된 리포트 파일 (정적 분석 결과)

| 파일 | 내용 |
|---|---|
| `AIMVP_SPY_QQQ_검증_5단계.docx` | **메인 리포트** — 적극+중립 통합 검증, 9개 차트, 6개 개선 영역 |
| `AIMVP_시그널_구조_재분석.docx` | Bull/Base/Bear 시그널 구조 분석 |
| `AIMVP_정량_백테스트_결과.docx` | 초기 5단계 정량 백테스트 |

### 차트 자산 (`chart_*.png`)
- `chart0_real_sim.png` — 적극형 53개월 시계열 (스왑 6단계)
- `chart1_cumulative.png` — 적극형 Bull 이벤트 순번별 누적
- `chart2_per_event.png`, `chart3_distribution.png`, `chart4_sharpe_curve.png` — 슬리브 분석
- `chart_base_sim.png`, `chart_base_cumulative.png` — 적극형 Base
- `chart_neutral_bull_sim.png`, `chart_neutral_bull_cumulative.png` — 중립형 Bull
- `chart_neutral_base_sim.png`, `chart_neutral_base_cumulative.png` — 중립형 Base
- `chart_regime.png`, `chart_rolling_horizons.png` — SPY vs QQQ 장기 regime
- `chart_sharpe_compare.png` — 스왑비율별 Sharpe 비교

---

## 코드 컨벤션

### 한국어 폰트 (matplotlib, docx)
```python
import matplotlib as mpl
mpl.rcParams['font.family'] = 'AppleGothic'
mpl.rcParams['axes.unicode_minus'] = False
```

### Plotly 차트 표준 (대시보드)
- 시그널 색상: Bull `#10B981` (초록) / Base `#F59E0B` (앰버) / Bear `#DC2626` (빨강)
- Profile 색상: 적극형 `#1F3A68` / 중립형 `#C48D43`
- 벤치마크: SPY `#9CA3AF` / ACWI `#6B7280` (점선)
- Bull 시그널 음영: `#FEE2E2` opacity 0.3
- 0 reference line: `line_dash='dash', line_color='black', opacity=0.5`
- 모든 `st.plotly_chart`에 고유 `key=` 필수 (StreamlitDuplicateElementId 방지)

### docx 생성 표준
- 폰트: 맑은 고딕
- 페이지 여백: 2.0 cm
- 표 스타일: `Light Grid Accent 1`
- 색상: 헤딩 `#1F3A68`(L1), `#2E5A88`(L2), 하이라이트 셀 `#FFF3CD`

### Ticker 매핑 (TICKER_MAP)
Excel의 ticker 문자열 → yfinance 심볼:
- `'spy us equity'` → `'SPY'`
- `'a357870'`, `'357870 ks equity'` → `'_CASH_KR'`
- `'a114820'`, `'a451540'` → `'_KRBOND'`
- 등 32개 US ETF + 3개 KR proxy + 팩터 proxy 5개 + BNDW

### 카테고리 상수 (slv 분해용)
```python
EQUITY_SLEEVE_CATS    = {'광역 주식', '국가 주식'}
BOND_SLEEVE_CATS      = {'채권'}
DEFENSIVE_SLEEVE_CATS = {'채권', '현금'}
BROAD_EQ_TICKERS      = {'spy/qqq/iwm/acwi/efa/vwo us equity'}
FACTOR_ORDER          = ['Value', 'Growth', 'Quality', 'Momentum', 'Size', 'Low Vol']
FF_FACTOR_ORDER       = ['MKT', 'Value', 'Size', 'Quality', 'Momentum', 'Low Vol']
```

---

## AIMVP 포트폴리오 6개 개선 영역 (Top 5 우선순위)

| 순위 | 항목 | 영역 |
|---|---|---|
| 1 | Bear 시그널 민감도 강화 (4년 중 2회만 발효) | 시그널 |
| 2 | 채권 다변화 효과 정상화 (HYG β=0.34, 위기 시 동반 손실) | 자산배분 |
| 3 | Bull 50% 스왑 룰 약관 명시 (적극+중립) + Base 확장 금지 | 룰 정합성 |
| 4 | 거래비용·규제 컴플라이언스 점검 (QQQ 0.20% vs SPY 0.09%) | 운영 |
| 5 | 표본 외 stress test 정례화 (닷컴/금리쇼크 미관측) | 위기 대응 |

---

## 작업 시 주의사항

1. **표본 외 위험**: 2022-2026 표본은 AI rally 편향 — 100% 풀스왑 결과는 regime 편향
   가능. 결론은 항상 한계 명시 ("닷컴/금리쇼크형 SPY-우세 regime 미관측").

2. **검정력 한계**: n=21 Bull 표본은 검정력 53%. 추가 trackrecord 12~18개월 누적 후
   재검증 권고.

3. **거래비용 미반영**: 백테스트는 그로스. QQQ-SPY 보수 차이 +11bps/년 + 양도세
   22% + 스프레드 5-10bps — net alpha 측정 필요.

4. **Bull/Base 적중률은 시그널 동일이므로 적극·중립에서 동일**. 단,
   portfolio-level 알파는 SPY/QQQ 비중에 비례해 다름.

5. **임시 파일 정리**: 분석 스크립트는 `_` 프리픽스로 명명 후 작업 완료 시 삭제.

6. **시그널 독립성**: 시그널은 매월 독립 산출 — 누적 시계열은 시각화 보조용으로만,
   통계적 결론은 월별 평균/distribution 기반.

7. **Look-through는 정적 매핑** (단일 스냅샷 + 수작업 라운딩). 광역 ETF 3개의
   월별 fact sheet 시계열 평균 보강 작업 backlog에 있음.

8. **CLAUDE.md 갱신**: 대시보드에 신규 섹션·tab·방법론을 추가할 때마다 본 파일도
   동기화 (사용자 요청 시 진행).

---

## 미해결 후속 작업 (Backlog)

- [ ] 거래비용·세금 정량 측정 (net alpha 산정)
- [ ] Bear 시그널 임계치 재캘리브레이션 (VIX + 신용 스프레드 결합)
- [ ] 펀드 약관 정합성 점검 (단일 ETF 비중 한도)
- [ ] 펀드 약관에 Bull 50% 스왑 룰 공식 명시
- [ ] Walk-forward / Deflated Sharpe Ratio 검증
- [ ] **iShares fact sheet 6개월 자동 다운로드 → LOOKTHROUGH 정밀화** (ACWI/EFA/VWO, 월별 시계열 평균)
- [ ] Morningstar Factor Box / MSCI Barra API 연동 → 팩터 정밀도 향상
- [ ] Pilot → 점진 도입 프레임워크 설계
- [ ] 거래비용 반영 옵션 토글 (대시보드 내)

---

_Last updated: 2026-06-23 (simulate_portfolio 연속체인 엔진 정정 — 누적/Sharpe/알파 수치 갱신)_
