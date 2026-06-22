"""AIMVP Commentary Agent Swarm — Stochastic (no LLM, no API key, non-deterministic).

각 agent가 phrase pool에서 random 선택 → 매 생성마다 narrative 변형.
숫자·facts는 고정 (실제 데이터 그대로), 표현·연결어·강조어만 variance.

Architecture (unchanged):
    [macro / country / bond / fx / signal] specialists (parallel)
        │
        ▼
    [synthesizer] — 5 drafts를 단일 narrative로 통합 + 헤더 (sentence order 변형)
        │
        ▼
    [editor] — 분량 검증 + 어미·강조 변형
        │
        ▼
    Final commentary

각 generation마다 다른 텍스트, 같은 facts.
"""

from __future__ import annotations
import random
from dataclasses import dataclass, field
from typing import Optional
from concurrent.futures import ThreadPoolExecutor


# ─── Context ─────────────────────────────────────────────────────────────
@dataclass
class CommentaryContext:
    kind: str
    period_label: str
    period_range: str
    signal: str
    signal_date: str
    agg_eq_pct: float
    neu_eq_pct: float
    market_returns: dict
    country_returns: dict
    bond_returns: dict
    fx_return: float
    signal_history: list = field(default_factory=list)
    seed: Optional[int] = None  # set for reproducibility (testing); None = stochastic


# ─── Phrase Pools (variance source) ──────────────────────────────────────
_TONE_POOLS = {
    'strong_bull': [
        '전반적으로 강세 흐름을 보였습니다',
        '뚜렷한 상승 모멘텀을 이어갔습니다',
        '견조한 강세 기조를 나타냈습니다',
        '전반에 걸쳐 긍정적 흐름이 두드러졌습니다',
    ],
    'mild_bull': [
        '상승세를 나타냈습니다',
        '완만한 상승 흐름을 이어갔습니다',
        '소폭 강세 흐름을 보였습니다',
        '점진적 상승세를 시현하였습니다',
    ],
    'neutral': [
        '제한적 등락 속 횡보 흐름을 보였습니다',
        '뚜렷한 방향성 없이 박스권 흐름을 이어갔습니다',
        '혼조세를 보이며 큰 변동성 없이 마감하였습니다',
        '방향성을 모색하는 횡보 흐름이 지속되었습니다',
    ],
    'mild_bear': [
        '약세 압력이 확대되었습니다',
        '하락 흐름이 우세하였습니다',
        '소폭 약세 흐름을 시현하였습니다',
        '조정 압력이 이어졌습니다',
    ],
    'strong_bear': [
        '큰 폭의 조정 흐름을 보였습니다',
        '뚜렷한 약세 흐름이 두드러졌습니다',
        '본격적인 하락 흐름이 전개되었습니다',
        '두드러진 조정 압력이 가속화되었습니다',
    ],
}
_CONNECTOR = ['또한', '더불어', '아울러', '이와 더불어', '이외에도']
_POS_VERB = ['긍정적으로 기여하였습니다', '포트폴리오 성과를 견인하였습니다',
             '포트폴리오 성과에 일조하였습니다', '플러스 요인으로 작용하였습니다']
_NEG_VERB = ['부정적으로 작용하였습니다', '포트폴리오 성과에 차감 요인으로 작용하였습니다',
             '마이너스 요인으로 작용하였습니다', '성과 부담 요인이 되었습니다']
_GOOD_PERF = ['양호한 성과를 시현', '견조한 성과를 기록', '우수한 흐름을 시현', '긍정적 성과를 시현']
_BAD_PERF = ['부진한 성과를 기록', '약세 흐름을 시현', '저조한 흐름을 보였으나', '미흡한 성과를 시현']
_MAGNITUDE_BIG = ['큰 폭으로', '뚜렷이', '두드러지게', '상당 폭']
_MAGNITUDE_SMALL = ['소폭', '완만하게', '제한적으로', '미미하게']

# 시그널별 emoji
_SIG_LBL = {'Bull': '🟢 Bull', 'Base': '🟡 Base', 'Bear': '🔴 Bear'}


def _rc(seq):
    """Random choice — wraps for clarity."""
    return random.choice(seq)


# ─── Specialist Agents ───────────────────────────────────────────────────

def macro_agent(ctx: CommentaryContext) -> dict:
    mr = ctx.market_returns
    spy, qqq = mr.get('SPY', 0), mr.get('QQQ', 0)
    acwi, efa, vwo = mr.get('ACWI', 0), mr.get('EFA', 0), mr.get('VWO', 0)
    avg_eq = (spy + qqq + acwi + efa + vwo) / 5
    tech_gap = qqq - spy

    if avg_eq > 2:    tone_key, tone_label = 'strong_bull', '강세'
    elif avg_eq > 0.5: tone_key, tone_label = 'mild_bull',   '상승'
    elif avg_eq > -0.5: tone_key, tone_label = 'neutral',    '횡보'
    elif avg_eq > -2:  tone_key, tone_label = 'mild_bear',   '약세'
    else:              tone_key, tone_label = 'strong_bear', '조정'
    tone = '글로벌 증시는 ' + _rc(_TONE_POOLS[tone_key])

    # US narrative — 4 variants per case
    if qqq > 1 and tech_gap > 0.5:
        us = _rc([
            f'미국 증시는 기술주 강세 흐름(QQQ {qqq:+.2f}%, S&P 500 {spy:+.2f}%)이 두드러졌고',
            f'미국 증시는 빅테크 모멘텀(QQQ {qqq:+.2f}%, S&P 500 {spy:+.2f}%)이 부각되며 상승세를 주도하였고',
            f'미국 시장은 나스닥 중심(QQQ {qqq:+.2f}%) 상승 흐름이 S&P 500({spy:+.2f}%) 대비 강하게 나타났고',
        ])
    elif spy > 0 and qqq > 0:
        us = _rc([
            f'미국 증시는 대형주 전반의 견조한 상승(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%) 흐름이 이어졌고',
            f'미국 증시는 광역 지수 동반 상승(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%)으로 긍정 분위기를 이어갔고',
            f'미국 증시는 전반적 상승(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%) 기조를 유지하였고',
        ])
    elif spy < -2:
        us = _rc([
            f'미국 증시는 큰 폭의 조정(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%) 흐름을 보였고',
            f'미국 증시는 광범위한 하락 압력(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%)이 확대되었고',
            f'미국 시장은 뚜렷한 조정(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%) 흐름이 지속되었고',
        ])
    elif spy < 0 and qqq < 0:
        us = _rc([
            f'미국 증시는 약세 흐름(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%)을 나타냈고',
            f'미국 증시는 동반 하락(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%) 압력이 이어졌고',
            f'미국 시장은 부진한 흐름(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%)을 보였고',
        ])
    else:
        us = _rc([
            f'미국 증시는 혼조세(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%)를 보였고',
            f'미국 시장은 방향성 없는 등락(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%) 흐름을 이어갔고',
            f'미국 증시는 박스권 흐름(S&P 500 {spy:+.2f}%, QQQ {qqq:+.2f}%)을 유지하였고',
        ])

    # Intl
    if efa > 0.5 and vwo > 0.5:
        intl = _rc([
            f'선진국(EFA {efa:+.2f}%)과 신흥국(VWO {vwo:+.2f}%) 증시도 동반 상승하였습니다',
            f'선진국(EFA {efa:+.2f}%)·신흥국(VWO {vwo:+.2f}%) 증시 모두 강세를 시현하였습니다',
        ])
    elif efa > 0 and vwo > 0:
        intl = _rc([
            f'선진국(EFA {efa:+.2f}%)과 신흥국(VWO {vwo:+.2f}%) 증시 모두 상승하였습니다',
            f'선진국(EFA {efa:+.2f}%), 신흥국(VWO {vwo:+.2f}%) 증시는 모두 양호한 흐름을 보였습니다',
        ])
    elif efa > 0:
        intl = _rc([
            f'선진국(EFA {efa:+.2f}%) 증시는 상승하였으나 신흥국(VWO {vwo:+.2f}%)은 상대적으로 부진하였습니다',
            f'선진국(EFA {efa:+.2f}%)은 강세를 시현한 반면 신흥국(VWO {vwo:+.2f}%)은 약세를 보였습니다',
        ])
    elif vwo > 0:
        intl = _rc([
            f'신흥국(VWO {vwo:+.2f}%)이 상대적으로 견조한 반면 선진국(EFA {efa:+.2f}%)은 약세를 시현하였습니다',
            f'신흥국(VWO {vwo:+.2f}%)은 양호한 흐름, 선진국(EFA {efa:+.2f}%)은 약세를 보였습니다',
        ])
    else:
        intl = _rc([
            f'선진국(EFA {efa:+.2f}%)과 신흥국(VWO {vwo:+.2f}%) 증시 모두 약세를 시현하였습니다',
            f'선진국(EFA {efa:+.2f}%), 신흥국(VWO {vwo:+.2f}%) 모두 부진한 흐름을 이어갔습니다',
        ])

    # Core US ETFs
    us_avg = (spy + qqq) / 2
    if us_avg > 1:
        core = (f'주식 포트폴리오 내 주요 편입 대상인 **S&P 500 ETF**(SPY {spy:+.2f}%)와 '
                f'**나스닥 100 ETF**(QQQ {qqq:+.2f}%)는 {_rc(_GOOD_PERF)}하며 {_rc(_POS_VERB)}')
    elif us_avg > 0:
        core = (f'주식 포트폴리오 내 주요 편입 대상인 **S&P 500 ETF**(SPY {spy:+.2f}%)와 '
                f'**나스닥 100 ETF**(QQQ {qqq:+.2f}%)는 {_rc(_MAGNITUDE_SMALL)} 상승하며 {_rc(_POS_VERB)}')
    elif us_avg > -1:
        core = (f'주식 포트폴리오 내 주요 편입 대상인 **S&P 500 ETF**(SPY {spy:+.2f}%)와 '
                f'**나스닥 100 ETF**(QQQ {qqq:+.2f}%)는 혼조 흐름을 보였습니다')
    else:
        core = (f'주식 포트폴리오 내 주요 편입 대상인 **S&P 500 ETF**(SPY {spy:+.2f}%)와 '
                f'**나스닥 100 ETF**(QQQ {qqq:+.2f}%)는 {_rc(_BAD_PERF)}하며 {_rc(_NEG_VERB)}')

    # Global ETFs
    glob_avg = (acwi + efa + vwo) / 3
    glob_sent = (
        f'{_rc(_CONNECTOR)} 전반적인 글로벌 증시에 투자하는 **MSCI ACWI ETF**(ACWI {acwi:+.2f}%), '
        f'미국 외 선진국 증시 **MSCI EAFE ETF**(EFA {efa:+.2f}%), '
        f'**이머징 증시 ETF**(VWO {vwo:+.2f}%)도 '
        f'{"상승" if glob_avg > 0 else "하락"}하며 '
        f'포트폴리오 성과에 {"긍정적" if glob_avg > 0 else "부정적"}으로 기여하였습니다'
    )

    return {
        'role': 'macro',
        'narrative': f'{tone}. {us}, {intl}. {core}. {glob_sent}',
        'facts': {'tone': tone_label, 'avg_eq': round(avg_eq, 2),
                  'tech_gap': round(tech_gap, 2)},
    }


def country_agent(ctx: CommentaryContext) -> dict:
    if not ctx.country_returns:
        return {'role': 'country', 'narrative': '', 'facts': {}}
    items = sorted(ctx.country_returns.items(), key=lambda x: x[1], reverse=True)
    positives = [(c, r) for c, r in items if r > 0]
    negatives = [(c, r) for c, r in items if r <= 0]
    if positives and not negatives:
        top_names = ', '.join(f'**{c}**({r:+.2f}%)' for c, r in positives[:3])
        narrative = _rc([
            f'국가 배분 전략 관점에서 전술적으로 편입한 {top_names} 증시 ETF는 포트폴리오 성과를 더해준 요인이었습니다.',
            f'국가 배분 측면에서 {top_names} 증시 ETF가 양호한 성과로 포트폴리오에 기여하였습니다.',
            f'전술적으로 편입한 {top_names} 증시 ETF는 성과 기여 요인으로 작용하였습니다.',
        ])
        outcome = 'positive'
    elif positives and negatives:
        top_names = ', '.join(f'**{c}**({r:+.2f}%)' for c, r in positives[:2])
        bot_names = ', '.join(f'**{c}**({r:+.2f}%)' for c, r in negatives[-2:])
        narrative = _rc([
            f'국가 배분 측면에서 {top_names} 증시 ETF는 성과 기여 요인으로 작용하였으며, {bot_names} 증시는 상대적으로 부진하여 일부 차감 요인으로 나타났습니다.',
            f'국가 배분 측면에서는 {top_names} ETF가 플러스 요인으로, {bot_names} ETF는 마이너스 요인으로 작용한 혼조 흐름을 보였습니다.',
            f'전술적으로 편입한 {top_names} ETF는 성과에 기여한 반면, {bot_names} ETF는 약세를 시현하였습니다.',
        ])
        outcome = 'mixed'
    else:
        bot_names = ', '.join(f'**{c}**({r:+.2f}%)' for c, r in negatives[:3])
        narrative = _rc([
            f'국가 배분 측면에서는 {bot_names} 증시 ETF가 전반적으로 약세를 보이며 차감 요인으로 작용하였습니다.',
            f'국가 배분 측면에서 {bot_names} ETF의 부진이 성과 부담 요인이 되었습니다.',
        ])
        outcome = 'negative'
    return {'role': 'country', 'narrative': narrative,
            'facts': {'outcome': outcome}}


def bond_agent(ctx: CommentaryContext) -> dict:
    br = ctx.bond_returns
    ief_r = br.get('IEF', 0)

    if ief_r > 1:
        bond_macro = _rc([
            f'중장기 국채 금리는 큰 폭으로 하락(IEF {ief_r:+.2f}%)하며 채권 가격은 강세를 시현하였고',
            f'중장기 금리가 뚜렷이 하락(IEF {ief_r:+.2f}%)하며 채권 가격은 우상향 흐름을 보였고',
        ])
    elif ief_r > 0.3:
        bond_macro = _rc([
            f'중장기 국채 금리는 소폭 하락(IEF {ief_r:+.2f}%)하며 채권 가격은 강보합 흐름을 보였고',
            f'중장기 금리는 완만한 하락세(IEF {ief_r:+.2f}%)를 보이며 채권 가격은 소폭 상승하였고',
        ])
    elif ief_r > -0.3:
        bond_macro = _rc([
            f'중장기 국채 금리는 좁은 박스권에서 등락(IEF {ief_r:+.2f}%)을 보였고',
            f'중장기 금리는 뚜렷한 방향성 없이 박스권 흐름(IEF {ief_r:+.2f}%)을 이어갔고',
        ])
    elif ief_r > -1:
        bond_macro = _rc([
            f'중장기 국채 금리는 상승 압력(IEF {ief_r:+.2f}%)이 이어지며 채권 가격은 약보합세를 보였고',
            f'중장기 금리가 완만한 상승(IEF {ief_r:+.2f}%) 압력을 받으며 채권 가격은 부진하였고',
        ])
    else:
        bond_macro = _rc([
            f'중장기 국채 금리는 큰 폭의 상승(IEF {ief_r:+.2f}%)을 보이며 채권 가격은 약세를 기록하였고',
            f'중장기 금리가 뚜렷이 상승(IEF {ief_r:+.2f}%)하며 채권 가격은 본격적 조정을 받았고',
        ])
    bond_vol = _rc(['높은 금리 변동성 장세로 전개되었습니다', '제한적 변동성 흐름을 이어갔습니다',
                     '변동성이 다소 확대된 흐름을 보였습니다']) if abs(ief_r) > 1.5 else \
                _rc(['제한적 변동성 흐름을 이어갔습니다', '안정적인 변동성 흐름이 유지되었습니다'])

    # KR 종합채권
    kr_long_sent = ''
    if 'TIGER 종합채권' in br:
        if ief_r > 0.5:
            kr_long_sent = _rc([
                '이러한 금리 흐름에 따라 중장기 듀레이션을 보유한 **TIGER 종합채권(AA-이상)액티브 ETF**는 양호한 성과를 시현하였습니다. ',
                '중장기 듀레이션 노출이 큰 **TIGER 종합채권(AA-이상)액티브 ETF**는 듀레이션 효과로 견조한 성과를 기록하였습니다. ',
            ])
        elif ief_r < -0.5:
            kr_long_sent = _rc([
                '이러한 금리 상승 압력으로 중장기 듀레이션을 보유한 **TIGER 종합채권(AA-이상)액티브 ETF**는 다소 부진한 성과를 기록하였습니다. ',
                '금리 상승 영향으로 **TIGER 종합채권(AA-이상)액티브 ETF**는 약세 흐름을 보였습니다. ',
            ])
        else:
            kr_long_sent = _rc([
                '중장기 듀레이션을 보유한 **TIGER 종합채권(AA-이상)액티브 ETF**는 박스권 흐름을 이어갔습니다. ',
                '**TIGER 종합채권(AA-이상)액티브 ETF**는 큰 변동 없이 박스권 흐름을 유지하였습니다. ',
            ])

    # US credit
    credit_rets, credit_labels = [], []
    if 'LQD' in br:
        credit_rets.append(br['LQD'])
        credit_labels.append(f'**미국 투자등급회사채 ETF**(LQD {br["LQD"]:+.2f}%)')
    hy_labels = []
    if 'HYG' in br:
        credit_rets.append(br['HYG'])
        hy_labels.append(f'HYG {br["HYG"]:+.2f}%')
    if 'PHYL' in br:
        credit_rets.append(br['PHYL'])
        hy_labels.append(f'PHYL {br["PHYL"]:+.2f}%')
    if hy_labels:
        credit_labels.append(f'**미국 투기등급회사채 ETF**({", ".join(hy_labels)})')
    if 'EMBD' in br:
        credit_rets.append(br['EMBD'])
        credit_labels.append(f'**이머징 채권 ETF**(EMBD {br["EMBD"]:+.2f}%)')

    credit_avg = sum(credit_rets) / len(credit_rets) if credit_rets else 0
    if credit_labels:
        if credit_avg > 0.3:
            outcome = _rc([
                f'{_rc(_GOOD_PERF)}하며 채권 포트폴리오 성과에 긍정적으로 기여하였습니다',
                '양호한 흐름으로 채권 포트폴리오 성과를 견인하였습니다',
            ])
        elif credit_avg < -0.3:
            outcome = _rc([
                f'{_rc(_BAD_PERF)}하며 채권 포트폴리오 성과에 부정적으로 작용하였습니다',
                '약세 흐름으로 채권 포트폴리오 성과에 차감 요인이 되었습니다',
            ])
        else:
            outcome = _rc([
                '중립적인 흐름을 보이며 채권 포트폴리오 성과 안정성에 기여하였습니다',
                '제한적 등락을 보이며 채권 포트폴리오 안정성을 유지하였습니다',
            ])
        credit_sent = f'{" 및 ".join(credit_labels)}는 {outcome}. '
    else:
        credit_sent = ''

    kr_cd_sent = ''
    if 'TIGER CD금리' in br:
        kr_cd_sent = _rc([
            f'{_rc(_CONNECTOR)} **TIGER CD금리투자KIS(합성) ETF**는 안정적인 단기금리 수익을 통해 채권 포트폴리오의 변동성을 완화하는 요인으로 작용하였습니다. ',
            f'{_rc(_CONNECTOR)} **TIGER CD금리투자KIS(합성) ETF**의 일정 비중 유지가 단기금리 수익을 통해 채권 포트폴리오 안정성에 기여하였습니다. ',
        ])

    return {'role': 'bond',
            'narrative': f'채권 시장의 경우 {bond_macro} {bond_vol} {kr_long_sent}{credit_sent}{kr_cd_sent}'.strip(),
            'facts': {'ief_r': ief_r, 'credit_avg': round(credit_avg, 2)}}


def fx_agent(ctx: CommentaryContext) -> dict:
    fx = ctx.fx_return
    if fx > 1:
        narrative = _rc([
            f'한편 원달러 환율은 월초 대비 {_rc(_MAGNITUDE_BIG)} 상승({fx:+.2f}%)하며 USD 표시 자산의 환차익이 발생해 전체 포트폴리오 성과에 긍정적으로 작용하였습니다.',
            f'원달러 환율의 {_rc(_MAGNITUDE_BIG)} 상승({fx:+.2f}%)으로 USD 자산 환차익이 발생하여 포트폴리오 성과에 플러스 요인으로 작용하였습니다.',
        ])
    elif fx > 0.3:
        narrative = _rc([
            f'한편 원달러 환율은 월초 대비 상승({fx:+.2f}%)하며 USD 표시 자산의 환차익이 발생해 전체 포트폴리오 성과에 {_rc(_MAGNITUDE_SMALL)} 긍정적으로 작용하였습니다.',
            f'원달러 환율의 상승({fx:+.2f}%)으로 USD 자산 환차익이 소폭 발생하였습니다.',
        ])
    elif fx < -1:
        narrative = _rc([
            f'한편 원달러 환율은 월초 대비 {_rc(_MAGNITUDE_BIG)} 하락({fx:+.2f}%)하며 USD 표시 자산의 환손실이 발생해 전체 포트폴리오 성과에 부정적으로 작용하였습니다.',
            f'원달러 환율의 {_rc(_MAGNITUDE_BIG)} 하락({fx:+.2f}%)으로 USD 자산 환손실이 발생하여 포트폴리오 성과에 차감 요인이 되었습니다.',
        ])
    elif fx < -0.3:
        narrative = _rc([
            f'한편 원달러 환율은 월초 대비 하락({fx:+.2f}%)하며 USD 표시 자산의 환손실이 발생해 전체 포트폴리오 성과에 {_rc(_MAGNITUDE_SMALL)} 부정적으로 작용하였습니다.',
            f'원달러 환율의 하락({fx:+.2f}%)으로 USD 자산 환손실이 일부 발생하였습니다.',
        ])
    else:
        narrative = _rc([
            f'한편 원달러 환율은 월초 대비 큰 변동 없이({fx:+.2f}%) 안정적인 흐름을 유지하였습니다.',
            f'원달러 환율은 ({fx:+.2f}%) 변동 폭이 제한적이었습니다.',
        ])
    return {'role': 'fx', 'narrative': narrative,
            'facts': {'fx': fx}}


def signal_agent(ctx: CommentaryContext) -> dict:
    hist = ctx.signal_history or []
    if not hist:
        narrative = _rc([
            f'AIMVP 시그널은 **{ctx.signal}** 상태(발효일 {ctx.signal_date})를 유지하였습니다.',
            f'현재 AIMVP 매크로 시그널은 **{ctx.signal}**(발효 {ctx.signal_date})입니다.',
        ])
        return {'role': 'signal', 'narrative': narrative,
                'facts': {'current': ctx.signal}}

    streak = 0
    for s in reversed(hist):
        if s == ctx.signal:
            streak += 1
        else:
            break
    bull_n, base_n, bear_n = hist.count('Bull'), hist.count('Base'), hist.count('Bear')
    dominant = max([(bull_n, 'Bull'), (base_n, 'Base'), (bear_n, 'Bear')])[1]

    if streak >= 3:
        ctx_phrase = _rc([
            f'{streak}개월 연속 **{ctx.signal}** 시그널이 발효되며',
            f'**{ctx.signal}** 시그널이 {streak}개월째 지속되며',
            f'{streak}개월간 **{ctx.signal}** 시그널 기조가 이어지며',
        ])
    elif streak == 2:
        ctx_phrase = _rc([
            f'2개월 연속 **{ctx.signal}** 시그널이 발효되며',
            f'**{ctx.signal}** 시그널이 2개월째 유지되며',
        ])
    else:
        ctx_phrase = _rc([
            f'**{ctx.signal}** 시그널이 발효되며(발효일 {ctx.signal_date})',
            f'AIMVP 매크로 시그널이 **{ctx.signal}** 상태로 전환되며',
        ])

    if dominant == ctx.signal:
        macro_phrase = _rc([
            f'(최근 12개월 중 {dominant} {hist.count(dominant)}회로 우세)',
            f'(지난 12개월간 {dominant} 시그널이 {hist.count(dominant)}회로 가장 빈번)',
        ])
    else:
        macro_phrase = _rc([
            f'(최근 12개월 분포: Bull {bull_n}/Base {base_n}/Bear {bear_n}회)',
            f'(지난 12개월 시그널 발효 — Bull {bull_n}회, Base {base_n}회, Bear {bear_n}회)',
        ])

    narrative = f'AIMVP 시그널은 {ctx_phrase} {macro_phrase}.'
    return {'role': 'signal', 'narrative': narrative,
            'facts': {'current': ctx.signal, 'streak': streak,
                      'bull': bull_n, 'base': base_n, 'bear': bear_n}}


# ─── Synthesizer ─────────────────────────────────────────────────────────

def synthesizer_agent(drafts: dict, ctx: CommentaryContext) -> str:
    """5 drafts를 단일 narrative로 통합 — 헤더 + 단락 1(주식) + 단락 2(채권/FX)."""
    macro_n   = drafts.get('macro',   {}).get('narrative', '')
    country_n = drafts.get('country', {}).get('narrative', '')
    bond_n    = drafts.get('bond',    {}).get('narrative', '')
    fx_n      = drafts.get('fx',      {}).get('narrative', '')
    signal_n  = drafts.get('signal',  {}).get('narrative', '')

    # 헤더
    if ctx.kind == 'weekly':
        header = (f'**📅 주간 시장 요약 ({ctx.period_range})**\n\n'
                  f'**ETF AI MVP (적극), ETF AI MVP (중립)**\n\n')
    else:
        header = (f'**📰 MVP시리즈 펀드 매니저 Comment ({ctx.period_label})**\n\n'
                  f'**ETF AI MVP (적극), ETF AI MVP (중립)**\n\n')

    # 단락 1 도입 — variants
    intro_variants = [
        f'ETF AI MVP (적극), ETF AI MVP (중립) 펀드는 위험자산의 비중을 각각 **{ctx.agg_eq_pct:.1f}%, {ctx.neu_eq_pct:.1f}%**로 운용하였으며,',
        f'ETF AI MVP (적극)·중립 펀드의 위험자산 비중은 각각 **{ctx.agg_eq_pct:.1f}%, {ctx.neu_eq_pct:.1f}%** 수준이며,',
        f'이번 기간 ETF AI MVP (적극), ETF AI MVP (중립) 펀드는 위험자산 비중을 **{ctx.agg_eq_pct:.1f}% / {ctx.neu_eq_pct:.1f}%**로 운용하였고,',
    ]
    para1 = f'{_rc(intro_variants)} {signal_n} {macro_n}. {country_n}'.strip()
    para2 = f'{bond_n} {fx_n}'.strip()
    return header + para1 + '\n\n' + para2


# ─── Editor ──────────────────────────────────────────────────────────────

def editor_agent(synthesized: str, ctx: CommentaryContext) -> str:
    """공백 정리 + 미세 lexical swap."""
    text = synthesized.replace('\n\n\n', '\n\n').replace('  ', ' ').strip()
    # 마지막 마무리 phrase optional 추가 (monthly only, 가끔)
    if ctx.kind == 'monthly' and random.random() < 0.4:
        tail = _rc([
            ' 향후에도 시그널 기반 자산 배분 전략을 일관되게 유지할 계획입니다.',
            ' 향후에도 매크로 시그널에 따른 자산 배분 원칙을 지속할 예정입니다.',
            ' 펀드는 향후에도 시그널 기반 자산 배분 프로세스를 유지합니다.',
        ])
        text = text + tail
    return text


# ─── Swarm Orchestrator ──────────────────────────────────────────────────

class CommentarySwarm:
    SPECIALISTS = [
        ('macro',   macro_agent),
        ('country', country_agent),
        ('bond',    bond_agent),
        ('fx',      fx_agent),
        ('signal',  signal_agent),
    ]

    def __init__(self, ctx: CommentaryContext):
        self.ctx = ctx
        if ctx.seed is not None:
            random.seed(ctx.seed)
        # else: 시스템 entropy → 매 generation 다른 결과

    def generate(self) -> dict:
        drafts = {}
        with ThreadPoolExecutor(max_workers=5) as ex:
            futures = {ex.submit(fn, self.ctx): role
                        for role, fn in self.SPECIALISTS}
            for fut in futures:
                role = futures[fut]
                try:
                    drafts[role] = fut.result()
                except Exception as e:
                    drafts[role] = {'role': role,
                                    'narrative': f'[{role} 실패: {e}]', 'facts': {}}
        synthesized = synthesizer_agent(drafts, self.ctx)
        final = editor_agent(synthesized, self.ctx)
        return {'final': final, 'drafts': drafts, 'synthesized': synthesized}


def generate_commentary_sync(ctx: CommentaryContext) -> dict:
    swarm = CommentarySwarm(ctx)
    return swarm.generate()


# ─── Self-test (반복 실행 시 변형 확인) ────────────────────────────────
if __name__ == '__main__':
    ctx = CommentaryContext(
        kind='monthly',
        period_label='2026년 6월',
        period_range='6/1 ~ 6/8',
        signal='Bull',
        signal_date='2026-06-01',
        agg_eq_pct=89.55, neu_eq_pct=59.72,
        market_returns={'SPY': 1.2, 'QQQ': 2.3, 'ACWI': 1.0, 'EFA': 0.5, 'VWO': 0.3},
        country_returns={'일본': 0.7, '캐나다': 0.3, '영국': 1.1, '한국': -0.2},
        bond_returns={'HYG': 0.2, 'LQD': 0.1, 'IEF': -0.3,
                      'TIGER 종합채권': 0.0, 'TIGER CD금리': 0.0},
        fx_return=0.6,
        signal_history=['Base']*5 + ['Bull']*7,
    )
    print('=== Run 1 ===')
    print(generate_commentary_sync(ctx)['final'])
    print('\n=== Run 2 (재실행 → 다른 narrative) ===')
    print(generate_commentary_sync(ctx)['final'])
