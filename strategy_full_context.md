# 50 EMA Rejection Strategy — Full Context for Continuation

## Overview
This is a day trading indicator for TradingView (Pine Script v6) that identifies high-probability entries based on 50 EMA rejection setups with volume confirmation, built through 6 iterative versions with real trade data analysis.

---

## Core Strategy Framework

### Setup Logic
1. **Bias**: Price above 50 EMA = bullish (long entries), below = bearish (short entries)
2. **Rejection**: Price touches/wicks through 50 EMA but closes back on the trending side. Must be a strong candle (hammer, shooting star, engulfing, pin bar)
3. **Volume**: Rejection candle must show above-average volume (1.2× the 10-bar average)
4. **Entry**: Break of the confirmation candle's high (long) or low (short) after the rejection
5. **Stops**: Below/above the rejection wick
6. **Targets**: T1 at 1:1 R/R, T2 at 2:1 R/R

### Timeframes
- 5-minute: Primary setup identification
- 1-minute: Entry refinement and the main timeframe used for backtesting

---

## Version History & What Each Solved

### v1 — Base indicator
- 50 EMA + rejection candle detection + volume filter
- Entry/SL/T1/T2 lines on chart
- Grade A/B system, session filter, EMA slope check
- **Problem**: Generated false signals on flat EMAs and counter-trend bounces

### v2 — Anti-false-signal filters
Added 5 filters:
1. **EMA Curvature** (2nd derivative) — blocks shorts when EMA curling up
2. **Counter-Momentum** — blocks when prior candles oppose trade direction
3. **Min Body Size** — rejection candle must have real body (≥0.2× ATR)
4. **RSI Extreme** — blocks shorts when oversold, longs when overbought
5. **Structure Confirmation** — requires lower-highs for shorts, higher-lows for longs
- **Problem**: Killed valid setups where price pulled back into EMA (pullback candles are supposed to be counter-trend)

### v3 — Rubber Band filter
Added:
1. **Rubber Band Separation** — measures max distance price traveled from EMA before pullback (normalized to ATR). Min 1.5× ATR required.
2. **Approach Angle** — pullback must have directional intent, not sideways drift
3. **First Touch Detection** — scans 50 bars back for prior EMA touches
- **Problem**: Still not triggering on the best setups because momentum filter was backwards

### v4 — Fixed momentum logic (critical fix)
**Root cause found**: The momentum filter checked candles RIGHT BEFORE the rejection — but those ARE the pullback and SHOULD be counter-trend. It was blocking the exact setups we wanted.

Fixes:
1. **Trend-Leg Momentum** — now checks bars 10-20 ago (the actual trend leg), not the recent pullback
2. **RSI reworked** — only blocks at extreme levels (RSI < 30 AND still falling), not moderate zones
3. **Rejection detection expanded** — alternative detection if candle closes in bottom 35% of range near EMA
4. **Proximity tolerance** — near-touches within 0.1% of EMA count
5. **Near-miss debug mode** — shows where rejection formed but confirmation didn't fire yet

### v5 — Trade tracking engine + win rate stats
Added complete trade simulation:
- Tracks each trade bar-by-bar: checks if T1, T2, or SL hit first
- Stores last 100 trades with full metadata (direction, grade, entry, outcome, R-multiple, separation ATR, RSI, volume ratio, first touch flag)
- **Stats table**: Overall win rate, T1/T2 hit rates, Grade A vs B, Long vs Short, First Touch, Avg R-multiple
- **Trade log**: Last 10 (later 20) trades with per-trade details
- Fixed Pine v6 compilation errors: `wrColor` function moved to global scope, all int divisions wrapped in `float()` casts

### v6 — Data-driven optimization (current version)
Based on analysis of 12 SPY trades (67% WR, 0.67R avg):

**Key findings from data:**
- Grade A was too strict (only 1/12 trades qualified, and it lost)
- Sweet spot for separation: 3.0–5.0× ATR (all T2 hits were in this zone)
- RSI didn't differentiate winners from losers
- Winners resolved in 3-5 bars, losers dragged out 7-12 bars
- Overextended trades (>6.5× ATR separation) were losers

**Changes:**
1. **Grade A now driven by separation sweet spot** (3.0–5.0× ATR) + volume OR first touch
2. **Overextension cap** — blocks signals above 6.5× ATR separation
3. **RSI removed as entry filter** (kept as display metric only)
4. **Time-based early exit** — if trade hasn't moved 0.3R toward target by bar 8, close at market (outcome code 4)
5. **First-touch lookback reduced** from 50 to 20 bars
6. **Signal queue system** — signals that fire while a trade is active get queued and opened immediately when current trade resolves (no more missed trades)
7. **Max resolve reduced** from 50 to 25 bars
8. **Trade log expanded** to 20 rows
9. **Sweet spot stats** tracked separately in the stats table

---

## Performance Data

### SPY Results (v6, 11 trades, 200 shares)
- **Overall**: 73.1% win rate, 0.89R avg, +$549.80 total P&L
- **Grade A**: 67.1% (2W/0L from 3 trades)
- **Grade B**: 75.1% (6W/2L from 8 trades)
- **Sweet Spot**: 71.1% (5W/1L from 7 trades)
- **Longs**: 67.1% | **Shorts**: 67.1%
- **Avg bars**: Winners 6.5, Losers 5.5

### NVDA Results (v6, 17 trades, 300 shares)
- **Overall**: 88.1% win rate, 1.35R avg, 23R total
- **Grade A**: 100% (5W/0L from 5 trades) ← PERFECT
- **Grade B**: 83.1% (10W/2L from 12 trades)
- **Sweet Spot**: 93.1% (13W/1L from 14 trades)
- **1st Touch**: 100% (5W/0L from 5 trades) ← PERFECT
- **Longs**: 83.1% (5W/1L) | **Shorts**: 91.1% (10W/1L)
- **Avg bars**: Winners 10.2, Losers 5.0
- **Key insight**: The only low-separation trade (#8, 1.2× ATR) was one of only 2 losses. Consider raising min separation to 1.5×.

### NVDA Trade Log (all 17 trades):
```
#   Dir  Gr  Entry    Result    R    Bars  Sep    RSI   Sweet  1stTouch
1   S    B   182.42   T2        +2R  10b   2.5x   43.9  no     no
2   L    A   178.85   T1        +1R  25b   5.0x   53.6  no     yes
3   L    B   182.99   T2        +2R  8b    3.2x   58.7  yes    no
4   L    B   180.33   T2        +2R  1b    3.5x   56.3  yes    no
5   S    B   180.95   T1        +1R  10b   5.3x   47.4  no     yes
6   S    B   180.79   SL        -1R  6b    3.3x   44.0  yes    no
7   S    B   180.52   T2        +2R  2b    2.8x   52.6  no     no
8   L    B   185.61   SL        -1R  4b    1.2x   52.6  no     no
9   S    B   185.68   T2        +2R  4b    3.5x   43.9  yes    no
10  S    A   183.12   T1        +1R  2b    3.5x   45.6  yes    yes
11  L    A   184.30   T1        +1R  20b   4.1x   56.0  yes    no
12  L    A   184.80   T1        +1R  19b   4.4x   55.6  yes    yes
13  S    B   183.90   T2        +2R  18b   3.1x   42.8  yes    no
14  S    B   182.62   T2        +2R  12b   3.7x   44.9  yes    no
15  S    A   180.71   T2        +2R  16b   3.4x   37.7  yes    no
16  S    B   175.35   T2        +2R  3b    3.3x   43.1  yes    no
17  S    A   173.55   T2        +2R  3b    4.6x   50.1  yes    yes
```

---

## Current Filters (v6 Active)

| Filter | Purpose | Default |
|--------|---------|---------|
| EMA Slope | Must be trending (not flat) | 0.01% min over 5 bars |
| EMA Curvature | Block when EMA curling against trade | 3-bar lookback |
| Min Body Size | No doji rejection candles | 0.15× ATR |
| Structure | Lower-highs for shorts, higher-lows for longs | 10-bar lookback |
| Trend-Leg Momentum | Trend leg (bars 10-20 ago) must match direction | 50% threshold |
| Rubber Band | Min separation before pullback | 1.0× ATR |
| Sweet Spot | Grade A boost zone | 3.0–5.0× ATR |
| Overextension Cap | Block if too far from EMA | 6.5× ATR |
| Approach Angle | Pullback must have directional intent | 0.3× ATR min |
| First Touch | Boost if first EMA test in 20 bars | 20-bar lookback |
| Early Exit | Cut if <0.3R progress by bar 8 | 8 bars, 0.3R min |
| Session | Only trade 9:35-15:55 ET | Configurable |

---

## Known Issues & Next Steps

1. **Min separation might be too low** — NVDA trade #8 at 1.2× ATR was a loss. Consider raising from 1.0× to 1.5×.
2. **Long trades on NVDA resolve slower** (19-25 bars) vs shorts (2-4 bars for T2). May need different early exit settings per direction.
3. **P&L calculation needs actual SL distances** from the rejection candle range, not estimated percentages. The indicator stores entry and SL — the risk is `|entry - SL|`.
4. **Sample size still small** — need 50+ trades on each ticker to confirm edge is real.
5. **Multi-ticker testing needed** — works on SPY and NVDA, should test AAPL, QQQ, TSLA, AMD.
6. **Could add**: VWAP confluence filter, time-of-day analysis (which hours produce best setups), consecutive loss circuit breaker.

---

## File Locations

- Current Pine Script: `50_EMA_Rejection_v6.pine`
- All versions preserved: v1 through v6 in conversation history
- The indicator runs on TradingView, Pine Script v6, overlay mode

---

## How to Continue Development

Paste this context into Claude Code and you can:
1. Ask for v7 with specific improvements based on new trade data
2. Request P&L analysis with actual share sizes and risk calculations
3. Add new filters or modify existing ones
4. Port the strategy to a different platform (Python backtest, etc.)
5. Build a companion dashboard or alert system
