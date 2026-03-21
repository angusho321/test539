#!/usr/bin/env python3
import argparse
from dataclasses import dataclass
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd


DRAW_COLS = ["號碼1", "號碼2", "號碼3", "號碼4", "號碼5"]


@dataclass
class EvalResult:
    name: str
    n: int
    days: int
    hits: int
    cost: int
    prize: int
    net: int
    roi: float
    params: Dict[str, float]


def load_hist(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期"]).copy()
    for c in DRAW_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df.dropna(subset=DRAW_COLS).copy()
    for c in DRAW_COLS:
        df[c] = df[c].astype(int)
    return df.sort_values("日期").reset_index(drop=True)


def num_freq_scores(train_df: pd.DataFrame, decay: float) -> np.ndarray:
    scores = np.zeros(40, dtype=float)
    max_date = train_df["日期"].max()
    days_ago = (max_date - train_df["日期"]).dt.days.clip(lower=0).to_numpy()
    row_w = np.power(decay, days_ago)
    draws = train_df[DRAW_COLS].to_numpy(dtype=int)
    for i in range(draws.shape[0]):
        w = row_w[i]
        for num in draws[i]:
            scores[num] += w
    scores[0] = -1e9
    return scores


def weekday_scores(train_df: pd.DataFrame, weekday: int) -> np.ndarray:
    sub = train_df[train_df["日期"].dt.weekday == weekday]
    scores = np.zeros(40, dtype=float)
    if len(sub) == 0:
        return scores
    draws = sub[DRAW_COLS].to_numpy(dtype=int)
    for row in draws:
        for num in row:
            scores[num] += 1.0
    scores[0] = -1e9
    return scores / len(sub)


def recent_momentum_scores(train_df: pd.DataFrame, k: int) -> np.ndarray:
    scores = np.zeros(40, dtype=float)
    if k <= 0:
        return scores
    sub = train_df.tail(k)
    draws = sub[DRAW_COLS].to_numpy(dtype=int)
    for row in draws:
        for num in row:
            scores[num] += 1.0
    scores[0] = -1e9
    return scores / max(1, len(sub))


def overdue_scores(train_df: pd.DataFrame, cap: int = 60) -> np.ndarray:
    # 出現間隔越久，分數越高（冷號回歸假設）
    scores = np.zeros(40, dtype=float)
    draws = train_df[DRAW_COLS].to_numpy(dtype=int)
    last_seen = np.full(40, -1, dtype=int)
    for i in range(draws.shape[0]):
        for num in draws[i]:
            last_seen[num] = i
    end = len(draws) - 1
    for num in range(1, 40):
        if last_seen[num] < 0:
            dist = cap
        else:
            dist = min(cap, end - last_seen[num])
        scores[num] = float(dist)
    scores[0] = -1e9
    return scores / max(1.0, scores[1:].max())


def pick_top_n(scores: np.ndarray, n: int) -> List[int]:
    return sorted(int(x) for x in np.argsort(scores)[::-1][:n])


def evaluate_strategy(
    df: pd.DataFrame,
    n: int,
    lookback: int,
    cost_per_num: int,
    prize_per_hit: int,
    decay: float,
    w_weekday: float,
    w_momentum: float,
    momentum_k: int,
    w_overdue: float,
) -> EvalResult:
    hits = 0
    cost = 0
    prize = 0
    valid_days = 0

    for i in range(len(df)):
        target_date = df.iloc[i]["日期"]
        start = target_date - pd.Timedelta(days=lookback)
        train = df[(df["日期"] < target_date) & (df["日期"] >= start)]
        if len(train) < 50:
            continue

        base = num_freq_scores(train, decay=decay)
        wday = weekday_scores(train, weekday=int(target_date.weekday()))
        mom = recent_momentum_scores(train, k=momentum_k)
        ovd = overdue_scores(train)
        combo = base + w_weekday * wday + w_momentum * mom + w_overdue * ovd

        pred = pick_top_n(combo, n=n)
        actual = set(int(df.iloc[i][c]) for c in DRAW_COLS)
        h = len(actual.intersection(pred))
        hits += h
        cost += cost_per_num * n
        prize += prize_per_hit * h
        valid_days += 1

    net = prize - cost
    roi = (net / cost) if cost > 0 else 0.0
    name = "hybrid_ev"
    return EvalResult(
        name=name,
        n=n,
        days=valid_days,
        hits=hits,
        cost=cost,
        prize=prize,
        net=net,
        roi=roi,
        params={
            "lookback": lookback,
            "decay": decay,
            "w_weekday": w_weekday,
            "w_momentum": w_momentum,
            "momentum_k": momentum_k,
            "w_overdue": w_overdue,
        },
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Search profitable Fantasy5 strategy")
    parser.add_argument("--history", default="fantasy5_hist.xlsx")
    parser.add_argument("--cost-per-num", type=int, default=285)
    parser.add_argument("--prize-per-hit", type=int, default=2100)
    args = parser.parse_args()

    df = load_hist(args.history)
    print(f"rows={len(df)} range={df['日期'].min().date()}~{df['日期'].max().date()}")

    results: List[EvalResult] = []
    lookbacks = [120, 150, 180, 240, 300]
    decays = [0.97, 0.98, 0.985, 0.99, 0.995]
    w_weekdays = [0.0, 0.2, 0.5, 0.8]
    w_moms = [0.0, 0.2, 0.5, 1.0]
    mom_ks = [7, 14, 21, 30]
    w_overdues = [0.0, 0.1, 0.2, 0.4]

    total = (
        2
        * len(lookbacks)
        * len(decays)
        * len(w_weekdays)
        * len(w_moms)
        * len(mom_ks)
        * len(w_overdues)
    )
    done = 0
    for n in [7, 9]:
        for lb in lookbacks:
            for d in decays:
                for ww in w_weekdays:
                    for wm in w_moms:
                        for mk in mom_ks:
                            for wo in w_overdues:
                                done += 1
                                if done % 200 == 0:
                                    print(f"progress {done}/{total}")
                                r = evaluate_strategy(
                                    df=df,
                                    n=n,
                                    lookback=lb,
                                    cost_per_num=args.cost_per_num,
                                    prize_per_hit=args.prize_per_hit,
                                    decay=d,
                                    w_weekday=ww,
                                    w_momentum=wm,
                                    momentum_k=mk,
                                    w_overdue=wo,
                                )
                                results.append(r)

    results.sort(key=lambda x: x.net, reverse=True)
    top = results[:10]
    print("\nTop 10 by net profit:")
    for i, r in enumerate(top, 1):
        print(
            f"{i}. n={r.n} net={r.net:,} roi={r.roi*100:.2f}% "
            f"hits={r.hits} days={r.days} params={r.params}"
        )

    positives = [r for r in results if r.net > 0]
    print(f"\npositive_count={len(positives)} / {len(results)}")
    if positives:
        best = positives[0]
        print(
            f"best_positive: n={best.n} net={best.net:,} roi={best.roi*100:.2f}% params={best.params}"
        )


if __name__ == "__main__":
    main()
