#!/usr/bin/env python3
# DJI T25 Flight Analyzer - fixed local/Termux version
# Usage:
#   python3 dji_analyzer_fixed.py FlightRecord.xlsx [more_FlightRecord.xlsx]
# Output:
#   t25_flight_flags.csv
#   t25_inferred_battery_segments.csv
#
# 修正重點：
# 1) 不再把 Serial Number 誤判成 Battery SN。Battery SN 空白時，只做「逐趟/推定電池段」判斷。
# 2) 支援 compact / full DJI FlightRecord 欄位。
# 3) 缺起降電量時，不硬算老化電池。
# 4) 可一次丟多個 xlsx，會用 Flight time + Serial Number 去重。

import sys
from pathlib import Path
import pandas as pd

FEN_M2 = 969.917
JIA_M2 = 9699.17


def parse_dur(value):
    if pd.isna(value):
        return 0.0
    parts = str(value).split(':')
    if len(parts) == 2:
        return int(parts[0]) + int(parts[1]) / 60.0
    try:
        return float(value)
    except Exception:
        return 0.0


def parse_time(value):
    try:
        date_part, range_part = str(value).strip().rsplit(' ', 1)
        start_s, end_s = range_part.split('-')
        return pd.to_datetime(f'{date_part} {start_s}'), pd.to_datetime(f'{date_part} {end_s}')
    except Exception:
        return pd.NaT, pd.NaT


def loc_short(loc):
    loc = '' if pd.isna(loc) else str(loc)
    if '南興街' in loc:
        return '福興南興街'
    if 'Hanbao' in loc or '漢寶' in loc or 'Fangyuan' in loc:
        return '芳苑漢寶'
    if 'Pouzi' in loc or 'Puyan' in loc:
        return '埔鹽埔子'
    if '福三路' in loc:
        return '福興福三路'
    if '鹿和路' in loc or 'Lugang' in loc:
        return '鹿港'
    if 'Zhanglu' in loc or 'Fuxing' in loc:
        return '福興'
    return loc[:18]


def load_files(paths):
    frames = []
    for p in paths:
        df = pd.read_excel(p)
        df['Source File'] = Path(p).name
        frames.append(df)
    if not frames:
        raise SystemExit('請提供 FlightRecord.xlsx 檔案路徑')
    df = pd.concat(frames, ignore_index=True)
    dedup_cols = [c for c in ['Flight time', 'Serial Number'] if c in df.columns]
    if dedup_cols:
        df = df.drop_duplicates(subset=dedup_cols, keep='first')
    return df


def enrich(df):
    required = ['Flight time', 'Sprayed area', 'Total Amount(L/Kg)', 'Flight duration(min:sec)']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise SystemExit(f'缺少必要欄位: {missing}')

    starts, ends = zip(*df['Flight time'].map(parse_time))
    df = df.copy()
    df['start_dt'] = starts
    df['end_dt'] = ends
    df['date'] = pd.to_datetime(df['start_dt']).dt.date.astype(str)
    df['dur_min'] = df['Flight duration(min:sec)'].map(parse_dur)
    df['area_m2'] = pd.to_numeric(df['Sprayed area'], errors='coerce').fillna(0)
    df['area_fen'] = df['area_m2'] / FEN_M2
    df['amount_l'] = pd.to_numeric(df['Total Amount(L/Kg)'], errors='coerce').fillna(0)
    df['amount_l_per_fen'] = df['amount_l'] / df['area_fen'].replace(0, pd.NA)
    df['location_short'] = df.get('Location', '').map(loc_short) if 'Location' in df.columns else ''

    df['start_batt'] = pd.to_numeric(df.get('Starting Battery Level'), errors='coerce')
    df['end_batt'] = pd.to_numeric(df.get('Ending Battery Level'), errors='coerce')
    df['drop_pct'] = df['start_batt'] - df['end_batt']
    df['drain_pct_min'] = df['drop_pct'] / df['dur_min'].replace(0, pd.NA)
    df['eff_m2_min'] = df['area_m2'] / df['dur_min'].replace(0, pd.NA)

    def flags(r):
        out = []
        if r['area_m2'] == 0:
            out.append('空飛/測試')
        if pd.notna(r['end_batt']) and r['end_batt'] <= 20:
            out.append('低電降落<=20')
        if pd.notna(r['drain_pct_min']) and r['drain_pct_min'] > 9:
            out.append('高耗電>9%/min')
        if pd.notna(r['eff_m2_min']) and r['area_m2'] > 0 and r['eff_m2_min'] < 150:
            out.append('效率低')
        if pd.isna(r['start_batt']) or pd.isna(r['end_batt']):
            out.append('缺電量資料')
        if pd.notna(r['start_batt']) and r['start_batt'] >= 96:
            out.append('滿電起飛')
        if pd.notna(r['start_batt']) and r['start_batt'] < 30:
            out.append('低電起飛<30')
        return '、'.join(out)

    df['flags'] = df.apply(flags, axis=1)
    return df.sort_values('start_dt').reset_index(drop=True)


def infer_segments(df, jump_threshold=5):
    rows = []
    current = []
    cid = 0
    last_end = None

    for _, r in df.iterrows():
        missing = pd.isna(r['start_batt']) or pd.isna(r['end_batt'])
        if missing:
            if current:
                rows.append((cid, pd.DataFrame(current)))
                current = []
            cid += 1
            rows.append((cid, pd.DataFrame([r])))
            last_end = None
            continue
        if not current:
            cid += 1
            current = [r]
            last_end = r['end_batt']
            continue
        if pd.notna(last_end) and r['start_batt'] > last_end + jump_threshold:
            rows.append((cid, pd.DataFrame(current)))
            cid += 1
            current = [r]
        else:
            current.append(r)
        last_end = r['end_batt']
    if current:
        rows.append((cid, pd.DataFrame(current)))

    summary = []
    for cid, g in rows:
        start_batt = g.iloc[0]['start_batt']
        end_batt = g.iloc[-1]['end_batt']
        drop = start_batt - end_batt if pd.notna(start_batt) and pd.notna(end_batt) else pd.NA
        dur = g['dur_min'].sum()
        drain = drop / dur if pd.notna(drop) and dur > 0 else pd.NA
        flags = []
        if pd.notna(end_batt) and end_batt <= 20:
            flags.append('低電收尾')
        if pd.notna(drain) and drain > 9:
            flags.append('整段高耗電')
        if g['start_batt'].isna().any() or g['end_batt'].isna().any():
            flags.append('缺電量')
        summary.append({
            'segment': cid,
            'date': str(g.iloc[0]['date']),
            'start_time': str(g.iloc[0]['start_dt'].time()) if pd.notna(g.iloc[0]['start_dt']) else '',
            'end_time': str(g.iloc[-1]['end_dt'].time()) if pd.notna(g.iloc[-1]['end_dt']) else '',
            'location': g.iloc[0]['location_short'],
            'flights': len(g),
            'area_m2': g['area_m2'].sum(),
            'area_fen': g['area_fen'].sum(),
            'amount_l': g['amount_l'].sum(),
            'dur_min': dur,
            'start_batt': start_batt,
            'end_batt': end_batt,
            'drop_pct': drop,
            'drain_pct_min': drain,
            'empty_flights': int((g['area_m2'] == 0).sum()),
            'flags': '、'.join(flags),
        })
    return pd.DataFrame(summary)


def print_summary(df, segments):
    spray = df[df['area_m2'] > 0]
    batt = df[df['drop_pct'].notna()]
    print('\n=== T25 作業總結 ===')
    print(f"總趟次: {len(df)} | 有效噴灑: {len(spray)} | 空飛/測試: {(df['area_m2']==0).sum()}")
    print(f"總面積: {df['area_m2'].sum():.0f} m² = {df['area_fen'].sum():.2f} 分 = {df['area_m2'].sum()/JIA_M2:.2f} 甲")
    print(f"總用藥: {df['amount_l'].sum():.2f} L | 平均 {df['amount_l'].sum()/df['area_fen'].sum():.2f} L/分")
    print(f"總飛行: {df['dur_min'].sum():.1f} min | 有效效率: {df['area_m2'].sum()/spray['dur_min'].sum():.0f} m²/min")
    if len(batt):
        print(f"平均耗電: {batt['drop_pct'].sum()/batt['dur_min'].sum():.2f}%/min")
    print(f"低電降落(<=20%): {(df['end_batt']<=20).sum()} 趟")
    print(f"高耗電(>9%/min): {(df['drain_pct_min']>9).sum()} 趟")
    print(f"缺電量資料: {df['drop_pct'].isna().sum()} 趟")

    if 'Battery SN' in df.columns and df['Battery SN'].notna().any():
        print('\nBattery SN 有資料，可另做實體電池排名。')
    else:
        print('\n注意: Battery SN 欄位沒有資料。不要把 Serial Number 當電池編號。')
        print('以下只輸出「推定電池段」，不是實體電池健康排名。')

    print('\n=== 推定電池段 ===')
    show = segments.copy()
    for c in ['area_m2', 'area_fen', 'amount_l', 'dur_min', 'drain_pct_min']:
        if c in show.columns:
            show[c] = pd.to_numeric(show[c], errors='coerce').round(2)
    print(show.to_string(index=False))

    print('\n=== 需要回看逐趟 ===')
    flagged = df[df['flags'].astype(str).str.len() > 0]
    cols = ['Flight time','location_short','area_m2','dur_min','start_batt','end_batt','drain_pct_min','eff_m2_min','flags']
    tmp = flagged[cols].copy()
    for c in ['area_m2','dur_min','drain_pct_min','eff_m2_min']:
        tmp[c] = pd.to_numeric(tmp[c], errors='coerce').round(2)
    print(tmp.to_string(index=False))


def main():
    paths = sys.argv[1:] or ['FlightRecord.xlsx']
    missing = [p for p in paths if not Path(p).exists()]
    if missing:
        raise SystemExit(f'找不到檔案: {missing}')
    df = enrich(load_files(paths))
    segments = infer_segments(df)
    print_summary(df, segments)

    cols = ['Source File','Flight time','location_short','Aircraft name','area_m2','area_fen','amount_l','amount_l_per_fen','dur_min','Serial Number','Battery SN','start_batt','end_batt','drop_pct','drain_pct_min','eff_m2_min','flags']
    cols = [c for c in cols if c in df.columns]
    df[cols].to_csv('t25_flight_flags.csv', index=False, encoding='utf-8-sig')
    segments.to_csv('t25_inferred_battery_segments.csv', index=False, encoding='utf-8-sig')
    print('\n已輸出: t25_flight_flags.csv, t25_inferred_battery_segments.csv')


if __name__ == '__main__':
    main()
