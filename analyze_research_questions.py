#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
YAMAPアウトドア保険 加入者アンケート分析スクリプト
リサーチクエスチョンに基づく詳細分析
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# 日本語フォントの設定
plt.rcParams['font.family'] = 'DejaVu Sans'
sns.set_style("whitegrid")
sns.set_palette("husl")

# CSVファイルのパス
CSV_PATH = Path.home() / "Downloads" / "20251031_YAMAPアウトドア保険 加入者アンケート（回答） - フォームの回答 1.csv"

def load_data():
    """データを読み込む"""
    print("データを読み込んでいます...")
    df = pd.read_csv(CSV_PATH, encoding='utf-8')
    print(f"データ読み込み完了: {len(df)}件の回答")
    return df

def split_multiple_choice(value):
    """複数選択の回答を分割"""
    if pd.isna(value):
        return []
    if isinstance(value, str):
        return [v.strip() for v in value.split(',')]
    return []

def analyze_by_attribute(df):
    """①属性ごとの加入動機、価値、加入タイミング、経路の分析"""
    print("\n" + "="*100)
    print("① 属性ごとの加入動機、価値（便益・独自性）、加入タイミング、経路の分析")
    print("="*100)
    
    results = {}
    
    # 属性定義
    attributes = {
        '年代': '年代をお選びください。',
        '性別': '性別をお選びください。',
        '地域': 'お住まいの地域をお選びください。',
        '登山歴': 'あなたの登山歴に最も近いものをお選びください。',
        '登山頻度': '直近1年以内に、どのくらいの頻度で登山・ハイキングをしていますか？'
    }
    
    # 分析対象の項目
    join_reason_col = 'あなたがYAMAPアウトドア保険に加入した理由を教えてください。（当てはまるものに全てチェックをしてください）[MA]'
    timing_col = 'ヤマップグループの「外あそびレジャー保険」「山歩保険」にご加入されたタイミングについて教えてください。'
    channel_col = 'YAMAPアウトドア保険を知ったきっかけをすべてお選びください。（複数選択可）[MA]'
    benefit_col = '保険加入後、保険から感じるメリットとして、以下のどれを最も実感しますか？'
    decision_col = '保険のご案内ページで、加入の「決め手となった情報」を1つ選んでお答えください。'
    
    for attr_name, attr_col in attributes.items():
        if attr_col not in df.columns:
            continue
        
        print(f"\n【{attr_name}別の分析】")
        print("-" * 80)
        
        attr_results = {}
        
        # 各属性値ごとに分析
        for attr_value in df[attr_col].dropna().unique():
            subset = df[df[attr_col] == attr_value]
            
            if len(subset) == 0:
                continue
            
            print(f"\n■ {attr_name}: {attr_value} (n={len(subset)})")
            
            # 加入理由
            if join_reason_col in df.columns:
                all_reasons = []
                for reasons_str in subset[join_reason_col].dropna():
                    all_reasons.extend(split_multiple_choice(reasons_str))
                if all_reasons:
                    reason_counts = pd.Series(all_reasons).value_counts()
                    print("\n  【加入理由（上位3）】")
                    for reason, count in reason_counts.head(3).items():
                        pct = count / len(subset) * 100
                        print(f"    {reason}: {count}回 ({pct:.1f}%)")
            
            # 加入タイミング
            if timing_col in df.columns:
                timing_counts = subset[timing_col].value_counts()
                print("\n  【加入タイミング】")
                for timing, count in timing_counts.items():
                    pct = count / len(subset) * 100
                    print(f"    {timing}: {count}人 ({pct:.1f}%)")
            
            # 認知経路
            if channel_col in df.columns:
                all_channels = []
                for channels_str in subset[channel_col].dropna():
                    all_channels.extend(split_multiple_choice(channels_str))
                if all_channels:
                    channel_counts = pd.Series(all_channels).value_counts()
                    print("\n  【認知経路（上位3）】")
                    for channel, count in channel_counts.head(3).items():
                        pct = count / len(subset) * 100
                        print(f"    {channel}: {count}回 ({pct:.1f}%)")
            
            # 価値（便益）
            if benefit_col in df.columns:
                benefit_counts = subset[benefit_col].value_counts()
                print("\n  【感じた価値・便益】")
                for benefit, count in benefit_counts.head(3).items():
                    pct = count / len(subset) * 100
                    print(f"    {benefit}: {count}人 ({pct:.1f}%)")
            
            # 決め手となった情報
            if decision_col in df.columns:
                decision_counts = subset[decision_col].value_counts()
                print("\n  【決め手となった情報】")
                for decision, count in decision_counts.head(3).items():
                    pct = count / len(subset) * 100
                    print(f"    {decision}: {count}人 ({pct:.1f}%)")
    
    return results

def analyze_upsell_experience(df):
    """②7日プランから年プランへのアップセル経験者のインサイト"""
    print("\n" + "="*100)
    print("② 7日プランから年プランへのアップセル経験者のインサイト")
    print("="*100)
    
    # 現在の加入状況から短期プラン経験者を特定
    status_col = '以下から、現在のご加入状況について1つお選びください。'
    switch_trigger_col = '短期契約の後に1年契約に切り替えようと思ったきっかけを教えてください。（複数選択可）[MA]'
    switch_timing_col = '実際に短期契約の後に1年契約に切り替えたのはいつですか？'
    hesitation_col = 'どのような点で迷われましたか？（複数選択可）[MA]'
    future_intention_col = '今後、1年契約に切り替えるご意向はありますか？'
    
    # アップセル経験者を特定
    # 現在年契約に加入している人で、短期プランを経験した人
    year_plan = df[df[status_col].str.contains('1年契約', na=False)]
    
    # 短期プラン経験者を特定（切り替えた人）
    switched = year_plan[year_plan[switch_timing_col].notna()]
    
    print(f"\n【アップセル経験者数】")
    print(f"  短期プランから年プランに切り替えた人: {len(switched)}人")
    print(f"  年契約加入者全体: {len(year_plan)}人")
    if len(year_plan) > 0:
        print(f"  切り替え率: {len(switched)/len(year_plan)*100:.1f}%")
    
    if len(switched) > 0:
        print(f"\n【切り替えきっかけ】")
        all_triggers = []
        for triggers_str in switched[switch_trigger_col].dropna():
            all_triggers.extend(split_multiple_choice(triggers_str))
        if all_triggers:
            trigger_counts = pd.Series(all_triggers).value_counts()
            for trigger, count in trigger_counts.items():
                pct = count / len(switched) * 100
                print(f"  {trigger}: {count}回 ({pct:.1f}%)")
        
        print(f"\n【切り替えタイミング】")
        timing_counts = switched[switch_timing_col].value_counts()
        for timing, count in timing_counts.items():
            pct = count / len(switched) * 100
            print(f"  {timing}: {count}人 ({pct:.1f}%)")
        
        # 迷った点
        print(f"\n【迷った点（短期→年契約への切り替え時）】")
        all_hesitations = []
        for hesitations_str in switched[hesitation_col].dropna():
            all_hesitations.extend(split_multiple_choice(hesitations_str))
        if all_hesitations:
            hesitation_counts = pd.Series(all_hesitations).value_counts()
            for hesitation, count in hesitation_counts.items():
                pct = count / len(switched) * 100
                print(f"  {hesitation}: {count}回 ({pct:.1f}%)")
    
    # 現在短期プランに加入している人の将来意向
    short_plan = df[df[status_col].str.contains('7日契約|30日契約', na=False)]
    if len(short_plan) > 0:
        print(f"\n【現在短期プラン加入者の年契約への切り替え意向】")
        print(f"  分母（短期プラン加入者総数）: {len(short_plan)}人")
        intention_counts = short_plan[future_intention_col].value_counts()
        for intention, count in intention_counts.items():
            pct = count / len(short_plan) * 100
            print(f"  {intention}: {count}人 ({pct:.1f}%)")
        
        # あまり/全く検討していない人の合計
        not_considering = intention_counts.get('あまり検討していない', 0) + intention_counts.get('全く検討していない', 0)
        print(f"\n  【あまり/全く検討していない人の合計】")
        print(f"    分子: {not_considering}人")
        print(f"    分母: {len(short_plan)}人")
        print(f"    割合: {not_considering/len(short_plan)*100:.1f}%")
    
    return switched

def analyze_continuation(df):
    """③外あそび1年の継続・非継続理由"""
    print("\n" + "="*100)
    print("③ 外あそびレジャー保険1年契約の継続・非継続理由")
    print("="*100)
    
    status_col = '以下から、現在のご加入状況について1つお選びください。'
    
    # 継続している人
    continuing = df[df[status_col].str.contains('外あそびレジャー保険の1年契約に加入し、現在も加入中', na=False)]
    
    # 非継続した人（解約した人）
    discontinued = df[df[status_col].str.contains('契約が終了している|解約', na=False)]
    
    # 継続率の計算
    total = len(continuing) + len(discontinued)
    if total > 0:
        continuation_rate = len(continuing) / total * 100
        print(f"\n【継続率】")
        print(f"  継続者数（分子）: {len(continuing)}人")
        print(f"  非継続者数: {len(discontinued)}人")
        print(f"  合計（分母）: {total}人")
        print(f"  継続率: {continuation_rate:.1f}%")
        print(f"    = {len(continuing)}人 / {total}人")
    
    # 継続理由を分析
    if len(continuing) > 0:
        print(f"\n【継続している人】")
        print(f"  継続者数: {len(continuing)}人")
        
        # 1年契約を選択した決め手
        reason_col = '1年契約を選択した決め手を教えてください。（当てはまるものに全てチェックをしてください）[MA]'
        if reason_col in df.columns:
            all_reasons = []
            for reasons_str in continuing[reason_col].dropna():
                all_reasons.extend(split_multiple_choice(reasons_str))
            if all_reasons:
                reason_counts = pd.Series(all_reasons).value_counts()
                print("\n  【1年契約を選んだ決め手】")
                for reason, count in reason_counts.items():
                    pct = count / len(continuing) * 100
                    print(f"    {reason}: {count}回 ({pct:.1f}%)")
        
        # 属性別の継続者特徴
        print("\n  【継続者の属性特徴】")
        print(f"    年代: {continuing['年代をお選びください。'].value_counts().to_dict()}")
        print(f"    性別: {continuing['性別をお選びください。'].value_counts().to_dict()}")
        print(f"    登山頻度: {continuing['直近1年以内に、どのくらいの頻度で登山・ハイキングをしていますか？'].value_counts().to_dict()}")
    
    # 非継続理由を分析
    if len(discontinued) > 0:
        print(f"\n【非継続（解約）した人】")
        print(f"  非継続者数: {len(discontinued)}人")
        
        cancel_reason_col = '解約した理由を上位3つまで選んで教えてください。'
        if cancel_reason_col in df.columns:
            all_reasons = []
            for reasons_str in discontinued[cancel_reason_col].dropna():
                all_reasons.extend(split_multiple_choice(reasons_str))
            if all_reasons:
                reason_counts = pd.Series(all_reasons).value_counts()
                print("\n  【解約理由】")
                for reason, count in reason_counts.items():
                    pct = count / len(discontinued) * 100
                    print(f"    {reason}: {count}回 ({pct:.1f}%)")
        
        # 解約理由の詳細
        detail_col = '上記で選んだ選択肢について、より具体的に教えてください。'
        if detail_col in df.columns:
            details = discontinued[detail_col].dropna()
            if len(details) > 0:
                print(f"\n  【解約理由の詳細（例）】")
                for idx, detail in details.head(5).items():
                    if detail and len(str(detail)) > 10:
                        print(f"    - {str(detail)[:100]}...")
        
        # 属性別の非継続者特徴
        print("\n  【非継続者の属性特徴】")
        print(f"    年代: {discontinued['年代をお選びください。'].value_counts().to_dict()}")
        print(f"    性別: {discontinued['性別をお選びください。'].value_counts().to_dict()}")
        print(f"    登山頻度: {discontinued['直近1年以内に、どのくらいの頻度で登山・ハイキングをしていますか？'].value_counts().to_dict()}")
    
    return continuing, discontinued

def create_summary_report(df, switched, continuing, discontinued):
    """サマリーレポートを作成"""
    print("\n" + "="*100)
    print("サマリーレポートの作成")
    print("="*100)
    
    # Excel形式でレポートを作成
    output_path = Path("yamap_analysis_report.xlsx")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 基本統計
        summary_data = {
            '項目': ['総回答数', '年契約継続者', '年契約非継続者', 'アップセル経験者'],
            '人数': [
                len(df),
                len(continuing),
                len(discontinued),
                len(switched)
            ],
            '割合': [
                100.0,
                len(continuing)/len(df)*100 if len(df) > 0 else 0,
                len(discontinued)/len(df)*100 if len(df) > 0 else 0,
                len(switched)/len(df)*100 if len(df) > 0 else 0
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='サマリー', index=False)
        
        # 属性別集計
        attr_summary = df.groupby('年代をお選びください。').agg({
            'ユーザーID': 'count'
        }).reset_index()
        attr_summary.columns = ['年代', '回答者数']
        attr_summary.to_excel(writer, sheet_name='属性別集計', index=False)
    
    print(f"✓ レポートを保存: {output_path}")

def create_visualizations(df, switched, continuing, discontinued):
    """可視化を作成"""
    print("\n" + "="*100)
    print("グラフを作成しています...")
    print("="*100)
    
    fig_dir = Path("visualizations")
    fig_dir.mkdir(exist_ok=True)
    
    # 1. 加入タイミングの分布
    timing_col = 'ヤマップグループの「外あそびレジャー保険」「山歩保険」にご加入されたタイミングについて教えてください。'
    if timing_col in df.columns:
        plt.figure(figsize=(12, 6))
        timing_counts = df[timing_col].value_counts()
        timing_counts.plot(kind='barh', color='skyblue', edgecolor='black')
        plt.title('加入タイミングの分布', fontsize=14, fontweight='bold')
        plt.xlabel('回答者数', fontsize=12)
        plt.tight_layout()
        plt.savefig(fig_dir / 'joining_timing.png', dpi=300, bbox_inches='tight')
        print(f"✓ 加入タイミング分布: {fig_dir / 'joining_timing.png'}")
        plt.close()
    
    # 2. 認知経路の分布
    channel_col = 'YAMAPアウトドア保険を知ったきっかけをすべてお選びください。（複数選択可）[MA]'
    if channel_col in df.columns:
        plt.figure(figsize=(12, 8))
        all_channels = []
        for channels_str in df[channel_col].dropna():
            all_channels.extend(split_multiple_choice(channels_str))
        channel_counts = pd.Series(all_channels).value_counts()
        channel_counts.plot(kind='barh', color='lightgreen', edgecolor='black')
        plt.title('認知経路の分布', fontsize=14, fontweight='bold')
        plt.xlabel('回答数', fontsize=12)
        plt.tight_layout()
        plt.savefig(fig_dir / 'channel_distribution.png', dpi=300, bbox_inches='tight')
        print(f"✓ 認知経路分布: {fig_dir / 'channel_distribution.png'}")
        plt.close()
    
    # 3. 継続 vs 非継続の比較
    if len(continuing) > 0 and len(discontinued) > 0:
        plt.figure(figsize=(10, 6))
        status_data = pd.DataFrame({
            '継続': [len(continuing)],
            '非継続': [len(discontinued)]
        })
        status_data.T.plot(kind='bar', color=['green', 'red'], edgecolor='black')
        plt.title('1年契約の継続 vs 非継続', fontsize=14, fontweight='bold')
        plt.ylabel('人数', fontsize=12)
        plt.xticks(rotation=0)
        plt.legend().remove()
        plt.tight_layout()
        plt.savefig(fig_dir / 'continuation_status.png', dpi=300, bbox_inches='tight')
        print(f"✓ 継続状況: {fig_dir / 'continuation_status.png'}")
        plt.close()

def main():
    """メイン処理"""
    print("="*100)
    print("YAMAPアウトドア保険 加入者アンケート分析")
    print("リサーチクエスチョンに基づく詳細分析")
    print("="*100)
    
    # データ読み込み
    df = load_data()
    
    # ①属性ごとの加入動機、価値、加入タイミング、経路の分析
    analyze_by_attribute(df)
    
    # ②アップセル経験者のインサイト
    switched = analyze_upsell_experience(df)
    
    # ③継続・非継続理由
    continuing, discontinued = analyze_continuation(df)
    
    # サマリーレポート作成
    create_summary_report(df, switched, continuing, discontinued)
    
    # 可視化
    create_visualizations(df, switched, continuing, discontinued)
    
    print("\n" + "="*100)
    print("分析が完了しました！")
    print("="*100)

if __name__ == "__main__":
    main()

