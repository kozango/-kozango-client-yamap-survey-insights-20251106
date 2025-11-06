#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
YAMAPアウトドア保険 加入者アンケート分析スクリプト
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
    print(f"列数: {len(df.columns)}")
    return df

def basic_statistics(df):
    """基本統計情報を表示"""
    print("\n" + "="*80)
    print("基本統計情報")
    print("="*80)
    print(f"総回答数: {len(df)}")
    print(f"回答期間: {df['タイムスタンプ'].min()} ～ {df['タイムスタンプ'].max()}")
    
    # 年代別の分布
    print("\n【年代別の分布】")
    age_counts = df['年代をお選びください。'].value_counts().sort_index()
    for age, count in age_counts.items():
        print(f"  {age}: {count}人 ({count/len(df)*100:.1f}%)")
    
    # 性別の分布
    print("\n【性別の分布】")
    gender_counts = df['性別をお選びください。'].value_counts()
    for gender, count in gender_counts.items():
        print(f"  {gender}: {count}人 ({count/len(df)*100:.1f}%)")
    
    # 地域別の分布
    print("\n【地域別の分布（上位10）】")
    region_counts = df['お住まいの地域をお選びください。'].value_counts().head(10)
    for region, count in region_counts.items():
        print(f"  {region}: {count}人 ({count/len(df)*100:.1f}%)")

def insurance_analysis(df):
    """保険関連の分析"""
    print("\n" + "="*80)
    print("保険関連の分析")
    print("="*80)
    
    # 加入タイミング
    print("\n【加入タイミング】")
    timing_col = 'ヤマップグループの「外あそびレジャー保険」「山歩保険」にご加入されたタイミングについて教えてください。'
    if timing_col in df.columns:
        timing_counts = df[timing_col].value_counts()
        for timing, count in timing_counts.items():
            print(f"  {timing}: {count}人 ({count/len(df)*100:.1f}%)")
    
    # 初めての加入かどうか
    print("\n【初めての登山保険加入かどうか】")
    first_col = '登山保険への加入は今回が初めてですか？'
    if first_col in df.columns:
        first_counts = df[first_col].value_counts()
        for val, count in first_counts.items():
            print(f"  {val}: {count}人 ({count/len(df)*100:.1f}%)")
    
    # 現在の加入状況
    print("\n【現在の加入状況】")
    status_col = '以下から、現在のご加入状況について1つお選びください。'
    if status_col in df.columns:
        status_counts = df[status_col].value_counts()
        for status, count in status_counts.items():
            print(f"  {status}: {count}人 ({count/len(df)*100:.1f}%)")

def satisfaction_analysis(df):
    """満足度・推奨度の分析"""
    print("\n" + "="*80)
    print("満足度・推奨度の分析")
    print("="*80)
    
    # 加入手続きの簡単さ
    print("\n【加入手続きの簡単さ】")
    easy_col = 'YAMAPアウトドア保険への加入手続きは簡単でしたか？'
    if easy_col in df.columns:
        easy_counts = df[easy_col].value_counts().sort_index()
        for val, count in easy_counts.items():
            print(f"  {val}: {count}人 ({count/len(df)*100:.1f}%)")
    
    # 推奨意向
    print("\n【家族・友人への推奨意向】")
    recommend_col = '加入中のYAMAPアウトドア保険を家族や友人、山仲間に勧めたいですか？'
    if recommend_col in df.columns:
        recommend_counts = df[recommend_col].value_counts()
        for val, count in recommend_counts.items():
            print(f"  {val}: {count}人 ({count/len(df)*100:.1f}%)")

def hiking_experience_analysis(df):
    """登山経験に関する分析"""
    print("\n" + "="*80)
    print("登山経験に関する分析")
    print("="*80)
    
    # 登山頻度
    print("\n【登山頻度】")
    freq_col = '直近1年以内に、どのくらいの頻度で登山・ハイキングをしていますか？'
    if freq_col in df.columns:
        freq_counts = df[freq_col].value_counts()
        for val, count in freq_counts.items():
            print(f"  {val}: {count}人 ({count/len(df)*100:.1f}%)")
    
    # 登山歴
    print("\n【登山歴】")
    history_col = 'あなたの登山歴に最も近いものをお選びください。'
    if history_col in df.columns:
        history_counts = df[history_col].value_counts()
        for val, count in history_counts.items():
            print(f"  {val}: {count}人 ({count/len(df)*100:.1f}%)")

def motivation_analysis(df):
    """加入動機の分析"""
    print("\n" + "="*80)
    print("加入動機の分析")
    print("="*80)
    
    # 加入理由
    reason_col = 'あなたがYAMAPアウトドア保険に加入した理由を教えてください。（当てはまるものに全てチェックをしてください）[MA]'
    if reason_col in df.columns:
        # 複数選択の回答を分割してカウント
        all_reasons = []
        for reasons_str in df[reason_col].dropna():
            if isinstance(reasons_str, str):
                reasons = [r.strip() for r in reasons_str.split(',')]
                all_reasons.extend(reasons)
        
        reason_counts = pd.Series(all_reasons).value_counts()
        print("\n【加入理由（複数選択可）】")
        for reason, count in reason_counts.head(10).items():
            print(f"  {reason}: {count}回 ({count/len(df)*100:.1f}%)")

def create_visualizations(df):
    """可視化を作成"""
    print("\n" + "="*80)
    print("グラフを作成しています...")
    print("="*80)
    
    fig_dir = Path("visualizations")
    fig_dir.mkdir(exist_ok=True)
    
    # 1. 年代別の分布
    plt.figure(figsize=(10, 6))
    age_counts = df['年代をお選びください。'].value_counts().sort_index()
    age_counts.plot(kind='bar', color='skyblue', edgecolor='black')
    plt.title('年代別の回答者分布', fontsize=14, fontweight='bold')
    plt.xlabel('年代', fontsize=12)
    plt.ylabel('回答者数', fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(fig_dir / 'age_distribution.png', dpi=300, bbox_inches='tight')
    print(f"✓ 年代分布グラフを保存: {fig_dir / 'age_distribution.png'}")
    plt.close()
    
    # 2. 性別の分布
    plt.figure(figsize=(8, 6))
    gender_counts = df['性別をお選びください。'].value_counts()
    plt.pie(gender_counts.values, labels=gender_counts.index, autopct='%1.1f%%', 
            startangle=90, colors=['lightblue', 'lightcoral'])
    plt.title('性別の分布', fontsize=14, fontweight='bold')
    plt.tight_layout()
    plt.savefig(fig_dir / 'gender_distribution.png', dpi=300, bbox_inches='tight')
    print(f"✓ 性別分布グラフを保存: {fig_dir / 'gender_distribution.png'}")
    plt.close()
    
    # 3. 地域別の分布（上位10）
    plt.figure(figsize=(12, 6))
    region_counts = df['お住まいの地域をお選びください。'].value_counts().head(10)
    region_counts.plot(kind='barh', color='lightgreen', edgecolor='black')
    plt.title('地域別の回答者分布（上位10）', fontsize=14, fontweight='bold')
    plt.xlabel('回答者数', fontsize=12)
    plt.ylabel('地域', fontsize=12)
    plt.tight_layout()
    plt.savefig(fig_dir / 'region_distribution.png', dpi=300, bbox_inches='tight')
    print(f"✓ 地域分布グラフを保存: {fig_dir / 'region_distribution.png'}")
    plt.close()

def main():
    """メイン処理"""
    print("="*80)
    print("YAMAPアウトドア保険 加入者アンケート分析")
    print("="*80)
    
    # データ読み込み
    df = load_data()
    
    # 基本統計
    basic_statistics(df)
    
    # 保険関連の分析
    insurance_analysis(df)
    
    # 満足度分析
    satisfaction_analysis(df)
    
    # 登山経験分析
    hiking_experience_analysis(df)
    
    # 加入動機分析
    motivation_analysis(df)
    
    # 可視化
    create_visualizations(df)
    
    # データの概要をCSVで保存
    summary_path = Path("data_summary.csv")
    summary_data = {
        '項目': ['総回答数', '年代数', '性別数', '地域数'],
        '値': [
            len(df),
            df['年代をお選びください。'].nunique(),
            df['性別をお選びください。'].nunique(),
            df['お住まいの地域をお選びください。'].nunique()
        ]
    }
    pd.DataFrame(summary_data).to_csv(summary_path, index=False, encoding='utf-8-sig')
    print(f"\n✓ サマリーデータを保存: {summary_path}")
    
    print("\n" + "="*80)
    print("分析が完了しました！")
    print("="*80)

if __name__ == "__main__":
    main()


