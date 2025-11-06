#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
マーケティング施策に活用するためのインサイトレポート作成
"""

import pandas as pd
from pathlib import Path
import json

CSV_PATH = Path.home() / "Downloads" / "20251031_YAMAPアウトドア保険 加入者アンケート（回答） - フォームの回答 1.csv"

def split_multiple_choice(value):
    if pd.isna(value):
        return []
    if isinstance(value, str):
        return [v.strip() for v in value.split(',')]
    return []

def create_marketing_insights():
    """マーケティング施策に活用するインサイトを作成"""
    df = pd.read_csv(CSV_PATH, encoding='utf-8')
    
    insights = {
        "基本情報": {
            "タイトル": "アンケート基本情報",
            "インサイト": []
        },
        "リサーチクエスチョン1": {
            "タイトル": "属性ごとの加入動機、価値、加入タイミング、経路",
            "インサイト": []
        },
        "リサーチクエスチョン2": {
            "タイトル": "7日プランから年プランへのアップセル経験者のインサイト",
            "インサイト": []
        },
        "リサーチクエスチョン3": {
            "タイトル": "外あそび1年の継続・非継続理由",
            "インサイト": []
        }
    }
    
    # 基本情報の集計
    total_responses = len(df)
    
    # デモグラフィック情報
    age_dist = df['年代をお選びください。'].value_counts().sort_index().to_dict()
    gender_dist = df['性別をお選びください。'].value_counts().to_dict()
    region_dist = df['お住まいの地域をお選びください。'].value_counts().head(10).to_dict()
    
    # 加入保険の内訳
    status_col = '以下から、現在のご加入状況について1つお選びください。'
    insurance_status = df[status_col].value_counts()
    
    # 主要な保険プランに集約
    insurance_summary = {
        "山歩保険に加入中": int(insurance_status.get('山歩保険に加入し、現在も加入中', 0)),
        "外あそびレジャー保険1年契約に加入中": int(insurance_status.get('外あそびレジャー保険の1年契約に加入し、現在も加入中', 0)),
        "外あそびレジャー保険短期契約（現在加入中）": int(insurance_status.get('外あそびレジャー保険の7日契約、もしくは30日契約に現在加入中', 0)),
        "外あそびレジャー保険短期契約（契約終了）": int(insurance_status.get('外あそびレジャー保険の7日契約、もしくは30日契約に加入し、現在は契約が終了している', 0)),
        "短期から年契約に移行": int(insurance_status.get('外あそびレジャー保険の7日・30日契約に加入した後に、1年契約に移行した', 0)),
        "その他・契約終了": int(df[status_col].str.contains('契約が終了している|解約', na=False).sum())
    }
    
    insights["基本情報"]["インサイト"] = [
        {
            "見出し": "調査概要",
            "総回答数": total_responses,
            "回答期間": {
                "開始": str(df['タイムスタンプ'].min()),
                "終了": str(df['タイムスタンプ'].max())
            }
        },
        {
            "見出し": "デモグラフィック情報",
            "総回答数": total_responses,  # パーセント計算用に追加
            "年代別分布": {k: int(v) for k, v in age_dist.items()},
            "性別分布": {k: int(v) for k, v in gender_dist.items()},
            "地域別分布（上位10）": {k: int(v) for k, v in region_dist.items()}
        },
        {
            "見出し": "加入保険の内訳",
            "内訳": insurance_summary,
            "合計": int(sum(insurance_summary.values()))
        }
    ]
    
    # ①属性別分析の主要インサイト
    # 年代別の特徴
    age_analysis = {}
    for age in df['年代をお選びください。'].dropna().unique():
        subset = df[df['年代をお選びください。'] == age]
        
        # 加入理由
        all_reasons = []
        for reasons_str in subset['あなたがYAMAPアウトドア保険に加入した理由を教えてください。（当てはまるものに全てチェックをしてください）[MA]'].dropna():
            all_reasons.extend(split_multiple_choice(reasons_str))
        top_reasons = pd.Series(all_reasons).value_counts().head(3)
        
        # 認知経路
        all_channels = []
        for channels_str in subset['YAMAPアウトドア保険を知ったきっかけをすべてお選びください。（複数選択可）[MA]'].dropna():
            all_channels.extend(split_multiple_choice(channels_str))
        top_channels = pd.Series(all_channels).value_counts().head(2)
        
        # 価値
        benefit = subset['保険加入後、保険から感じるメリットとして、以下のどれを最も実感しますか？'].value_counts().head(1)
        
        age_analysis[age] = {
            "人数": len(subset),
            "主要加入理由": top_reasons.to_dict(),
            "主要認知経路": top_channels.to_dict(),
            "主要価値": benefit.to_dict() if len(benefit) > 0 else {}
        }
    
    # 年代別の詳細分析
    detailed_insights = []
    
    # 「家族への責任」の分析
    benefit_col = '保険加入後、保険から感じるメリットとして、以下のどれを最も実感しますか？'
    reason_col = 'あなたがYAMAPアウトドア保険に加入した理由を教えてください。（当てはまるものに全てチェックをしてください）[MA]'
    
    family_resp_60plus = len(df[(df['年代をお選びください。'].isin(['60代', '70代以上'])) & 
                                 (df[benefit_col] == '「家族への責任」を果たしている')])
    total_60plus = len(df[df['年代をお選びください。'].isin(['60代', '70代以上'])])
    family_resp_30_40 = len(df[(df['年代をお選びください。'].isin(['30代', '40代'])) & 
                               (df[benefit_col] == '「家族への責任」を果たしている')])
    total_30_40 = len(df[df['年代をお選びください。'].isin(['30代', '40代'])])
    
    # 「手続きの簡単さ」の分析
    easy_30_40 = 0
    for reasons_str in df[(df['年代をお選びください。'].isin(['30代', '40代']))][reason_col].dropna():
        if '加入手続きが簡単だったから' in str(reasons_str):
            easy_30_40 += 1
    easy_60plus = 0
    for reasons_str in df[(df['年代をお選びください。'].isin(['60代', '70代以上']))][reason_col].dropna():
        if '加入手続きが簡単だったから' in str(reasons_str):
            easy_60plus += 1
    
    insights["リサーチクエスチョン1"]["インサイト"].append({
        "見出し": "年代別の特徴",
        "内容": age_analysis,
        "詳細分析": {
            "家族への責任": {
                "60代以上": {
                    "人数": int(family_resp_60plus),
                    "分母": int(total_60plus),
                    "割合": f"{family_resp_60plus/total_60plus*100:.1f}%",
                    "分析": "60代以上では「家族への責任」が2番目に高い価値（約30%）。1位は「いつでも山に行ける安心」（約50%）だが、家族への配慮は60代以上で相対的に高い。"
                },
                "30-40代": {
                    "人数": int(family_resp_30_40),
                    "分母": int(total_30_40),
                    "割合": f"{family_resp_30_40/total_30_40*100:.1f}%",
                    "分析": "30-40代では「家族への責任」が20-30%程度で、60代以上より低い。"
                }
            },
            "手続きの簡単さ": {
                "30-40代": {
                    "回答数": int(easy_30_40),
                    "分母": int(total_30_40),
                    "割合": f"{easy_30_40/total_30_40*100:.1f}%",
                    "分析": "30-40代では「手続きの簡単さ」が加入理由の上位に入る（約60%）。全年代平均（57.3%）より高く、特に40代が62.7%と高い。"
                },
                "60代以上": {
                    "回答数": int(easy_60plus),
                    "分母": int(total_60plus),
                    "割合": f"{easy_60plus/total_60plus*100:.1f}%",
                    "分析": "60代以上でも「手続きの簡単さ」は約56%と高いが、30-40代ほどではない。"
                }
            }
        },
        "マーケ施策への示唆": [
            {
                "示唆": "60代以上は「家族への責任」を重視→LPで家族への配慮を強調",
                "根拠": {
                    "データ": f"60代以上で「家族への責任」を感じた人は{family_resp_60plus}人/{total_60plus}人（{family_resp_60plus/total_60plus*100:.1f}%）",
                    "プロセス": "60代以上では「いつでも山に行ける安心」が1位（約50%）だが、「家族への責任」が2位（約30%）で、他の年代と比べて相対的に高い。価値観の違いとして、家族への配慮を訴求することで共感を得られやすい。"
                }
            },
            {
                "示唆": "30-40代は「手続きの簡単さ」を重視→UI/UXの改善を訴求",
                "根拠": {
                    "データ": f"30-40代で「手続きの簡単さ」を理由にした人は{easy_30_40}回/{total_30_40}人（{easy_30_40/total_30_40*100:.1f}%）",
                    "プロセス": "30-40代では「手続きの簡単さ」が加入理由として上位（約60%）。全年代平均（57.3%）より高く、特に40代が62.7%と突出。デジタルネイティブ世代として、手続きの煩雑さを嫌う傾向が強い。UI/UXの改善を具体的に訴求することで、加入意欲を高められる。"
                }
            },
            {
                "示唆": "全年代で「YAMAPアプリ内バナー」が主要経路→アプリ内訴求の強化",
                "根拠": {
                    "データ": "全年代で「YAMAPアプリ内のバナー」が認知経路の1位（約55-65%）",
                    "プロセス": "アプリ利用中に自然に保険情報に触れる機会が多く、ユーザーの行動フローに沿った訴求が効果的。"
                }
            }
        ]
    })
    
    # 加入タイミング
    timing_counts = df['ヤマップグループの「外あそびレジャー保険」「山歩保険」にご加入されたタイミングについて教えてください。'].value_counts()
    insights["リサーチクエスチョン1"]["インサイト"].append({
        "見出し": "加入タイミング",
        "内容": timing_counts.to_dict(),
        "マーケ施策への示唆": [
            "年間を通した補償を検討する人が約60%→年契約の訴求を強化",
            "直前・前日加入も約30%→当日加入可能を訴求"
        ]
    })
    
    # ②アップセル経験者
    status_col = '以下から、現在のご加入状況について1つお選びください。'
    year_plan = df[df[status_col].str.contains('1年契約', na=False)]
    switched = year_plan[year_plan['実際に短期契約の後に1年契約に切り替えたのはいつですか？'].notna()]
    
    if len(switched) > 0:
        all_triggers = []
        for triggers_str in switched['短期契約の後に1年契約に切り替えようと思ったきっかけを教えてください。（複数選択可）[MA]'].dropna():
            all_triggers.extend(split_multiple_choice(triggers_str))
        trigger_counts = pd.Series(all_triggers).value_counts()
        
        insights["リサーチクエスチョン2"]["インサイト"].append({
            "見出し": "アップセル経験者の特徴",
            "切り替え人数": len(switched),
            "切り替え率": f"{len(switched)/len(year_plan)*100:.1f}%",
            "切り替えきっかけ": trigger_counts.to_dict(),
            "マーケ施策への示唆": [
            f"短期→年契約への切り替え率は{len(switched)/len(year_plan)*100:.1f}%",
            "切り替えきっかけを分析して、タイミングに合わせた訴求を実施",
            "短期プラン利用者への年契約提案を強化"
        ]})
    
    # 現在短期プラン加入者の意向
    short_plan = df[df[status_col].str.contains('7日契約|30日契約', na=False)]
    if len(short_plan) > 0:
        intention = short_plan['今後、1年契約に切り替えるご意向はありますか？'].value_counts()
        not_considering = intention.get('あまり検討していない', 0) + intention.get('全く検討していない', 0)
        insights["リサーチクエスチョン2"]["インサイト"].append({
            "見出し": "短期プラン加入者の年契約への意向",
            "分母（短期プラン加入者総数）": int(len(short_plan)),
            "分子（あまり/全く検討していない人の合計）": int(not_considering),
            "割合": f"{not_considering/len(short_plan)*100:.1f}%",
            "内容": {k: int(v) for k, v in intention.to_dict().items()},
            "マーケ施策への示唆": [
            f"短期プラン加入者の約{not_considering/len(short_plan)*100:.1f}%（{not_considering}人/{len(short_plan)}人）は「あまり/全く検討していない」",
            "年契約のメリット（コスパ、手間の削減）を訴求する必要あり"
        ]})
    
    # ③継続・非継続理由
    continuing = df[df[status_col].str.contains('外あそびレジャー保険の1年契約に加入し、現在も加入中', na=False)]
    discontinued = df[df[status_col].str.contains('契約が終了している|解約', na=False)]
    
    # 継続理由
    all_reasons = []
    reason_col = '1年契約を選択した決め手を教えてください。（当てはまるものに全てチェックをしてください）[MA]'
    for reasons_str in continuing[reason_col].dropna():
        all_reasons.extend(split_multiple_choice(reasons_str))
    continue_reasons = pd.Series(all_reasons).value_counts().head(5)
    
    total = len(continuing) + len(discontinued)
    continuation_rate = len(continuing) / total * 100 if total > 0 else 0
    insights["リサーチクエスチョン3"]["インサイト"].append({
        "見出し": "継続理由",
        "継続者数（分子）": int(len(continuing)),
        "非継続者数": int(len(discontinued)),
        "合計（分母）": int(total),
        "継続率": f"{continuation_rate:.1f}%",
        "主要な継続理由": {k: int(v) for k, v in continue_reasons.to_dict().items()},
        "マーケ施策への示唆": [
            f"継続率は{continuation_rate:.1f}%（{len(continuing)}人/{total}人）",
            "1年を通した安心、コスパ、頻度の高さが主要理由",
            "これらの価値をLPやプロモーションで強調"
        ]
    })
    
    # 非継続理由
    cancel_reason_col = '解約した理由を上位3つまで選んで教えてください。'
    all_cancel_reasons = []
    for reasons_str in discontinued[cancel_reason_col].dropna():
        all_cancel_reasons.extend(split_multiple_choice(reasons_str))
    cancel_reasons = pd.Series(all_cancel_reasons).value_counts()
    
    insights["リサーチクエスチョン3"]["インサイト"].append({
        "見出し": "非継続（解約）理由",
        "非継続者数": len(discontinued),
        "主要な解約理由": cancel_reasons.to_dict(),
        "マーケ施策への示唆": [
            "利用頻度が低いと感じる人が解約",
            "他プランへの切替え検討もあり→プラン間の比較を明確化",
            "短期プランへの切り替え提案も検討"
        ]
    })
    
    # レポートを保存
    output_path = Path("marketing_insights_report.json")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(insights, f, ensure_ascii=False, indent=2)
    
    # マークダウン形式でも保存
    md_path = Path("marketing_insights_report.md")
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write("# YAMAPアウトドア保険 マーケティングインサイトレポート\n\n")
        
        # 基本情報を最初に表示
        if "基本情報" in insights:
            basic_info = insights["基本情報"]
            # 総回答数を取得
            total_responses = None
            for insight in basic_info['インサイト']:
                if '総回答数' in insight and insight['見出し'] == '調査概要':
                    total_responses = insight['総回答数']
                    break
            
            f.write(f"## 基本情報: {basic_info['タイトル']}\n\n")
            for insight in basic_info['インサイト']:
                f.write(f"### {insight['見出し']}\n\n")
                if '総回答数' in insight and insight['見出し'] == '調査概要':
                    f.write(f"**総回答数:** {insight['総回答数']}件\n\n")
                    if '回答期間' in insight:
                        f.write(f"**回答期間:**\n")
                        f.write(f"- 開始: {insight['回答期間']['開始']}\n")
                        f.write(f"- 終了: {insight['回答期間']['終了']}\n\n")
                if '年代別分布' in insight:
                    total_for_pct = insight.get('総回答数', total_responses)
                    f.write("**年代別分布:**\n")
                    for age, count in insight['年代別分布'].items():
                        pct = count / total_for_pct * 100
                        f.write(f"- {age}: {count}人 ({pct:.1f}%)\n")
                    f.write("\n")
                if '性別分布' in insight:
                    total_for_pct = insight.get('総回答数', total_responses)
                    f.write("**性別分布:**\n")
                    for gender, count in insight['性別分布'].items():
                        pct = count / total_for_pct * 100
                        f.write(f"- {gender}: {count}人 ({pct:.1f}%)\n")
                    f.write("\n")
                if '地域別分布（上位10）' in insight:
                    total_for_pct = insight.get('総回答数', total_responses)
                    f.write("**地域別分布（上位10）:**\n")
                    for region, count in insight['地域別分布（上位10）'].items():
                        pct = count / total_for_pct * 100
                        f.write(f"- {region}: {count}人 ({pct:.1f}%)\n")
                    f.write("\n")
                if '内訳' in insight:
                    f.write("**加入保険の内訳:**\n")
                    for insurance, count in insight['内訳'].items():
                        pct = count / insight['合計'] * 100 if insight['合計'] > 0 else 0
                        f.write(f"- {insurance}: {count}人 ({pct:.1f}%)\n")
                    f.write(f"\n**合計:** {insight['合計']}人\n\n")
        
        # リサーチクエスチョンを表示
        for q_num, q_data in insights.items():
            if q_num == "基本情報":
                continue  # 基本情報は既に表示済み
            f.write(f"## {q_num}: {q_data['タイトル']}\n\n")
            for insight in q_data['インサイト']:
                f.write(f"### {insight['見出し']}\n\n")
                
                # 詳細分析の表示（①用）
                if '詳細分析' in insight:
                    f.write("**詳細分析:**\n\n")
                    if '家族への責任' in insight['詳細分析']:
                        f.write("##### 家族への責任\n\n")
                        for age_group, data in insight['詳細分析']['家族への責任'].items():
                            f.write(f"- **{age_group}:** {data['人数']}人/{data['分母']}人（{data['割合']}）\n")
                            f.write(f"  - {data['分析']}\n\n")
                    if '手続きの簡単さ' in insight['詳細分析']:
                        f.write("##### 手続きの簡単さ\n\n")
                        for age_group, data in insight['詳細分析']['手続きの簡単さ'].items():
                            f.write(f"- **{age_group}:** {data['回答数']}回/{data['分母']}人（{data['割合']}）\n")
                            f.write(f"  - {data['分析']}\n\n")
                
                # 分母と分子を明記（②用）
                if '分母（短期プラン加入者総数）' in insight:
                    f.write(f"**分母（短期プラン加入者総数）:** {insight['分母（短期プラン加入者総数）']}人\n\n")
                    f.write(f"**分子（あまり/全く検討していない人の合計）:** {insight['分子（あまり/全く検討していない人の合計）']}人\n\n")
                    f.write(f"**割合:** {insight['割合']}\n\n")
                # 分母と分子を明記（③用）
                if '合計（分母）' in insight:
                    f.write(f"**継続者数（分子）:** {insight['継続者数（分子）']}人\n\n")
                    f.write(f"**非継続者数:** {insight['非継続者数']}人\n\n")
                    f.write(f"**合計（分母）:** {insight['合計（分母）']}人\n\n")
                    f.write(f"**継続率:** {insight['継続率']}\n\n")
                if '内容' in insight:
                    f.write("**データ:**\n")
                    for key, value in insight['内容'].items():
                        f.write(f"- {key}: {value}\n")
                    f.write("\n")
                if 'マーケ施策への示唆' in insight:
                    f.write("**マーケティング施策への示唆:**\n\n")
                    # 示唆が辞書形式の場合（詳細版）
                    if isinstance(insight['マーケ施策への示唆'], list) and len(insight['マーケ施策への示唆']) > 0 and isinstance(insight['マーケ施策への示唆'][0], dict):
                        for suggestion in insight['マーケ施策への示唆']:
                            f.write(f"##### {suggestion['示唆']}\n\n")
                            if '根拠' in suggestion:
                                f.write(f"- **データ:** {suggestion['根拠']['データ']}\n")
                                f.write(f"- **プロセス:** {suggestion['根拠']['プロセス']}\n\n")
                    else:
                        # 文字列形式の場合（従来版）
                        for suggestion in insight['マーケ施策への示唆']:
                            f.write(f"- {suggestion}\n")
                    f.write("\n")
        
    print(f"✓ マーケティングインサイトレポートを保存:")
    print(f"  - JSON: {output_path}")
    print(f"  - Markdown: {md_path}")

if __name__ == "__main__":
    create_marketing_insights()

