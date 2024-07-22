import pandas as pd
from scipy.stats import spearmanr
import matplotlib.pyplot as plt
from pandas.plotting import table

# ファイルのパスを指定
file_path = 'stats.xlsx'

# データの読み込み
data = pd.read_excel(file_path)

# 質問のスコアリング
media_frequency_mapping = {
    '毎日 / Every day': 4,
    '週に数回 / Several times per week': 3,
    '月に数回 / Several times per month': 2,
    'ほとんど利用しない / Seldom': 1
}
empathized_mapping = {
    'よくある / Often': 4,
    'たまにある / Sometimes': 3,
    'めったにない / Rarely': 2,
    'まったくない / Never': 1
}
change_mapping = {
    'はい / Yes': 1,
    'いいえ / No': 0
}

# スコアのマッピングを適用
data['media_frequency_score'] = data['frequency'].map(media_frequency_mapping)
data['empathized_score'] = data['empathized'].map(empathized_mapping)
data['change_score'] = data['change'].map(change_mapping)

# 相関の計算
empathized_corr, empathized_p = spearmanr(data['media_frequency_score'], data['empathized_score'], nan_policy='omit')
anxiety_corr, anxiety_p = spearmanr(data['media_frequency_score'], data['anxiety'], nan_policy='omit')
change_corr, change_p = spearmanr(data['media_frequency_score'], data['change_score'], nan_policy='omit')

alpha = 0.05
empathized_significant = empathized_p < alpha
anxiety_significant = anxiety_p < alpha
change_significant = change_p < alpha

# 結果を表にまとめる
results_summary = pd.DataFrame({
    'Variable': ['Freq & Empathy', 'Freq & Anxiety', 'Freq & Change'],
    'Spearman Correlation': [round(empathized_corr, 3), round(anxiety_corr, 3), round(change_corr, 3)],
    'p-value': [round(empathized_p, 3), round(anxiety_p, 3), round(change_p, 3)],
    'Significant': ['Yes' if empathized_significant else 'No', 'Yes' if anxiety_significant else 'No', 'Yes' if change_significant else 'No']
})

# プロットの作成
fig, ax = plt.subplots(figsize=(10, 4))  # プロットのサイズを指定
ax.axis('tight')
ax.axis('off')
tbl = table(ax, results_summary, loc='center', cellLoc='center', colWidths=[0.2]*len(results_summary.columns))
tbl.auto_set_font_size(False)
tbl.set_fontsize(12)
tbl.scale(1.2, 1.2)  # テーブルのサイズを調整

# プロットの表示
plt.title('Spearman Correlation Results Summary', fontsize=16)
plt.show()


result_lines = [
    f"Correlation between media frequency and empathized: {empathized_corr} (p-value: {empathized_p})",
    f"Correlation between media frequency and anxiety: {anxiety_corr} (p-value: {anxiety_p})",
    f"Correlation between media frequency and change: {change_corr} (p-value: {change_p})",
    f"Significant correlation (media frequency and empathized): {empathized_p < 0.05}",
    f"Significant correlation (media frequency and anxiety): {anxiety_p < 0.05}",
    f"Significant correlation (media frequency and change): {change_p < 0.05}"
]

# 結果をファイルに出力
output_file_path = 'correlation_results.txt'
with open(output_file_path, 'w') as f:
    for line in result_lines:
        f.write(line + '\n')

print(f"結果がファイルに出力されました: {output_file_path}")
