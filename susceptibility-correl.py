import pandas as pd
from scipy.stats import spearmanr
import matplotlib.pyplot as plt
from pandas.plotting import table

# ファイルのパスを指定
file_path = 'stats.xlsx'

# データの読み込み
data = pd.read_excel(file_path)

# スコアリングのマッピング
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
data['empathized_score'] = data['empathized'].map(empathized_mapping)
data['change_score'] = data['change'].map(change_mapping)

# 相関の計算
empathized_anxiety_corr, empathized_anxiety_p = spearmanr(data['empathized_score'], data['anxiety'], nan_policy='omit')
empathized_change_corr, empathized_change_p = spearmanr(data['empathized_score'], data['change_score'], nan_policy='omit')
anxiety_change_corr, anxiety_change_p = spearmanr(data['anxiety'], data['change_score'], nan_policy='omit')

alpha = 0.05
empathized_anxiety_significant = empathized_anxiety_p < alpha
empathized_change_significant = empathized_change_p < alpha
anxiety_change_significant = anxiety_change_p < alpha

# 結果を表にまとめる
results_summary = pd.DataFrame({
    'Variable': ['Empathy & Anxiety', 'Empathy & Change', 'Anxiety & Change'],
    'Spearman Correlation': [round(empathized_anxiety_corr, 3), round(empathized_change_corr, 3), round(anxiety_change_corr, 3)],
    'p-value': [round(empathized_anxiety_p, 3), round(empathized_change_p, 3), round(anxiety_change_p, 3)],
    'Significant': ['Yes' if empathized_anxiety_significant else 'No', 'Yes' if empathized_change_significant else 'No', 'Yes' if anxiety_change_significant else 'No']
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