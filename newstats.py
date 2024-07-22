import pandas as pd
from scipy.stats import chi2_contingency
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from pandas.plotting import table

# ファイルのパスを指定
file_path = 'stats.xlsx'

# データの読み込み
data = pd.read_excel(file_path)

# くらめーる
def cramer(ct, chi2):
    n = ct.sum().sum()
    k = min(ct.shape)
    return np.sqrt(chi2 / (n * (k - 1)))

# クロス集計表の作成とカイ二乗検定の実行
def perform_chi2_test(column):
    empathized_ct = pd.crosstab(data[column], data['empathized'])
    anxiety_ct = pd.crosstab(data[column], data['anxiety'])
    change_ct = pd.crosstab(data[column], data['change'])

    empathized_chi2, empathized_p, _, _ = chi2_contingency(empathized_ct)
    anxiety_chi2, anxiety_p, _, _ = chi2_contingency(anxiety_ct)
    change_chi2, change_p, _, _ = chi2_contingency(change_ct)

    return {
        'empathized': (empathized_chi2, empathized_p, empathized_p < 0.05, cramer(empathized_ct, empathized_chi2)),
        'anxiety': (anxiety_chi2, anxiety_p, anxiety_p < 0.05, cramer(anxiety_ct, anxiety_chi2)), 
        'change': (change_chi2, change_p, change_p < 0.05, cramer(change_ct, change_chi2))
    }

# 結果の集計
age_results = perform_chi2_test('age')
gender_results = perform_chi2_test('gender')
occupation_results = perform_chi2_test('occupation')
source_results = perform_chi2_test('source')

# 結果の出力
result_lines = [
    f"Chi-Square Test Results for Age:\n"
    f"  Empathized: Chi2: {age_results['empathized'][0]}, p-value: {age_results['empathized'][1]}, Significant: {age_results['empathized'][2]}, Cramér's V: {age_results['empathized'][3]}\n"
    f"  Anxiety: Chi2: {age_results['anxiety'][0]}, p-value: {age_results['anxiety'][1]}, Significant: {age_results['anxiety'][2]}, Cramér's V: {age_results['anxiety'][3]}\n"
    f"  Change: Chi2: {age_results['change'][0]}, p-value: {age_results['change'][1]}, Significant: {age_results['change'][2]}, Cramér's V: {age_results['change'][3]}\n",
    
    f"Chi-Square Test Results for Gender:\n"
    f"  Empathized: Chi2: {gender_results['empathized'][0]}, p-value: {gender_results['empathized'][1]}, Significant: {gender_results['empathized'][2]}, Cramér's V: {gender_results['empathized'][3]}\n"
    f"  Anxiety: Chi2: {gender_results['anxiety'][0]}, p-value: {gender_results['anxiety'][1]}, Significant: {gender_results['anxiety'][2]}, Cramér's V: {gender_results['anxiety'][3]}\n"
    f"  Change: Chi2: {gender_results['change'][0]}, p-value: {gender_results['change'][1]}, Significant: {gender_results['change'][2]}, Cramér's V: {gender_results['change'][3]}\n",
    
    f"Chi-Square Test Results for Occupation:\n"
    f"  Empathized: Chi2: {occupation_results['empathized'][0]}, p-value: {occupation_results['empathized'][1]}, Significant: {occupation_results['empathized'][2]}, Cramér's V: {occupation_results['empathized'][3]}\n"
    f"  Anxiety: Chi2: {occupation_results['anxiety'][0]}, p-value: {occupation_results['anxiety'][1]}, Significant: {occupation_results['anxiety'][2]}, Cramér's V: {occupation_results['anxiety'][3]}\n"
    f"  Change: Chi2: {occupation_results['change'][0]}, p-value: {occupation_results['change'][1]}, Significant: {occupation_results['change'][2]}, Cramér's V: {occupation_results['change'][3]}\n",
    
    f"Chi-Square Test Results for Source:\n"
    f"  Empathized: Chi2: {source_results['empathized'][0]}, p-value: {source_results['empathized'][1]}, Significant: {source_results['empathized'][2]}, Cramér's V: {source_results['empathized'][3]}\n"
    f"  Anxiety: Chi2: {source_results['anxiety'][0]}, p-value: {source_results['anxiety'][1]}, Significant: {source_results['anxiety'][2]}, Cramér's V: {source_results['anxiety'][3]}\n"
    f"  Change: Chi2: {source_results['change'][0]}, p-value: {source_results['change'][1]}, Significant: {source_results['change'][2]}, Cramér's V: {source_results['change'][3]}\n"
]

# 結果をファイルに出力
output_file_path = 'chi2_test_results_all_respondents.txt'
with open(output_file_path, 'w', encoding='utf-8') as f:
    for line in result_lines:
        f.write(line + '\n')

print(f"結果がファイルに出力されました: {output_file_path}")

################

results_df = pd.DataFrame({
    'Variable': ['Age', 'Gender', 'Occupation', 'Source'],
    'Empathized Chi2': [round(age_results['empathized'][0], 3), round(gender_results['empathized'][0], 3), round(occupation_results['empathized'][0], 3), round(source_results['empathized'][0], 3)],
    'Empathized p-value': [round(age_results['empathized'][1], 3), round(gender_results['empathized'][1], 3), round(occupation_results['empathized'][1], 3), round(source_results['empathized'][1], 3)],
    'Empathized Significant': [age_results['empathized'][2], gender_results['empathized'][2], occupation_results['empathized'][2], source_results['empathized'][2]],
    'Empathized Cramer\'s V': [round(age_results['empathized'][3], 3), round(gender_results['empathized'][3], 3), round(occupation_results['empathized'][3], 3), round(source_results['empathized'][3], 3)],
    'Anxiety Chi2': [round(age_results['anxiety'][0], 3), round(gender_results['anxiety'][0], 3), round(occupation_results['anxiety'][0], 3), round(source_results['anxiety'][0], 3)],
    'Anxiety p-value': [round(age_results['anxiety'][1], 3), round(gender_results['anxiety'][1], 3), round(occupation_results['anxiety'][1], 3), round(source_results['anxiety'][1], 3)],
    'Anxiety Significant': [age_results['anxiety'][2], gender_results['anxiety'][2], occupation_results['anxiety'][2], source_results['anxiety'][2]],
    'Anxiety Cramer\'s V': [round(age_results['anxiety'][3], 3), round(gender_results['anxiety'][3], 3), round(occupation_results['anxiety'][3], 3), round(source_results['anxiety'][3], 3)],
    'Change Chi2': [round(age_results['change'][0], 3), round(gender_results['change'][0], 3), round(occupation_results['change'][0], 3), round(source_results['change'][0], 3)],
    'Change p-value': [round(age_results['change'][1], 3), round(gender_results['change'][1], 3), round(occupation_results['change'][1], 3), round(source_results['change'][1], 3)],
    'Change Significant': [age_results['change'][2], gender_results['change'][2], occupation_results['change'][2], source_results['change'][2]],
    'Change Cramer\'s V': [round(age_results['change'][3], 3), round(gender_results['change'][3], 3), round(occupation_results['change'][3], 3), round(source_results['change'][3], 3)],
})

results_emp = pd.DataFrame({
    'Variable': ['Age', 'Gender', 'Occupation', 'Source'],
    'Empathized Chi2': [round(age_results['empathized'][0], 3), round(gender_results['empathized'][0], 3), round(occupation_results['empathized'][0], 3), round(source_results['empathized'][0], 3)],
    'Empathized p-value': [round(age_results['empathized'][1], 3), round(gender_results['empathized'][1], 3), round(occupation_results['empathized'][1], 3), round(source_results['empathized'][1], 3)],
    'Empathized Significant': [age_results['empathized'][2], gender_results['empathized'][2], occupation_results['empathized'][2], source_results['empathized'][2]],
    'Empathized Cramer\'s V': [round(age_results['empathized'][3], 3), round(gender_results['empathized'][3], 3), round(occupation_results['empathized'][3], 3), round(source_results['empathized'][3], 3)],
})

results_anx = pd.DataFrame({
    'Variable': ['Age', 'Gender', 'Occupation', 'Source'],
    'Anxiety Chi2': [round(age_results['anxiety'][0], 3), round(gender_results['anxiety'][0], 3), round(occupation_results['anxiety'][0], 3), round(source_results['anxiety'][0], 3)],
    'Anxiety p-value': [round(age_results['anxiety'][1], 3), round(gender_results['anxiety'][1], 3), round(occupation_results['anxiety'][1], 3), round(source_results['anxiety'][1], 3)],
    'Anxiety Significant': [age_results['anxiety'][2], gender_results['anxiety'][2], occupation_results['anxiety'][2], source_results['anxiety'][2]],
    'Anxiety Cramer\'s V': [round(age_results['anxiety'][3], 3), round(gender_results['anxiety'][3], 3), round(occupation_results['anxiety'][3], 3), round(source_results['anxiety'][3], 3)],
})

results_cha = pd.DataFrame({
    'Variable': ['Age', 'Gender', 'Occupation', 'Source'],
    'Change Chi2': [round(age_results['change'][0], 3), round(gender_results['change'][0], 3), round(occupation_results['change'][0], 3), round(source_results['change'][0], 3)],
    'Change p-value': [round(age_results['change'][1], 3), round(gender_results['change'][1], 3), round(occupation_results['change'][1], 3), round(source_results['change'][1], 3)],
    'Change Significant': [age_results['change'][2], gender_results['change'][2], occupation_results['change'][2], source_results['change'][2]],
    'Change Cramer\'s V': [round(age_results['change'][3], 3), round(gender_results['change'][3], 3), round(occupation_results['change'][3], 3), round(source_results['change'][3], 3)],
})

# プロットの作成
fig, axes = plt.subplots(3, 1, figsize=(15, 15))  # プロットのサイズを指定

# Empathizedの結果を表示
axes[0].axis('tight')
axes[0].axis('off')
tbl_empathized = table(axes[0], results_emp, loc='center', cellLoc='center', colWidths=[0.1]*len(results_emp.columns))
tbl_empathized.auto_set_font_size(False)
tbl_empathized.set_fontsize(12)
tbl_empathized.scale(1.2, 1.2)  # テーブルのサイズを調整
axes[0].set_title('Empathized Chi-Square Test Results Summary', fontsize=16)

# Anxietyの結果を表示
axes[1].axis('tight')
axes[1].axis('off')
tbl_anxiety = table(axes[1], results_anx, loc='center', cellLoc='center', colWidths=[0.1]*len(results_anx.columns))
tbl_anxiety.auto_set_font_size(False)
tbl_anxiety.set_fontsize(12)
tbl_anxiety.scale(1.2, 1.2)  # テーブルのサイズを調整
axes[1].set_title('Anxiety Chi-Square Test Results Summary', fontsize=16)

# Changeの結果を表示
axes[2].axis('tight')
axes[2].axis('off')
tbl_change = table(axes[2], results_cha, loc='center', cellLoc='center', colWidths=[0.1]*len(results_cha.columns))
tbl_change.auto_set_font_size(False)
tbl_change.set_fontsize(12)
tbl_change.scale(1.2, 1.2)  # テーブルのサイズを調整
axes[2].set_title('Change Chi-Square Test Results Summary', fontsize=16)

# プロットの表示
plt.tight_layout()
plt.show()

##############

empathized_ct = pd.crosstab(data['age'], data['empathized'])

plt.figure(figsize=(10, 6))
sns.heatmap(empathized_ct, annot=True, fmt="d", cmap="YlGnBu")
plt.title('Contingency Table: Age vs. Empathized')
plt.ylabel('Age')
plt.xlabel('Empathized')
plt.savefig('contingency_table_age_empathized.png')
plt.show()

###############

# プロットの作成
fig, ax = plt.subplots(3, 1, figsize=(12, 18))

# Empathizedのカイ二乗値のプロット
ax[0].bar(results_df['Variable'], results_df['Empathized Chi2'], color=['red' if sig else 'blue' for sig in results_df['Empathized Significant']])
ax[0].set_title('for Empathized')
ax[0].set_ylabel('Chi-Square Value')

# Anxietyのカイ二乗値のプロット
ax[1].bar(results_df['Variable'], results_df['Anxiety Chi2'], color=['red' if sig else 'blue' for sig in results_df['Anxiety Significant']])
ax[1].set_title('for Anxiety')
ax[1].set_ylabel('Chi-Square Value')

# Changeのカイ二乗値のプロット
ax[2].bar(results_df['Variable'], results_df['Change Chi2'], color=['red' if sig else 'blue' for sig in results_df['Change Significant']])
ax[2].set_title('for Change')
ax[2].set_ylabel('Chi-Square Value')

# プロットの表示
plt.tight_layout()
plt.show()