import matplotlib.pyplot as plt
import numpy as np
import os

# 設定中文微軟正黑體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'PingFang TC', 'SimHei', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

def generate_blank_radar_chart():
    categories = ['生態自然感', '便利繁華度', '生存壓力感', '人際親密感', '悠活慢調度']
    N = len(categories)

    # 角度（正上方 90 度開始，順時針）
    angles = np.linspace(np.pi / 2, np.pi / 2 - 2 * np.pi, N, endpoint=False).tolist()
    angles += angles[:1]  # 閉合

    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True), dpi=300)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    # 設定刻度 0 到 10
    ax.set_ylim(0, 10)
    ax.set_yticks([2, 4, 6, 8, 10])
    ax.set_yticklabels(['2', '4', '6', '8', '10'], color='#64748b', size=9, weight='bold')
    ax.set_rlabel_position(45)

    # 設置網格樣式
    ax.grid(color='#cbd5e1', linestyle='--', linewidth=1.0)
    ax.spines['polar'].set_color('#94a3b8')
    ax.spines['polar'].set_linewidth(1.5)

    # 設置各軸角度與標籤
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, size=11.5, weight='bold', color='#1e293b')

    # 調整標籤距離
    ax.tick_params(pad=14)

    # 標記中心點
    ax.plot(0, 0, marker='o', markersize=4, color='#64748b')

    plt.tight_layout()
    output_path = r"c:\2026_透析臺北──文本比較\radar_grid_blank.png"
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print("Blank radar chart generated successfully at:", output_path)

if __name__ == '__main__':
    generate_blank_radar_chart()
