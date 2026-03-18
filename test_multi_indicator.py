"""
生成多指标测试数据，验证自动识别功能
"""

import pandas as pd
import numpy as np

def generate_test_data():
    """生成包含多个指标的测试数据"""
    np.random.seed(42)

    # 创建数据 - 包含明显的样品列和多个指标列
    data = []
    groups = ['0K', '0NO', '1K', '1NO', '3K', '3NO']

    for group in groups:
        # 每个组3个平行样
        for replicate in range(1, 4):
            data.append({
                '样品名称': f'{group}{replicate}',  # 0K1, 0K2, 0K3, ...
                '蛋白质含量': round(np.random.normal(100 + groups.index(group) * 5, 10), 2),
                '酶活性_U_g': round(np.random.normal(50 + groups.index(group) * 3, 5), 2),
                '水分含量_%': round(np.random.normal(80 - groups.index(group) * 2, 3), 2),
                'pH值': round(np.random.normal(6.5 + groups.index(group) * 0.1, 0.2), 2),
            })

    df = pd.DataFrame(data)

    # 保存为Excel
    filename = 'test_multi_indicator.xlsx'
    df.to_excel(filename, index=False, engine='openpyxl')

    print(f"测试数据已生成: {filename}")
    print(f"\n数据预览:")
    print(df.head(10))
    print(f"\n数据形状: {df.shape}")
    print(f"\n列信息:")
    for col in df.columns:
        print(f"  - {col}: {df[col].dtype}, {df[col].nunique()} 唯一值")

    return df

if __name__ == '__main__':
    print("=" * 60)
    print("生成多指标测试数据")
    print("=" * 60)
    generate_test_data()
    print("\n" + "=" * 60)
    print("测试数据生成完成！")
    print("请在网站中上传 test_multi_indicator.xlsx 进行测试")
    print("=" * 60)
