"""
生成示例测试数据
用于测试ANOVA分析功能
"""

import pandas as pd
import numpy as np

def generate_anova_sample_data(filename='sample_data.xlsx'):
    """
    生成用于ANOVA分析的示例数据
    包含三组样品，每组有显著差异的均值
    """
    np.random.seed(42)  # 确保可重复性

    # 定义三组数据，具有不同的均值（模拟真实实验数据）
    groups = {
        '对照组': np.random.normal(50, 5, 10),      # 均值50，标准差5
        '处理组A': np.random.normal(55, 5, 10),     # 均值55，与对照组有差异
        '处理组B': np.random.normal(62, 5, 10),     # 均值62，差异更大
    }

    # 创建DataFrame
    data = []
    for group_name, values in groups.items():
        for value in values:
            data.append({
                '样品名称': group_name,
                '指标结果': round(value, 2),
                '重复次数': len([v for v in values])  # 每组重复次数
            })

    df = pd.DataFrame(data)

    # 保存为Excel
    df.to_excel(filename, index=False, engine='openpyxl')
    print(f"示例数据已生成: {filename}")
    print(f"\n数据预览:")
    print(df.head(15))
    print(f"\n描述性统计:")
    print(df.groupby('样品名称')['指标结果'].describe())
    print(f"\n提示: 这组数据中三组之间的差异应该是显著的")
    return df

def generate_multi_indicator_data(filename='multi_indicator_sample.xlsx'):
    """
    生成包含多个指标的示例数据
    """
    np.random.seed(123)

    groups = ['对照组', '低剂量组', '中剂量组', '高剂量组']

    data = []
    for group in groups:
        for i in range(8):  # 每组8个重复
            data.append({
                '样品名称': group,
                '蛋白质含量': round(np.random.normal(
                    {'对照组': 100, '低剂量组': 110, '中剂量组': 125, '高剂量组': 140}[group], 10), 2),
                '酶活性': round(np.random.normal(
                    {'对照组': 50, '低剂量组': 58, '中剂量组': 70, '高剂量组': 85}[group], 8), 2),
                '细胞活力': round(np.random.normal(
                    {'对照组': 90, '低剂量组': 88, '中剂量组': 82, '高剂量组': 75}[group], 5), 2),
            })

    df = pd.DataFrame(data)
    df.to_excel(filename, index=False, engine='openpyxl')
    print(f"\n多指标示例数据已生成: {filename}")
    print(f"\n数据预览:")
    print(df.head(12))
    return df

if __name__ == '__main__':
    print("=" * 60)
    print("生成ANOVA分析示例数据")
    print("=" * 60)

    # 生成单指标示例数据
    generate_anova_sample_data('sample_data.xlsx')

    # 生成多指标示例数据
    generate_multi_indicator_data('multi_indicator_sample.xlsx')

    print("\n" + "=" * 60)
    print("数据生成完成！")
    print("=" * 60)
    print("\n您可以使用以下文件测试网站功能:")
    print("1. sample_data.xlsx - 单指标数据")
    print("2. multi_indicator_sample.xlsx - 多指标数据")
    print("\n启动服务器后，在 http://localhost:5000 上传这些文件进行测试")
