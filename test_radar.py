import requests
import json

# 测试数据
test_data = [
    {'指标': 'W1C', '0d': 2.5, '1W': 2.0, '1NO': 2.8},
    {'指标': 'W3S', '0d': 3.0, '1W': 3.5, '1NO': 2.5},
    {'指标': 'W2W', '0d': 1.5, '1W': 2.0, '1NO': 1.8}
]

payload = {
    'chart_type': 'radar',
    'data': test_data,
    'axes_col': '指标',
    'series_cols': ['0d', '1W', '1NO'],
    'colors': ['#000000', '#FF0000', '#00FF00'],
    'line_styles': ['-', '-', '-'],
    'line_widths': [2, 2, 2],
    'marker_styles': ['o', 's', '^'],
    'transpose': False,
    'title': '',
    'font_sizes': {
        'title': 14,
        'axis_label': 13,
        'tick': 11,
        'legend': 10
    }
}

print("Sending request to http://localhost:5000/api/plot")
print("Payload:", json.dumps(payload, indent=2, ensure_ascii=False))

try:
    response = requests.post('http://localhost:5000/api/plot', json=payload)
    print(f"\nStatus Code: {response.status_code}")
    print(f"Response: {response.text[:500]}")

    if response.status_code == 200:
        result = response.json()
        if 'image' in result:
            print("\n✓ Success! Image generated")
        else:
            print("\n✗ Error: No image in response")
    else:
        print(f"\n✗ Error: {response.text}")
except Exception as e:
    print(f"\n✗ Exception: {e}")
