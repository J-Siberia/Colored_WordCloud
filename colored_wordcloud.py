import pandas as pd
import matplotlib.pyplot as plt
from wordcloud import WordCloud

CHAP_COLOR = (230, 0, 18)   # 赤色
CALU_COLOR = (235, 97, 0)   # 淡い赤色
SECT_COLOR = (243, 152, 0)  # 橙色
BOLD_COLOR = (252, 200, 0)  # 黄色

# 単語と色のマッピングを定義する
color_mapping = {
    # 章タイトル
    '色': CHAP_COLOR,
    # 節タイトル
    '目の構造と色覚': CALU_COLOR,
    '色の見え方の性質': CALU_COLOR,
    '色空間': CALU_COLOR,
    # 項タイトル
    '三原色と混色': SECT_COLOR,
    '目の構造': SECT_COLOR,
    '三色説': SECT_COLOR,
    '反対色説': SECT_COLOR,
    '段階説': SECT_COLOR,
    '色覚異常': SECT_COLOR,
    '色の三属性': SECT_COLOR,
    'コンテキストの影響': SECT_COLOR,
    '大きさの影響': SECT_COLOR,
    'コントラストの影響': SECT_COLOR,
    '色の与える印象': SECT_COLOR,
    '色覚異常者への配慮': SECT_COLOR,
    # 太字
    '加法混色': BOLD_COLOR,
    '減法混色': BOLD_COLOR,
    '桿体視細胞': BOLD_COLOR,
    '錐体視細胞': BOLD_COLOR,
    '色相': BOLD_COLOR,
    '彩度': BOLD_COLOR,
    '明度': BOLD_COLOR,
    'RGB色空間': BOLD_COLOR,
    'HSV色空間': BOLD_COLOR,
    'HSB色空間': BOLD_COLOR,
    'CIE1976L*a*b* 均等色空間': BOLD_COLOR,
    'CIE1976L*u*v* 均等色空間': BOLD_COLOR,
    'CIEDE2000 色差式': BOLD_COLOR
}

# Excelファイルのパスを指定してデータを読み込む
df = pd.read_excel('PassToData')

# フォントの指定
WORDCLOUD_FONT_PATH = "PassToFONT"

# 複数の列のデータを結合する
text = ' '.join(df[1].str.strip()) + ' '.join(df[2].str.strip()) + ' '.join(df[3].str.strip())

# WordCloudオブジェクトを作成
wordcloud = WordCloud(prefer_horizontal=1, background_color='white', width=2560, height=1440, font_path=WORDCLOUD_FONT_PATH).generate(text).recolor(color_func=lambda word, font_size, position, orientation, random_state=None, **kwargs: color_mapping.get(word, (0, 0, 0)))

# WordCloudを表示する
plt.figure(figsize=(20.0, 6.8))
plt.imshow(wordcloud, interpolation='bilinear')
plt.axis('off')
plt.show()