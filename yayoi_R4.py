import pandas as pd
import numpy as np
import jaconv
import streamlit as st
from io import BytesIO

st.markdown("""
    弥生の仕訳データは、アップロードする前に**「日付」**行を**「短い日付形式」**に変換してからアップロードしてね。
    """)


# CSVファイルのアップロード
uploaded_file1 = st.file_uploader("弥生の仕訳データのcsvをアップロード", type=['csv'])
uploaded_file2 = st.file_uploader("財務のMKAMOKUのcsvをアップロード", type=['csv'])
uploaded_file3 = st.file_uploader("財務のSKAMOKUのcsvをアップロード", type=['csv'])

if st.checkbox('OK') and uploaded_file1 is not None and uploaded_file2 is not None and uploaded_file3 is not None:
    try:
        df = pd.read_csv(uploaded_file1 , encoding = 'cp932')
        df_master = pd.read_csv(uploaded_file2 , encoding = 'cp932')
        df_sub = pd.read_csv(uploaded_file3 , encoding = 'cp932')

        new_columns = [
            '日付', '借方科目コード', '借方科目名', '借方補助コード', '借方補助科目名', '借方部門', '借方税区分','借方税計算区分',
            '借方金額', '貸方科目コード', '貸方科目名', '貸方補助コード', '貸方補助科目名', '貸方部門', 
            '貸方税区分','貸方税計算区分', '貸方金額', '摘要'
        ]

        # 条件に基づいてデータをフィルタリング
        mask = df['[表題行]'] == '[明細行]'
        df_filtered = df.loc[mask].copy()

        # 新しいDataFrameを作成し、列名リストに基づいて列を再配置
        df_provi = pd.DataFrame(index=df_filtered.index, columns=new_columns)

        # 必要なデータを各列に転記
        mapping = {
            '日付': '日付',
            '借方科目名': '借方勘定科目',
            '借方補助科目名': '借方補助科目',
            '借方部門': '借方部門',
            '借方税区分': '借方税区分',
            '借方税計算区分': '借方税計算区分',
            '借方金額': '借方金額',
            '貸方科目名': '貸方勘定科目',
            '貸方補助科目名': '貸方補助科目',
            '貸方部門': '貸方部門',
            '貸方税区分': '貸方税区分',
            '貸方税計算区分': '貸方税計算区分',
            '貸方金額': '貸方金額',
            '摘要': '摘要'
        }

        for key, value in mapping.items():
            df_provi[key] = df_filtered[value]

        # 未定義の列にNaNを設定
        df_provi.fillna(np.nan, inplace=True)

        # インデックスをリセット
        df_provi.reset_index(drop=True, inplace=True)

        df_provi['借方科目名'] = df_provi['借方科目名'].fillna('諸口')
        df_provi['貸方科目名'] = df_provi['貸方科目名'].fillna('諸口')
        df_provi.loc[df_provi['借方科目名'] == '諸口', '借方科目コード'] = 999
        df_provi.loc[df_provi['貸方科目名'] == '諸口', '貸方科目コード'] = 999
        df_provi['貸方金額'] = df_provi['貸方金額'].fillna(df_provi['借方金額'])
        df_provi['借方金額'] = df_provi['借方金額'].fillna(df_provi['貸方金額'])

        # df_master からマッピング用の辞書を作成
        master_mapping = df_master.set_index('勘定科目名')['勘定科目コード'].to_dict()
        # df_provi の '借方科目名' 列にマッピングを適用
        df_provi['借方科目コード'] = df_provi['借方科目名'].map(master_mapping)

        # df_master からマッピング用の辞書を作成
        master_mapping = df_master.set_index('勘定科目名')['勘定科目コード'].to_dict()
        # df_provi の '借方科目名' 列にマッピングを適用
        df_provi['貸方科目コード'] = df_provi['貸方科目名'].map(master_mapping)

        # df_master の '勘定科目名' を全角から半角に変換（またはその逆）
        df_master['勘定科目名'] = df_master['勘定科目名'].apply(lambda x: jaconv.zen2han(x, kana=False, ascii=True, digit=True))

        # df_provi の '借方科目名' と '貸方科目名' も同様に変換
        df_provi['借方科目名'] = df_provi['借方科目名'].apply(lambda x: jaconv.zen2han(x, kana=False, ascii=True, digit=True))
        df_provi['貸方科目名'] = df_provi['貸方科目名'].apply(lambda x: jaconv.zen2han(x, kana=False, ascii=True, digit=True))

        # マッピング用の辞書を作成
        master_mapping = df_master.set_index('勘定科目名')['勘定科目コード'].to_dict()

        # マッピングを適用
        df_provi['借方科目コード'] = df_provi['借方科目名'].map(master_mapping)
        df_provi['貸方科目コード'] = df_provi['貸方科目名'].map(master_mapping)

        # ミスマッチの行を識別
        mismatched_rows = df_provi['借方金額'] != df_provi['貸方金額']

        # 貸方科目名が特定の条件を満たす場合
        conditions_credit = df_provi['貸方科目名'].isin(['売上高', '固定資産売却益', '有価証券売却益']) & mismatched_rows
        df_provi.loc[conditions_credit, '借方金額'] = df_provi.loc[conditions_credit, '貸方金額']

        # 借方科目名が特定の条件を満たす場合
        conditions_debit = df_provi['借方科目名'].isin(['仕入高', '固定資産売却損']) & mismatched_rows
        df_provi.loc[conditions_debit, '貸方金額'] = df_provi.loc[conditions_debit, '借方金額']

        # 上記いずれの条件にも当てはまらない場合
        conditions_neither = ~(conditions_credit | conditions_debit) & mismatched_rows
        df_provi.loc[conditions_neither, '貸方金額'] = df_provi.loc[conditions_neither, '借方金額']
        df_provi.loc[conditions_neither, '摘要'] = df_provi.loc[conditions_neither, '摘要'] + ' ●金額確認必須'

        df_provi.loc[(df_provi['貸方科目名'] == '売上高') & (df_provi['貸方補助科目名'] == '賃貸収入'), '貸方科目コード'] = 811
        df_provi.loc[(df_provi['貸方科目名'] == '売上高') & (df_provi['貸方補助科目名'] == '製品売上高'), '貸方科目コード'] = 810
        df_provi.loc[(df_provi['借方科目名'] == '売上高') & (df_provi['借方補助科目名'] == '賃貸収入'), '借方科目コード'] = 811
        df_provi.loc[(df_provi['借方科目名'] == '売上高') & (df_provi['借方補助科目名'] == '製品売上高'), '借方科目コード'] = 810

        # 貸方補助科目名に対する借方補助科目コードのマッピング
        codes = {
            '横浜ゴム㈱': 1, '横浜MBジャパン㈱': 1, '横浜MBジャパン㈱茨城': 1,
            'ハイテック': 2, '大阪高圧ホース株式会社': 2, '北興商事': 2, '共栄産業株式会社': 2, '有限会社郡山高圧': 2, '山清工業': 2, '㈱サンテム': 2, 'ツツミ工業株式会社': 2,
            'タノ製作所美里工業': 3, 'ライフエレクトロ': 3, '市光工業': 3,
            '㈱宮入バルブ製作所': 4,
            '小林ｽﾌﾟﾘﾝｸﾞ': 5,
            'ｶﾝﾂｰﾙ': 6,
            'ダンレイ': 7,
            'ﾕﾀｶ': 8,
            '㈱テクノフレックス': 9, '阿部製作所': 9, '加藤製作所': 9, '株式会社富藤製作所': 9,
            'ニューマシン': 10,
            'ハマイ　大多喜': 11,
            '桂精機': 12,
            '佐々木発條': 13,
            '花岡車輌㈱': 15,
            '三笠': 16,
            '㈱宮入製作所': 17,
            'アルバック機工': 18
        }

        # 条件に一致する行をフィルタリング
        condition = (df_provi['借方科目名'] == '売上高') & (df_provi['借方補助科目名'] == '製品売上高')

        # 条件に一致した行に対してマッピングを適用
        df_provi.loc[condition, '借方補助コード'] = df_provi.loc[condition, '貸方補助科目名'].map(codes).fillna(19)

        # 条件に一致する行をフィルタリング
        condition = (df_provi['貸方科目名'] == '売上高') & (df_provi['貸方補助科目名'] == '製品売上高')

        # 条件に一致した行に対してマッピングを適用
        df_provi.loc[condition, '貸方補助コード'] = df_provi.loc[condition, '借方補助科目名'].map(codes).fillna(19)

        # 貸方補助科目名に対する借方補助科目コードのマッピング
        codes_purchase = {
            '日鉄物産ワイヤ＆': 1,
            '平野綱線': 2,
            'ステラ': 3,
            '株式会社小林スプリング': 4,
            '大木発条': 5,
            'ヒタチスプリン': 6,
            '東邦発条': 7,
            '佐野鍍金': 8,
            '㈱明光電化工業': 9,
            '大阪ばね工業㈱': 10,
            '日産ｽﾌﾟﾘﾝｸﾞ': 11,
            '宮精機': 12,
            'ライトスプリン': 13,
            '東亜製砥工業': 14,
            '㈱巧工業': 15,
            '長泉ﾊﾟｰｶｰ': 16,
        }

        # 条件に一致する行をフィルタリング
        condition = (df_provi['借方科目名'] == '材料仕入高')

        # 条件に一致した行に対してマッピングを適用
        df_provi.loc[condition, '借方補助コード'] = df_provi.loc[condition, '貸方補助科目名'].map(codes_purchase).fillna(99)

        # 条件に一致する行をフィルタリング
        condition = (df_provi['借方科目名'] == 'C外注加工費')

        # 条件に一致した行に対してマッピングを適用
        df_provi.loc[condition, '借方補助コード'] = df_provi.loc[condition, '貸方補助科目名'].map(codes_purchase).fillna(99)

        # 条件に一致する行をフィルタリング
        condition = (df_provi['借方科目名'] == 'C消耗品費')

        # 条件に一致した行に対してマッピングを適用
        df_provi.loc[condition, '借方補助コード'] = df_provi.loc[condition, '貸方補助科目名'].map(codes_purchase).fillna(99)

        # 条件に一致する行をフィルタリング
        condition = (df_provi['貸方科目名'] == '材料仕入高')

        # 条件に一致した行に対してマッピングを適用
        df_provi.loc[condition, '貸方補助コード'] = df_provi.loc[condition, '借方補助科目名'].map(codes_purchase).fillna(99)

        # 条件に一致する行をフィルタリング
        condition = (df_provi['貸方科目名'] == 'C外注加工費')

        # 条件に一致した行に対してマッピングを適用
        df_provi.loc[condition, '貸方補助コード'] = df_provi.loc[condition, '借方補助科目名'].map(codes_purchase).fillna(99)

        # 条件に一致する行をフィルタリング
        condition = (df_provi['貸方科目名'] == 'C消耗品費')

        # 条件に一致した行に対してマッピングを適用
        df_provi.loc[condition, '貸方補助コード'] = df_provi.loc[condition, '借方補助科目名'].map(codes_purchase).fillna(99)

        output_columns =["月種別","種類","形式","作成方法","付箋","伝票日付","伝票番号","伝票摘要","枝番","借方部門","借方部門名","借方科目","借方科目名","借方補助","借方補助科目名","借方金額","借方消費税コード","借方消費税業種","借方消費税税率","借方資金区分","借方任意項目１","借方任意項目２","借方インボイス情報","貸方部門","貸方部門名","貸方科目","貸方科目名","貸方補助","貸方補助科目名","貸方金額","貸方消費税コード","貸方消費税業種","貸方消費税税率","貸方資金区分","貸方任意項目１","貸方任意項目２","貸方インボイス情報","摘要","期日","証番号","入力マシン","入力ユーザ","入力アプリ","入力会社","入力日付"
        ]

        output_df = pd.DataFrame(columns = output_columns, index = df_provi.index)

        df_provi['借方科目コード'] = df_provi['借方科目コード'].fillna(0).astype(int)
        df_provi['貸方科目コード'] = df_provi['貸方科目コード'].fillna(0).astype(int)
        df_provi['借方補助コード'] = df_provi['借方補助コード'].fillna(0).astype(int)
        df_provi['貸方補助コード'] = df_provi['貸方補助コード'].fillna(0).astype(int)
        df_provi['借方金額'] = df_provi['借方金額'].fillna(0).astype(int)
        df_provi['貸方金額'] = df_provi['貸方金額'].fillna(0).astype(int)

        # df_provi から output_df へ特定の列のデータを転記
        output_df['伝票日付'] = df_provi['日付']
        output_df['借方科目'] = df_provi['借方科目コード']
        output_df['借方科目名'] = df_provi['借方科目名']
        output_df['借方補助'] = df_provi['借方補助コード']
        output_df['借方金額'] = df_provi['借方金額']
        output_df['貸方科目'] = df_provi['貸方科目コード']
        output_df['貸方科目名'] = df_provi['貸方科目名']
        output_df['貸方補助'] = df_provi['貸方補助コード']
        output_df['貸方金額'] = df_provi['貸方金額']
        output_df['摘要'] = df_provi['摘要']

        # DataFrameをCSV形式のバイト列に変換
        to_download = BytesIO()
        output_df.to_csv(to_download, encoding='cp932', index=False)
        to_download.seek(0)

        # ダウンロードボタンを作成
        st.snow()
        st.download_button(
            label="ダウンロード",
            data=to_download,
            file_name='output.csv',
            mime='text/csv'
        )


    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
