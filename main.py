from __future__ import annotations
from fileinput import filename
from mimetypes import init
from traceback import print_tb
from turtle import back
import pandas as pd
import glob
import shutil
import os

import logging

# ログの設定
logging.basicConfig(filename='logfile.log', level=logging.DEBUG)

# GUIツールを読み込む
import tkinter as tk
import tkinter.filedialog

# 日付関連のライブラリを読み込む
import datetime
import locale

# PDF関連のライブラリを読み込む
from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
from pytest import skip

from reportlab.pdfgen import canvas
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import mm
from reportlab.lib.colors import red, blue, black, orange, green

from reportlab.platypus import BaseDocTemplate, PageTemplate
from reportlab.platypus import Paragraph, PageBreak, FrameBreak
from reportlab.platypus.flowables import Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import A4, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase import cidfonts

from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.high_level import extract_text
import PyPDF2
from PyPDF2 import PdfFileReader
# ここまでPDF操作関連

import re
import math

try:

    def main():
        pdf_path = './input_test2.pdf'
        pdf_path = pick_data_from_gui(
            "Shopifyからエクスポートした明細書(PDFを選択してください", file_type=".pdf")

        shopify_csv_path = "./input_test2.csv"
        shopify_csv_path = pick_data_from_gui(
            "Shopifyからエクスポートした注文情報(CSV)を選択してください", file_type=".csv")
        shopify_df = pd.read_csv(shopify_csv_path)

        instraction_csv_path = "梱包指示.csv"
        instraction_csv_path = pick_data_from_gui(
            "Spreadsheetからエクスポートした梱包指示・出荷前の変更(CSV)を選択してください", file_type=".csv")
        instraction_df = pd.read_csv(instraction_csv_path)

        editer = PdfEditer(pdf_path)

        # 最初と最後の注文IDを取得する
        editer.get_first_order_id_and_last_order_id(pdf_path=pdf_path)

        # 現在時刻JSTを取得する
        dt_now_jst_aware = editer.get_formated_date()

        # PDFの名前を最初の注文ID＋最後の注文ID＋日付にする
        pdf_filename = PdfEditer.first_order_id_in_process + "-" + \
            PdfEditer.last_order_id_in_process + "_" + dt_now_jst_aware + ".pdf"

        # 入力されたpdfの枚数を取得する
        number_of_pages = editer.get_pdf_info_from_path(pdf_path)

        # PDFを0ページ目から読み込む
        for current_page in range(number_of_pages):
            editer.set_output_pdf_name(
                "./tmp-pdf/" + str(current_page).zfill(4) + ".pdf")
            # print(f'ゼロ埋めした注文番号：{str(current_page).zfill(4)}')
            editer.cc.setFont("HeiseiKakuGo-W5", 20)

            # 記入するpdfページを読み込む
            editer.get_pdf_object_from_path(pdf_path, current_page)
            text = extract_text(pdf_path, page_numbers=[current_page])
            # 読み込んだPDFのページから注文番号(#○○○○○)を探す
            purchase_id_re_object = re.search(r'#\d{5}', text)

            # 現在のページを最後のページとして初期化する
            last_paage = current_page

            # PDFから注文番号を探すことができたら、各種csvから記載内容を探す。探せなければスキップ
            if purchase_id_re_object is not None:
                # 注文がどこまで続いているか調べる
                count_continue = 0
                for i in range(current_page, number_of_pages):
                    # print(f'注文がどこまで続いているかループ：{i}')
                    next_page_text = extract_text(pdf_path, page_numbers=[i])
                    next_purchase_id_re_object = re.search(
                        r'#\d{5}', next_page_text)
                    if next_purchase_id_re_object is None:
                        # 注文番号が無いページ=続いているページをカウント
                        count_continue += 1
                    else:
                        break
                print(f'1枚目以外のページ数:{count_continue}')
                last_paage = current_page + count_continue

                # re.Matchオブジェクトから注文番号を取り出す
                purchase_id = purchase_id_re_object.group()

                # 注文番号を頼りに各種CSVから記載事項を探す
                # shopifyからエクスポートしたcsvから該当する注文情報を探し、取得する
                one_subtotal, _ = editer.get_value_from_df_and_kay(check_df=shopify_df, search_key=purchase_id, key_column="Name", value_column="Subtotal")
                one_discount, _ = editer.get_value_from_df_and_kay(check_df=shopify_df, search_key=purchase_id, key_column="Name", value_column="Discount Amount")

                # 値段が20,000円以上の場合"おまけ有"とする
                #if one_subtotal + one_discount > 19999:
                    #editer.highlight_text(80*mm, 270*mm, "おまけ有", black, orange, 20)
                
                # 〇万円以上購入で各おまけあり仕様へ
                #if one_subtotal + one_discount > 49999:
                    #editer.highlight_text(80*mm, 270*mm, "5万円おまけ", black, orange, 20)
                if one_subtotal + one_discount > 29999:
                    editer.highlight_text(80*mm, 270*mm, "3万円おまけ", black, orange, 20)
                elif one_subtotal + one_discount > 19999:
                    editer.highlight_text(80*mm, 270*mm, "2万円おまけ", black, orange, 20)
                #elif  one_subtotal + one_discount > 9999:
                    #editer.highlight_text(80*mm, 300*mm, "1万円おまけ", black, orange, 20)

                # 注文番号を頼りにクーポンコードを受け取る
                sasutainable_coupon, _ = editer.get_value_from_df_and_kay(
                    check_df=shopify_df, search_key=purchase_id, key_column="Name", value_column="Discount Code")
                # クーポンコードが"ARAS30"または"ARASQA30"であれば簡易包装とする
                if sasutainable_coupon == "ARAS30":
                    editer.highlight_text(80*mm, 285*mm, "簡易包装", black, orange, 20)
                #elif sasutainable_coupon == "ARASQA30":
                    #editer.highlight_text(80*mm, 285*mm, "簡易包装", black, orange, 20)
                

                # 注文番号を頼りに決済ステータスを受け取る
                financial_status, _ = editer.get_value_from_df_and_kay(
                    check_df=shopify_df, search_key=purchase_id, key_column="Name", value_column="Financial Status")
                # 決済ステータスが pendingであれば代引きとする
                if financial_status == "pending":
                    # editer.highlight_text(80*mm, 285*mm, "備有", black, orange, 20)
                    editer.cc.setFont("HeiseiKakuGo-W5", 20)
                    editer.cc.setFillColor(red)
                    one_total, _ = editer.get_value_from_df_and_kay(check_df=shopify_df, search_key=purchase_id, key_column="Name", value_column="Total")
                    separated_total = editer.culc_collect_fee(one_total)
                    editer.cc.drawString(
                        120*mm, 284*mm, f'代引き: {separated_total}円')

                # 配送指示を注文番号を頼りに受け取る
                shipping_inst, _ = editer.get_value_from_df_and_kay(
                    check_df=shopify_df, search_key=purchase_id, key_column="Name", value_column="Note Attributes")
                if shipping_inst != '':
                    editer.cc.setFont("HeiseiKakuGo-W5", 11)
                    editer.cc.setFillColor(red)
                    editer.auto_indent(str(shipping_inst), 145*mm, 265*mm)

            # 2枚目以降のページである場合、記載対象であれば記載する
            # 記載事項
            # 最後のページ：出荷指示
            # 全ページ:ページ番号
            if current_page == last_paage:
                # 注文番号を頼りに出荷指示を確認する
                note_attributes, _ = editer.get_value_from_df_and_kay(
                    check_df=instraction_df, search_key=purchase_id, key_column="注文番号", value_column="指示", add_head_hash=True)
                print(f'注文番号：{purchase_id}, 指示：{note_attributes}')
                if note_attributes is str and note_attributes is not None:
                    editer.highlight_text(
                        80*mm, 285*mm, "備有", black, orange, 20)
                    editer.cc.setFont("HeiseiKakuGo-W5", 20)
                    editer.cc.setFillColor(red)

                    # 最後のページに注文情報が入りきらないようなら新しくpdfを作成する
                    item_count = editer.get_value_from_df_and_kay(
                        check_df=shopify_df, search_key=purchase_id, key_column="Name", value_column="Subtotal")
                    print(f'Item数：{item_count}')
                    if item_count > 7 and count_continue != 1:
                        # ページ情報を記載
                        editer.highlight_text(
                            10*mm, 285*mm, f'最後 注有：{last_paage - count_continue - current_page + 1}/{count_continue + 2}', black, orange, 20)
                        editer.cc.showPage
                        editer.cc.save()
                        editer.set_output_pdf_name(
                            './tmp-pdf/' + str(current_page) + '_1.pdf')
                        pp = pagexobj(editer.black_pdf[0])
                        editer.cc.doForm(makerl(editer.cc, pp))
                        editer.cc.setFont("HeiseiKakuGo-W5", 20)
                        editer.cc.setFillColor(black)
                        # ページ情報を記載
                        editer.highlight_text(
                            10*mm, 285*mm, f'最後 注有：{last_paage - count_continue - current_page + 2}/{count_continue + 2}, 注文ID：{purchase_id}', black, orange, 20)

                    # 備考を記入する
                    editer.cc.setFillColor(black)
                    editer.auto_indent(note_attributes)

            # 現在のページが最後のページでないとき、左上にページ番号を付す
            elif current_page < last_paage:
                # 注文番号を頼りに出荷指示を確認する
                note_attributes, _ = editer.get_value_from_df_and_kay(
                    check_df=instraction_df, search_key=purchase_id, key_column="注文番号", value_column="指示", add_head_hash=True)
                print(f'注文番号：{purchase_id}, 指示：{note_attributes}')
                if note_attributes is not None:
                    editer.highlight_text(
                        80*mm, 285*mm, "備有", black, orange, 20)
                    editer.cc.setFont("HeiseiKakuGo-W5", 20)
                    editer.cc.setFillColor(red)
                    # ページ情報を記載
                    editer.highlight_text(
                        10*mm, 285*mm, f'最後 注有：{last_paage - count_continue - current_page + 1}/{count_continue + 1}', black, orange, 20)

            editer.cc.showPage
            editer.cc.save()

        PdfEditer.merge_pdf_in_dir(r'C:\Users\s.nozeki\Desktop\Automated-shipping-process-for-Shopify-orders\tmp-pdf', pdf_filename)

        # 一時的に作成されたpdfを削除
        #target_dir = r'C:\Users\s.nozeki\Desktop\Automated-shipping-process-for-Shopify-orders\tmp-pdf'
        #shutil.rmtree (target_dir)
        #os.mkdir (target_dir)


    class OnePdfPage:
        number_of_page_at_this_order = 0
        first_page_of_this_order = 0
        text = []
        order_id = ""


    class PdfEditer:
        """_summary_
        """
        purchase_df = pd.DataFrame()
        note_instructions_df = pd.DataFrame()
        first_order_id_in_process = ""
        last_order_id_in_process = ""
        pointer_position_x, pointer_pointer_position_y = 0, 0
        # PDF新規作成
        cc = canvas.Canvas("output.pdf")
        cc.setFillColor(red)
        cc.setStrokeColor(red)

        # フォントの設定
        font_name = "HeiseiKakuGo-W5"
        pdfmetrics.registerFont(UnicodeCIDFont(font_name))
        cc.setFont(font_name, 8)

        def __init__(self, pdf_path):
            # アノテーション対象ページ読み込み
            self.page = PdfReader(pdf_path, decompress=False).pages
            self.black_pdf = PdfReader(
                './blanksheet-a4-portrait.pdf', decompress=False).pages

        def get_formated_date(self):
            """日付を取得して文字列にして返す
            """

            # ロケールを日本語に設定
            locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')
            # 現在時刻JSTを取得する
            now_date = datetime.datetime.now(
                datetime.timezone(datetime.timedelta(hours=9)))

            # 曜日をフォーマット
            days = ['月', '火', '水', '木', '金', '土', '日']
            day = now_date.weekday()
            formated_day = f'({days[day]})'

            # 曜日以外をフォーマット
            dt_now_jst_aware = now_date.strftime('%Y{0}%m{1}%d{2}').format(*'年月日')

            return(dt_now_jst_aware + formated_day)

        # 合計金額が1万円を超えている場合、各金額帯で代引き手数料を加算
        def culc_collect_fee(self, total):
            if total < 10000:
                added_collect_fee = total + 330
            elif total < 30000:
                added_collect_fee = total + 440
            elif total < 100000:
                added_collect_fee = total + 660
            elif total < 300000:
                added_collect_fee = total + 1100
            else:
                return("代引き不能金額 >30万円")

            separated_total = "{:,}".format(int(added_collect_fee))
            return(separated_total)

        def merge_pdf_in_dir(dir_path, dst_path):
            '''
            作成された一時PDFを連結
            '''
            l = glob.glob(os.path.join(dir_path, '*.pdf'))
            print(l)
            l.sort(reverse=False)
            print(l)

            merger = PyPDF2.PdfFileMerger()
            for p in l:
                if not PyPDF2.PdfFileReader(p).isEncrypted:
                    merger.append(p)

            merger.write(dst_path)
            merger.close()

        def get_first_order_id_and_last_order_id(self, pdf_path):
            """pdfから最初の注文IDと最後の注文IDをクラス変数に格納する
            """
            # 与えられたpathからページ数を取得する
            number_of_pdf_pages = PdfEditer.get_pdf_info_from_path(pdf_path)

            # ページを０ページからループさせ、注文番号を見つけると『最後』の注文番号として保存
            for current_page in range(number_of_pdf_pages):
                pdf_text = extract_text(pdf_path, page_numbers=[current_page])
                purchase_id = re.search(r'#\d{5}', pdf_text)
                # purchase_idにデータが格納されているかチェック
                if purchase_id:
                    PdfEditer.last_order_id_in_process = purchase_id.group()
                    break

            # ページを最後からループさせ、注文番号を見つけると『最初』の注文番号として保存
            for current_page in reversed(range(number_of_pdf_pages)):
                pdf_text = extract_text(pdf_path, page_numbers=[current_page])
                purchase_id = re.search(r'#\d{5}', pdf_text)
                # purchase_idにデータが格納されているかチェック
                if purchase_id:
                    PdfEditer.first_order_id_in_process = purchase_id.group()
                    break

        def set_output_pdf_name(self, output_pdf_name):
            """出力PDFの名前を設定する

            Parameters
            ----------
            output_pdf_name : str_
            """
            PdfEditer.cc = canvas.Canvas(output_pdf_name)

        def get_pdf_object_from_path(self, pdf_path, current_page):
            # 現在のページをオブジェクトに
            pp = pagexobj(self.page[current_page])
            PdfEditer.cc.doForm(makerl(PdfEditer.cc, pp))

        @ staticmethod
        def get_pdf_info_from_path(pdf_path):
            """
            pathを渡すとpdfの枚数を返す。
            """
            # 注釈を記入するPDFを開く
            with open(pdf_path, 'rb') as f:
                pdf = PdfFileReader(f)
                document_info = pdf.getDocumentInfo()
                number_of_pages = pdf.getNumPages()

                '''
                pdf_infomation_text = f"""
                PDFファイルパス: {"input.pdf"}:
                タイトル: {document_info.title}
                サブタイトル: {document_info.subject}
                著者: {document_info.author}
                ページ数: {number_of_pages}
                """
                print(pdf_infomation_text)
                '''

            return(number_of_pages)

        def get_value_from_df_and_kay(self, check_df=pd.DataFrame(), search_key='', key_column='', value_column='', add_head_hash=False):
            """与えられた search_key が key_column の中にあるか検索し、ヒットしたら最初に見つけた値を返す
            NaN値は削除して残った一つの値一つ返す
            NaN以外のデータがない場合はエラー

            Parameters
            ----------
            check_df : _type_, optional
                _description_, by default pd.DataFrame()
            search_key : str, optional
                _description_, by default ''
            key_column : str, optional
                _description_, by default ''
            value_column : str, optional
                _description_, by default ''
            add_head_hash : bool, optional
                _description_, by default False
            """
            # データフレームの中に検索ID(注文ID（#付き）)があるものを抽出する, 抽出された個数も返す
            _df = check_df[check_df[key_column].str.contains(
                str(search_key), na=False)]
            _df_count = len(_df.axes[0])

            # #付加オプションが有効で#付きで抽出できない場合,
            # データフレームの中に検索ID(注文ID（#なし）)があるものを抽出する, 抽出された個数も返す
            if (add_head_hash and _df_count == 0):
                added_hash_search_key = re.search(r'\d{5}', str(search_key))
                _df = check_df[check_df[key_column].str.contains(
                    str(added_hash_search_key.group()), na=False)]
                _df_count = len(_df.axes[0])

            # 欠損値（NaN）を除外して値を返す
            for index, row in _df.iterrows():
                if(row[value_column] is not None or row[value_column] is not str):
                    return(row[value_column], _df_count)

            # もし該当する値がない場合はNaNを返す
            return(None, _df_count)

        def auto_indent(self, text, cursor_x=10*mm, cursor_y=70*mm):
            """
            45文字 or \n ごとに自動改行する
            """
            remaining_text = text
            i = 1
            # ゼロ文字になるまで45文字ごとに改行しながら、出力する
            while (len(remaining_text) > 0):
                if "\n" in remaining_text[:27]:
                    PdfEditer.cc.drawString(
                        cursor_x, cursor_y - 7*i*mm, remaining_text[:remaining_text.find("\n")])
                    remaining_text = remaining_text[remaining_text.find("\n") + 2:]
                else:
                    PdfEditer.cc.drawString(
                        cursor_x, cursor_y - 7*i*mm, remaining_text[:27])
                    remaining_text = remaining_text[27:]
                i += 1

        def highlight_text(self, x_cursor, y_cursor, text, text_color, back_color, font_size):
            PdfEditer.cc.setFont("HeiseiKakuGo-W5", font_size)
            PdfEditer.cc.setFillColor(back_color)
            PdfEditer.cc.rect(x_cursor - 1*mm, y_cursor - 1*mm,
                            font_size*4.5, font_size+3*mm, fill=1)
            # PdfEditer.cc.rect(y_cursor - 1, x_cursor - 1, 30, 30, fill=1)
            PdfEditer.cc.setFillColor(text_color)
            PdfEditer.cc.drawString(x_cursor, y_cursor, text)

        def calc_purchase_long(Parsed_pdf_text):
            """
            注文書の長さを確認し、どのページのどの位置から記載すれば見えやすいか返す。

            parameters
            ----------
            text
            """

        def page_extract():
            pass


    def pick_data_from_gui(user_message, file_type):
        global input_file_path
        input_file_path = ''

        # ウインドウ作成
        root = tk.Tk()
        # ウインドウのタイトル
        root.title(user_message)
        # ウインドウサイズと位置指定 幅,高さ,x座標,y座標
        root.geometry("700x150+50+50")

        def get_file_path_trigger_trigger():
            global input_file_path
            f_path = tk.filedialog.askopenfilename(title="ファイル選択", initialdir="ディレクトリを入力", filetypes=[
                ("ファイルを選択してください", file_type)])
            name = os.path.basename(f_path)
            file_path.set(f_path)
            file_name.set(name)
            l_1 = tk.Label(root, textvariable=file_path, relief="flat")
            l_1.place(x=30, y=40)
            input_file_path = f_path
            print(input_file_path)

        def ok_button_trigger():
            global input_file_path
            if input_file_path != '':
                root.destroy()

        # ラベルテキスト更新
        file_path = tk.StringVar()
        file_name = tk.StringVar()

        # ボタン作成
        button = tk.Button(root, text="ファイル選択",
                        command=get_file_path_trigger_trigger)
        button.pack()
        ok_button = tk.Button(root, text="OK", command=ok_button_trigger)
        ok_button.pack(side='bottom', pady=5)

        # イベントループ
        root.mainloop()
        return input_file_path


    if __name__ == "__main__":
        main()

except Exception as e:
    logging.exception("エラーが発生しました：")