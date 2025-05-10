#!/usr/bin/env python3
"""
MD to Excel Converter

このスクリプトはMarkdownファイルを解析してExcelファイルに変換します。
Markdownの見出し（#）は階層構造として解釈し、表（|）は対応するExcelの表として変換されます。
また、箇条書き（-、*、+）もサポートしています。
"""

import os
import re
import sys
import argparse
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class MdToExcelConverter:
    def __init__(self, md_file, output_file=None, sheet_name="Sheet1"):
        """
        初期化メソッド
        
        Parameters:
            md_file (str): 入力Markdownファイルのパス
            output_file (str, optional): 出力Excelファイルのパス。デフォルトはNone（MDファイル名から自動生成）
            sheet_name (str, optional): 出力Excelファイルのシート名。デフォルトは"Sheet1"
        """
        self.md_file = md_file
        # 出力ファイル名が指定されていない場合は入力ファイル名から生成
        if output_file is None:
            md_path = Path(md_file)
            self.output_file = str(md_path.with_suffix('.xlsx'))
        else:
            self.output_file = output_file
        self.sheet_name = sheet_name
        self.sections = []  # 各セクション（見出し、段落、表など）を格納
        self.current_row = 1

    def parse_markdown(self):
        """
        Markdownファイルを解析し、セクション構造として抽出する
        """
        print(f"Markdownファイルを解析中: {self.md_file}")
        
        try:
            with open(self.md_file, 'r', encoding='utf-8') as file:
                lines = file.readlines()
        except Exception as e:
            print(f"エラー: ファイル '{self.md_file}' を開けませんでした。 {str(e)}")
            sys.exit(1)

        current_section = None
        current_heading_level = 0
        in_table = False
        current_table = []
        paragraph_lines = []
        in_list = False
        current_list = []
        
        for line in lines:
            line = line.rstrip()
            
            # 見出しの検出
            heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if heading_match:
                # 前の段落を処理
                if paragraph_lines:
                    if current_section:
                        current_section['paragraphs'].append('\n'.join(paragraph_lines))
                    paragraph_lines = []
                
                # 前のリストを処理
                if in_list and current_list:
                    if current_section:
                        current_section['lists'].append(current_list)
                    current_list = []
                    in_list = False
                
                # 前の表を処理
                if in_table and current_table:
                    if current_section:
                        current_section['tables'].append(current_table)
                    current_table = []
                    in_table = False
                
                # 新しいセクションを作成
                level = len(heading_match.group(1))
                text = heading_match.group(2)
                
                current_section = {
                    'heading': text,
                    'level': level,
                    'paragraphs': [],
                    'lists': [],
                    'tables': []
                }
                
                self.sections.append(current_section)
                current_heading_level = level
                continue
            
            # 表の処理
            if line.startswith('|') and line.endswith('|'):
                # 前の段落を処理
                if paragraph_lines:
                    if current_section:
                        current_section['paragraphs'].append('\n'.join(paragraph_lines))
                    paragraph_lines = []
                
                # 前のリストを処理
                if in_list and current_list:
                    if current_section:
                        current_section['lists'].append(current_list)
                    current_list = []
                    in_list = False
                
                if not in_table:
                    # 新しい表の開始
                    in_table = True
                    current_table = []
                
                # 表の行を追加
                cells = [cell.strip() for cell in line.strip('|').split('|')]
                
                # セパレータ行（---）をスキップ
                if all(re.match(r'^[-:\s]+$', cell) for cell in cells):
                    continue
                
                current_table.append(cells)
                continue
            elif in_table:
                # 表の終了
                if current_section and current_table:
                    current_section['tables'].append(current_table)
                current_table = []
                in_table = False
            
            # 箇条書きの処理
            list_match = re.match(r'^(\s*)[-*+]\s+(.+)$', line)
            if list_match:
                # 前の段落を処理
                if paragraph_lines:
                    if current_section:
                        current_section['paragraphs'].append('\n'.join(paragraph_lines))
                    paragraph_lines = []
                
                if not in_list:
                    # 新しいリストの開始
                    in_list = True
                    current_list = []
                
                # リストアイテムを追加
                indent = len(list_match.group(1))
                content = list_match.group(2)
                current_list.append((indent, content))
                continue
            elif in_list and line.strip():
                # リストの終了（空行でない別の内容があれば終了）
                if not re.match(r'^\s*[-*+]', line):
                    if current_section and current_list:
                        current_section['lists'].append(current_list)
                    current_list = []
                    in_list = False
                    
                    # 次の処理に続く（この行は通常のテキストとして扱う）
                else:
                    # まだリスト内の行である
                    continue
            
            # 空行の処理
            if not line.strip():
                # 前の段落を処理
                if paragraph_lines:
                    if current_section:
                        current_section['paragraphs'].append('\n'.join(paragraph_lines))
                    paragraph_lines = []
                
                # 前のリストを処理
                if in_list and current_list:
                    if current_section:
                        current_section['lists'].append(current_list)
                    current_list = []
                    in_list = False
                    
                continue
            
            # 通常のテキスト
            if not in_list:
                paragraph_lines.append(line)
        
        # 残りの段落、リスト、表を処理
        if paragraph_lines:
            if current_section:
                current_section['paragraphs'].append('\n'.join(paragraph_lines))
        
        if in_list and current_list:
            if current_section:
                current_section['lists'].append(current_list)
        
        if in_table and current_table:
            if current_section:
                current_section['tables'].append(current_table)
        
        # 結果を出力
        print(f"セクション数: {len(self.sections)}")
        for i, section in enumerate(self.sections):
            print(f"  セクション {i+1}: {section['heading']} (レベル {section['level']})")
            print(f"    段落数: {len(section['paragraphs'])}")
            print(f"    リスト数: {len(section['lists'])}")
            print(f"    表数: {len(section['tables'])}")

    def _get_column_width(self, text):
        """
        テキストの長さに基づいてカラム幅を計算する
        
        Parameters:
            text (str): セルのテキスト
            
        Returns:
            float: 推奨されるカラム幅
        """
        if text is None:
            return 10.0
            
        # 文字列に変換
        text = str(text)
        
        # 日本語の文字は英語の約2倍の幅として計算
        # 日本語文字の検出（全角文字）
        japanese_chars = len(re.findall(r'[\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\u4e00-\u9fff\uff00-\uff9f]', text))
        # 英数字の検出
        english_chars = len(text) - japanese_chars
        
        # 基本幅 + 日本語文字の追加幅 + 英数字の幅
        width = 1.0 + (japanese_chars * 2.0) + (english_chars * 1.0)
        return min(max(width, 10.0), 50.0)  # 最小10、最大50の幅に制限

    def create_excel(self):
        """
        解析したMarkdownの内容からExcelファイルを生成する
        """
        print(f"Excelファイルを作成中: {self.output_file}")
        
        # Excelワークブックの作成
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = self.sheet_name
        
        # スタイルの定義
        heading_fonts = {
            1: Font(bold=True, size=16, color="000000"),
            2: Font(bold=True, size=14, color="000000"),
            3: Font(bold=True, size=12, color="000000"),
            4: Font(bold=True, size=11, color="000000"),
            5: Font(bold=True, size=10, color="000000"),
            6: Font(bold=True, size=10, color="000000")
        }
        
        heading_fills = {
            1: PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"),
            2: PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
            3: PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
            4: PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
            5: PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"),
            6: PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        }
        
        # 境界線スタイル
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # ヘッダースタイル（表用）
        header_font = Font(bold=True)
        header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        
        # 中央揃え
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        # 箇条書きのスタイル
        list_font = Font(name='Calibri', size=11)
        
        # 各セクションを追加
        for section in self.sections:
            # 見出しを追加
            sheet.cell(row=self.current_row, column=1, value=section['heading'])
            
            # 見出しのスタイル設定
            cell = sheet.cell(row=self.current_row, column=1)
            cell.font = heading_fonts.get(section['level'], Font(bold=True))
            cell.fill = heading_fills.get(section['level'], PatternFill())
            
            # インデント設定（レベルに応じて）
            cell.alignment = Alignment(indent=section['level']-1)
            
            # 見出しの行の高さを調整
            sheet.row_dimensions[self.current_row].height = 24
            
            self.current_row += 1
            
            # 段落を追加
            for paragraph in section['paragraphs']:
                if paragraph.strip():  # 空の段落はスキップ
                    sheet.cell(row=self.current_row, column=1, value=paragraph)
                    self.current_row += 1
            
            # リストを追加
            for list_items in section['lists']:
                for indent, content in list_items:
                    # インデントレベルを計算（スペース4つを1レベルと仮定）
                    indent_level = max(1, indent // 2 + 1)
                    
                    # リストアイテムを追加
                    cell = sheet.cell(row=self.current_row, column=1, value=f"• {content}")
                    cell.font = list_font
                    cell.alignment = Alignment(indent=indent_level)
                    
                    self.current_row += 1
                
                # リストの後に空行を追加
                self.current_row += 1
            
            # 表を追加
            for table in section['tables']:
                if not table:  # 空の表はスキップ
                    continue
                    
                # ヘッダー行が存在するか確認
                if len(table) > 0:
                    # 表のヘッダー行
                    for col_index, header in enumerate(table[0], start=1):
                        cell = sheet.cell(row=self.current_row, column=col_index, value=header)
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.alignment = center_alignment
                        cell.border = thin_border
                        
                        # カラム幅の調整
                        col_letter = get_column_letter(col_index)
                        sheet.column_dimensions[col_letter].width = self._get_column_width(header)
                    
                    self.current_row += 1
                    
                    # 表のデータ行
                    for row_data in table[1:]:
                        for col_index, cell_value in enumerate(row_data, start=1):
                            # 列インデックスが範囲外の場合は空白セルを追加
                            if col_index <= len(row_data):
                                cell = sheet.cell(row=self.current_row, column=col_index, value=cell_value)
                                cell.border = thin_border
                                
                                # カラム幅の再調整
                                col_letter = get_column_letter(col_index)
                                current_width = sheet.column_dimensions[col_letter].width
                                new_width = self._get_column_width(cell_value)
                                sheet.column_dimensions[col_letter].width = max(current_width, new_width)
                        
                        self.current_row += 1
                
                # 表の後に空行を追加
                self.current_row += 1
        
        # ファイルの保存
        try:
            workbook.save(self.output_file)
            print(f"ファイルを保存しました: {self.output_file}")
        except Exception as e:
            print(f"エラー: ファイルの保存中にエラーが発生しました: {str(e)}")
            sys.exit(1)

    def convert(self):
        """
        変換プロセスを実行する
        """
        self.parse_markdown()
        self.create_excel()
        return self.output_file


def main():
    """
    メイン実行関数
    コマンドライン引数を解析して処理を実行する
    """
    parser = argparse.ArgumentParser(description="Markdownファイルをエクセルに変換するプログラム")
    parser.add_argument("input", help="入力Markdownファイルのパス")
    parser.add_argument("-o", "--output", help="出力Excelファイルのパス（指定しない場合は入力ファイルと同じ名前で.xlsxとして保存）")
    parser.add_argument("-s", "--sheet", default="Sheet1", help="出力Excelファイルのシート名（デフォルト: Sheet1）")
    
    args = parser.parse_args()
    
    converter = MdToExcelConverter(args.input, args.output, args.sheet)
    output_file = converter.convert()
    
    print(f"変換完了: {output_file}")


if __name__ == "__main__":
    main()
