# -*- coding: utf-8 -*-
import os
import win32com.client
import pandas as pd
from pandas import Series, DataFrame
from eunjeon import Mecab
# from konlpy.tag import Komoran  # For Test(2021-02-21)
import datetime
import re
import argparse
import time
import multiprocessing
import numpy as np

_version_ = '0.40'


# Version History
# v0.40(2021-08-29): MS Word, PowerPoint, Text 파일에서 단어 추출후 "단어빈도" 시트에 출처(Source) 항목 추가
# v0.30(2021-04-26): DB table, column comment 파일에서 단어 추출후 "단어빈도" 시트에 출처(Source) 항목 추가
# v0.20(2021-02-21): Multiprocessing 적용 버전
# v0.10(2021-01-10): 최초 작성 버전


def get_word_list(df_text) -> DataFrame:
    """
    text 추출결과 DataFrame에서 명사를 추출하여 최종 output을 DataFrame type으로 return
    :param df_text: 파일에서 추출한 text(DataFrame type)
    :return: 명사, 복합어(1개 이상의 명사, 접두사+명사+접미사) 추출결과(Dataframe type)
    """
    start_time = time.time()
    df_result = DataFrame()

    tagger = Mecab()
    # tagger = Komoran()
    row_idx = 0
    for index, row in df_text.iterrows():
        row_idx += 1
        if row_idx % 100 == 0:  # 100건마다 현재 진행상태 출력
            print('[pid:%d] current: %d, total: %d, progress: %3.2f%%' %
                  (os.getpid(), row_idx, df_text.shape[0], round(row_idx / df_text.shape[0] * 100, 2)))
        file_name = row['FileName']
        file_type = row['FileType']
        page = row['Page']
        text = str(row['Text'])
        source = (row['Source'])
        is_db = True if row['FileType'] in ('table', 'column') else False
        is_db_table = True if row['FileType'] == 'table' else False
        is_db_column = True if row['FileType'] == 'column' else False
        if is_db:
            db = row['DB']
            schema = row['Schema']
            table = row['Table']
            if is_db_column:
                column = row['Column']

        if text is None or text.strip() == '':
            continue
        try:
            # nouns = mecab.nouns(text)
            # [O]ToDo: 연속된 체언접두사(XPN), 명사파생접미사(XSN) 까지 포함하여 추출
            # [O]ToDo: 명사(NNG, NNP)가 연속될 때 각각 명사와 연결된 복합명사 함께 추출
            text_pos = tagger.pos(text)
            words = [pos for pos, tag in text_pos if tag in ['NNG', 'NNP', 'SL']]  # NNG: 일반명사, NNP: 고유명사
            pos_list = [x for (x, y) in text_pos]
            tag_list = [y for (x, y) in text_pos]
            pos_str = '/'.join(pos_list) + '/'
            tag_str = '/'.join(tag_list) + '/'
            iterator = re.finditer('(NNP/|NNG/)+(XSN/)*|(XPN/)+(NNP/|NNG/)+(XSN/)*|(SL/)+', tag_str)
            for mo in iterator:
                x, y = mo.span()
                if x == 0:
                    start_idx = 0
                else:
                    start_idx = tag_str[:x].count('/')
                end_idx = tag_str[:y].count('/')
                sub_pos = ''
                # if end_idx - start_idx > 1 and not (start_idx == 0 and end_idx == len(tag_list)):
                if end_idx - start_idx > 1:
                    for i in range(start_idx, end_idx):
                        sub_pos += pos_list[i]
                    # print('%s[sub_pos]' % sub_pos)
                    words.append('%s[복합어]' % sub_pos)  # 추가 형태소 등록

            if len(words) >= 1:
                # print(nouns, text)
                for word in words:
                    # print(noun, '\t', text)
                    if not is_db:
                        df_word = DataFrame(
                            {'FileName': [file_name], 'FileType': [file_type], 'Page': [page], 'Text': [text],
                             'Word': [word], 'Source': [source]})
                    elif is_db_table:
                        df_word = DataFrame(
                            {'FileName': [file_name], 'FileType': [file_type], 'Page': [page], 'Text': [text],
                             'Word': [word], 'DB': [db], 'Schema': [schema], 'Table': [table], 'Source': [source]})
                    elif is_db_column:
                        df_word = DataFrame(
                            {'FileName': [file_name], 'FileType': [file_type], 'Page': [page], 'Text': [text],
                             'Word': [word], 'DB': [db], 'Schema': [schema], 'Table': [table], 'Column': [column],
                             'Source': [source]})
                    df_result = pd.concat([df_result, df_word], ignore_index=True)
        except Exception as ex:
            print('[pid:%d] Exception has raised for text: %s' % (os.getpid(), text))
            print(ex)

    print(
        '[pid:%d] input text count:%d, extracted word count: %d' % (os.getpid(), df_text.shape[0], df_result.shape[0]))
    end_time = time.time()
    # elapsed_time = end_time - start_time
    elapsed_time = str(datetime.timedelta(seconds=end_time - start_time))
    print('[pid:%d] get_word_list finished. total: %d, elapsed time: %s' %
          (os.getpid(), df_text.shape[0], elapsed_time))
    return df_result


def get_current_datetime() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")


def get_ppt_text(file_name) -> DataFrame:
    """
    ppt 파일에서 text를 추출하여 DataFrame type으로 return
    :param file_name: 입력 파일명 (str type)
    :return: 입력 파일에서 추출한 text
    """
    # :return: 입력 파일에서 추출한 text에 형태소 분석기로 명사 추출한 DataFrame
    start_time = time.time()
    print('\r\nget_ppt_text: %s' % file_name)
    ppt_app = win32com.client.Dispatch('PowerPoint.Application')
    ppt_file = ppt_app.Presentations.Open(file_name, True)
    df_text = pd.DataFrame()
    page_count = 0
    for slide in ppt_file.Slides:
        slide_number = slide.SlideNumber
        page_count += 1
        for shape in slide.Shapes:
            shape_text = []
            text = ''
            if shape.HasTable:
                col_cnt = shape.Table.Columns.Count
                row_cnt = shape.Table.Rows.Count
                for row_idx in range(1, row_cnt + 1):
                    for col_idx in range(1, col_cnt + 1):
                        text = shape.Table.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Text
                        if text != '':
                            text = text.replace('\r', ' ')
                            shape_text.append(text)
            elif shape.HasTextFrame:
                for paragraph in shape.TextFrame.TextRange.Paragraphs():
                    text = paragraph.Text
                    if text != '':
                        shape_text.append(text)
            for text in shape_text:
                if text.strip() != '':
                    sr_text = Series([file_name, 'ppt', slide_number, text, f'{file_name}:{slide_number}:{text}'],
                                     index=['FileName', 'FileType', 'Page', 'Text', 'Source'])
                    df_text = df_text.append(sr_text, ignore_index=True)
    ppt_file.Close()
    print('text count: %s' % str(df_text.shape[0]))
    print('page count: %d' % page_count)
    end_time = time.time()
    # elapsed_time = end_time - start_time
    elapsed_time = str(datetime.timedelta(seconds=end_time - start_time))
    print('[pid:%d] get_ppt_text elapsed time: %s' % (os.getpid(), elapsed_time))
    return df_text


def get_doc_text(file_name) -> DataFrame:
    """
    doc 파일에서 text를 추출하여 DataFrame type으로 return
    :param file_name: 입력 파일명 (str type)
    :return: 입력 파일에서 추출한 text
    """
    # :return: 입력 파일에서 추출한 text에 형태소 분석기로 명사 추출한 DataFrame
    start_time = time.time()
    print('\r\nget_doc_text: %s' % file_name)
    word_app = win32com.client.Dispatch("Word.Application")
    word_file = word_app.Documents.Open(file_name, True)
    df_text = pd.DataFrame()
    page = 0
    for paragraph in word_file.Paragraphs:
        text = paragraph.Range.Text
        page = paragraph.Range.Information(3)  # 3: wdActiveEndPageNumber(Text의 페이지번호 확인)
        if text.strip() != '':
            sr_text = Series([file_name, 'doc', page, text, f'{file_name}:{page}:{text}'],
                             index=['FileName', 'FileType', 'Page', 'Text', 'Source'])
            df_text = df_text.append(sr_text, ignore_index=True)

    word_file.Close()
    print('text count: %s' % str(df_text.shape[0]))
    print('page count: %d' % page)
    end_time = time.time()
    # elapsed_time = end_time - start_time
    elapsed_time = str(datetime.timedelta(seconds=end_time - start_time))
    print('[pid:%d] get_doc_text elapsed time: %s' % (os.getpid(), elapsed_time))
    return df_text


def get_txt_text(file_name) -> DataFrame:
    """
    txt 파일에서 text를 추출하여 DataFrame type으로 return
    :param file_name: 입력 파일명 (str type)
    :return: 입력 파일에서 추출한 text
    """
    # :return: 입력 파일에서 추출한 text에 형태소 분석기로 명사 추출한 DataFrame
    start_time = time.time()
    print('\r\nget_txt_text: ' + file_name)
    df_text = pd.DataFrame()
    line_number = 0
    with open(file_name, 'rt', encoding='UTF8') as file:
        for text in file:
            line_number += 1
            if text.strip() != '':
                sr_text = Series([file_name, 'txt', line_number, text, f'{file_name}:{line_number}:{text}'],
                                 index=['FileName', 'FileType', 'Page', 'Text', 'Source'])
                df_text = df_text.append(sr_text, ignore_index=True)
    print('text count: %d' % df_text.shape[0])
    print('line count: %d' % line_number)
    end_time = time.time()
    # elapsed_time = end_time - start_time
    elapsed_time = str(datetime.timedelta(seconds=end_time - start_time))
    print('[pid:%d] get_txt_text elapsed time: %s' % (os.getpid(), elapsed_time))
    return df_text


def make_word_cloud(df_group, now_dt, out_path):
    """
    명사의 빈도를 구한 DataFrame으로 word cloud 그리기
    :param df_group: 명사 빈도 DataFrame
    :param now_dt: 현재 날짜 시각
    :param out_path: 출력경로
    :return: None
    """
    start_time = time.time()
    print('\r\nstart make_word_cloud...')
    from wordcloud import WordCloud
    # malgun.ttf # NanumSquare.ttf # NanumSquareR.ttf NanumMyeongjo.ttf # NanumBarunpenR.ttf # NanumBarunGothic.ttf
    wc = WordCloud(font_path='.\\font\\NanumBarunGothic.ttf',
                   background_color='white',
                   max_words=500,
                   width=1800,
                   height=1000
                   )

    words = df_group.to_dict()['Freq']
    wc.generate_from_frequencies(words)
    wc.to_file('%s\\wordcloud_%s.png' % (out_path, now_dt))
    end_time = time.time()
    # elapsed_time = end_time - start_time
    elapsed_time = str(datetime.timedelta(seconds=end_time - start_time))
    print('make_word_cloud elapsed time: %s' % elapsed_time)


# Todo: 아래아한글 파일(hwp)에서 text 추출
def get_hwp_text(file_name) -> DataFrame:
    pass


# Todo: PDF 파일에서 text 추출
def get_pdf_text(file_name) -> DataFrame:
    pass


# [O]ToDo: Table, column comment에서 text 추출
def get_db_comment_text(file_name) -> DataFrame:
    """
    db_comment 파일에서 text를 추출하여 DataFrame type으로 return
    :param file_name:  입력 파일명 (str type)
    :return: 입력 파일에서 추출한 text
    """
    # :return: 입력 파일에서 추출한 text에 형태소 분석기로 명사 추출한 DataFrame
    start_time = time.time()
    print('\r\nget_db_comment_text: %s' % file_name)
    excel_app = win32com.client.Dispatch('Excel.Application')
    full_path_file_name = os.path.abspath(file_name)
    excel_file = excel_app.Workbooks.Open(full_path_file_name, True)

    # region Table comment
    table_comment_sheet = excel_file.Worksheets(1)
    last_row = table_comment_sheet.Range("A1").End(-4121).Row  # -4121: xlDown
    table_comment_range = 'A2:D%s' % (str(last_row))
    print('table_comment_range : %s (%d rows)' % (table_comment_range, last_row - 1))
    table_comments = table_comment_sheet.Range(table_comment_range).Value2
    df_table = pd.DataFrame(list(table_comments),
                            columns=['DB', 'Schema', 'Table', 'Text'])
    df_table['FileName'] = full_path_file_name
    df_table['FileType'] = 'table'
    df_table['Page'] = 0
    df_table = df_table[df_table.Text.notnull()]  # Text 값이 없는 행 제거
    df_table['Source'] = df_table['DB'] + '.' + df_table['Schema'] + '.' + df_table['Table'] \
                         + '(' + df_table['Text'].astype(str) + ')'
    # print(df_table)
    # endregion

    # region Column comment
    column_comment_sheet = excel_file.Worksheets(2)
    last_row = column_comment_sheet.Range("A1").End(-4121).Row  # -4121: xlDown
    column_comment_range = 'A2:E%s' % (str(last_row))
    print('column_comment_range : %s (%d rows)' % (column_comment_range, last_row - 1))
    column_comments = column_comment_sheet.Range(column_comment_range).Value2
    df_column = pd.DataFrame(list(column_comments),
                             columns=['DB', 'Schema', 'Table', 'Column', 'Text'])
    df_column['FileName'] = full_path_file_name
    df_column['FileType'] = 'column'
    df_column['Page'] = 0
    df_column = df_column[df_column.Text.notnull()]  # Text 값이 없는 행 제거
    df_column['Source'] = df_column['DB'] + '.' + df_column['Schema'] + '.' + df_column['Table'] \
                          + '.' + df_column['Column'] + '(' + df_column['Text'].astype(str) + ')'
    # print(df_column)
    # endregion

    excel_file.Close()
    df_text = df_column.append(df_table, ignore_index=True)
    end_time = time.time()
    # elapsed_time = end_time - start_time
    elapsed_time = str(datetime.timedelta(seconds=end_time - start_time))
    print('[pid:%d] get_db_comment_text elapsed time: %s' % (os.getpid(), elapsed_time))
    print('text count: %s' % str(df_text.shape[0]))
    return df_text


def get_file_text(file_name) -> DataFrame:
    """
    MS Word, PowerPoint, Text, DB Comment(Excel) file에서 text를 추출하는 함수
    :param file_name: 파일명
    :return: file에서 추출한 text(DataFrame type)
    """
    df_text = DataFrame()
    if file_name.endswith(('.doc', '.docx')):
        df_text = get_doc_text(file_name)
    elif file_name.endswith(('.ppt', '.pptx')):
        df_text = get_ppt_text(file_name)
    elif file_name.endswith('.txt'):
        df_text = get_txt_text(file_name)
    elif file_name.endswith(('.xls', '.xlsx', '.xlsb')):
        df_text = get_db_comment_text(file_name)
    return df_text


def main():
    """
    지정한 경로 하위 폴더의 File들에서 Text를 추출하고 각 Text의 명사를 추출하여 엑셀파일로 저장
    :return: 없음
    """

    # region Args Parse & Usage set-up -------------------------------------------------------------
    # parser = argparse.ArgumentParser(usage='usage test', description='description test')
    usage_description = """--- Description ---
  * db_comment_file과 in_path중 하나는 필수로 입력

  * 실행 예시
    1. File에서 text, 단어 추출: in_path, out_path 지정
       python word_extractor.py --multi_process_count 4 --in_path .\\test_files --out_path .\out

    2. DB comment에서 text, 단어 추출: db_comment_file, out_path 지정
       python word_extractor.py --db_comment_file "table,column comments.xlsx" --out_path .\out

    3. File, DB comment 에서 text, 단어 추출: db_comment_file, in_path, out_path 지정
       python word_extractor.py --db_comment_file "table,column comments.xlsx" --in_path .\test_files --out_path .\out

  * DB Table, Column comment 파일 형식
    - 첫번째 sheet(Table comment): DBName, SchemaName, Tablename, TableComment
    - 두번째 sheet(Column comment): DBName, SchemaName, Tablename, ColumnName, ColumnComment"""

    # ToDo: 옵션추가: 복합어 추출할지 여부, 영문자 추출할지 여부, 영문자 길이 1자리 제외여부, ...
    parser = argparse.ArgumentParser(description=usage_description, formatter_class=argparse.RawTextHelpFormatter)
    # name argument 추가
    parser.add_argument('--multi_process_count', required=False, type=int,
                        help='text 추출, 단어 추출을 동시에 실행할 multi process 개수(지정하지 않으면 (logical)cpu 개수로 설정됨)')
    parser.add_argument('--db_comment_file', required=False,
                        help='DB Table, Column comment 정보 파일명(예: comment.xlsx)')
    parser.add_argument('--in_path', required=False, help='입력파일(ppt, doc, txt) 경로명(예: .\in) ')
    parser.add_argument('--out_path', required=True, help='출력파일(xlsx, png) 경로명(예: .\out)')

    args = parser.parse_args()

    multi_process_count = int(args.multi_process_count)
    if multi_process_count is None:
        multi_process_count = multiprocessing.cpu_count()

    db_comment_file = args.db_comment_file
    if db_comment_file is not None and not os.path.isfile(db_comment_file):
        print('db_comment_file not found: %s' % db_comment_file)
        exit(-1)

    in_path = args.in_path
    out_path = args.out_path
    print('------------------------------------------------------------')
    print('Word Extractor v%s start --- %s' % (_version_, get_current_datetime()))
    print('##### arguments #####')
    print('multi_process_count: %d' % multi_process_count)
    print('db_comment_file: %s' % db_comment_file)
    print('in_path: %s' % in_path)
    print('out_path: %s' % out_path)
    print('------------------------------------------------------------')
    # endregion Args Parse & Usage set-up -------------------------------------------------------------

    start_time = time.time()

    df_text = DataFrame()  # 파일에서 읽은 text
    df_result = DataFrame()  # df_text에서 추출한 단어
    file_list = []
    if in_path is not None and in_path.strip() != '':
        print('[%s] Start Get File List...' % get_current_datetime())
        in_abspath = os.path.abspath(in_path)  # os.path.abspath('.') + '\\test_files'
        file_types = ('.ppt', '.pptx', '.doc', '.docx', '.txt')
        for root, dir, files in os.walk(in_abspath):
            for file in sorted(files):
                # 제외할 파일
                if file.startswith('~'):
                    continue
                # 포함할 파일
                if file.endswith(file_types):
                    file_list.append(root + '\\' + file)

        print('[%s] Finish Get File List.' % get_current_datetime())
        print('--- File List ---')
        print('\n'.join(file_list))

    if db_comment_file is not None:
        file_list.append(db_comment_file)

    # ---------- text 추출 병렬 실행 ----------
    print('[%s] Start Get File Text...' % get_current_datetime())
    with multiprocessing.Pool(processes=multi_process_count) as pool:
        mp_text_result = pool.map(get_file_text, file_list)
    df_text = pd.concat(mp_text_result, ignore_index=True)
    print('[%s] Finish Get File Text.' % get_current_datetime())
    # 여기까지 text 추출완료. 아래에 단어 추출 시작

    # ---------- 단어 추출 병렬 실행 ----------
    print('[%s] Start Get Word from File Text...' % get_current_datetime())
    df_text_split = np.array_split(df_text, multi_process_count)
    # mp_result = []
    with multiprocessing.Pool(processes=multi_process_count) as pool:
        mp_result = pool.map(get_word_list, df_text_split)

    df_result = pd.concat(mp_result, ignore_index=True)
    if 'DB' not in df_result.columns:
        df_result['DB'] = ''
        df_result['Schema'] = ''
        df_result['Table'] = ''
        df_result['Column'] = ''

    print('[%s] Finish Get Word from File Text.' % get_current_datetime())
    # ------------------------------

    print('[%s] Start Get Word Frequency...' % get_current_datetime())
    df_result_subset = df_result[['Word', 'Source']]  # 빈도수를 구하기 위해 필요한 column만 추출
    df_group = df_result_subset.groupby(by='Word').agg(['count', lambda x: '\n'.join(list(x)[:10])])
    df_group.index.name = 'Word'  # index명 재지정
    df_group.columns = ['Freq', 'Source']  # column명 재지정
    df_group = df_group.sort_values(by='Freq', ascending=False)
    print('[%s] Finish Get Word Frequency.' % get_current_datetime())
    print('[%s] Start Make Word Cloud...' % get_current_datetime())
    now_dt = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    make_word_cloud(df_group, now_dt, out_path)
    print('[%s] Finish Make Word Cloud.' % get_current_datetime())

    print('[%s] Start Save the Extract result to Excel File...' % get_current_datetime())
    df_result.index += 1
    excel_style = {
        'font-size': '10pt'
    }
    df_result = df_result.style.set_properties(**excel_style)
    df_group = df_group.style.set_properties(**excel_style)
    out_file_name = '%s\\extract_result_%s.xlsx' % (out_path, now_dt)  # 'out\\extract_result_%s.xlsx' % now_dt

    print('start writing excel file...')
    with pd.ExcelWriter(path=out_file_name, engine='xlsxwriter') as writer:
        df_result.to_excel(writer,
                           header=True,
                           sheet_name='단어추출결과',
                           index=True,
                           index_label='No',
                           freeze_panes=(1, 0),
                           columns=['Word', 'FileName', 'FileType', 'Page', 'Text', 'DB', 'Schema', 'Table', 'Column'])
        df_group.to_excel(writer,
                          header=True,
                          sheet_name='단어빈도',
                          index=True,
                          index_label='단어',
                          freeze_panes=(1, 0))
        workbook = writer.book
        worksheet = writer.sheets['단어빈도']
        wrap_format = workbook.add_format({'text_wrap': True})
        worksheet.set_column("C:C", None, wrap_format)

    print('[%s] Finish Save the Extract result to Excel File...' % get_current_datetime())

    end_time = time.time()
    # elapsed_time = end_time - start_time
    elapsed_time = str(datetime.timedelta(seconds=end_time - start_time))
    print('------------------------------------------------------------')
    print('[%s] Finished.' % get_current_datetime())
    print('overall elapsed time: %s' % elapsed_time)
    print('------------------------------------------------------------')


if __name__ == '__main__':
    main()
