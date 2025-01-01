# -*- coding: utf-8 -*-
import os
from tqdm import tqdm
import pandas as pd

def get_total_courses():
    if not os.path.exists('StarClass_List.xls'):
        print('請先下載全校開課總表到此資料夾下')
        exit()
    if os.path.exists('總表.xlsx'):
        if 'Y' != input('總表已存在，是否要取代(Y/N)? '):
            exit()
        else:
            os.remove('總表.xlsx')
    df = pd.read_excel('StarClass_List.xls')
    df.columns = df.iloc[1]
    df = df.drop(index = 1).reset_index(drop = True) # 刪掉格式不對齊的第一行
    return df

def merge(total_course):
    # 合併所有修課名單至df
    class_code_UPs = []
    files = os.listdir()
    with tqdm(files, desc = '進度條') as pbar:
        for file in pbar:
            if file.endswith('_修課名單.xls'):
                # 設定進度調顯示訊息
                pbar.set_postfix({'正在處理': file})
                
                # 讀檔與初始化
                single_course_df = pd.read_excel(file)
                # 新增欄位
                single_course_df.insert(single_course_df.columns.get_loc('課程代碼') + 1, '系所課程代碼', '')
                single_course_df.insert(single_course_df.columns.get_loc('課程名稱') + 1, '必選修', '')
                single_course_df.insert(single_course_df.columns.get_loc('必選修') + 1, '開課學系', '')
                single_course_df.insert(single_course_df.columns.get_loc('開課學系') + 1, '開課班級', '')
                # 替換課程代碼開頭
                code_without_UP = file.replace('UP', '').replace('_修課名單.xls', '') # ex. 'UP106G_修課名單.xls' -> '106G'
                mapping = {'工業與系統工程學系': 'IE', '電子工程學系': 'EL', '電機工程學系': 'EE'} # 學系對應的課程代碼編號
                
                # 到全校開課總表找到課程對應的開課學系置換課程代碼
                for m_idx, m_row in single_course_df.iterrows():
                    for _, tc_row in total_course.iterrows():
                        if m_row['課程名稱'] == tc_row['課程名稱(中)'] and \
                            tc_row['開課學系'] in ['電機工程學系', '電子工程學系', '工業與系統工程學系']:
                            single_course_df.at[m_idx, '系所課程代碼'] = mapping[tc_row['開課學系']] + code_without_UP
                            single_course_df.at[m_idx, '必選修'] = tc_row['必選修']
                            single_course_df.at[m_idx, '開課學系'] = tc_row['開課學系']
                            single_course_df.at[m_idx, '開課班級'] = tc_row['開課班級']
                            break
                class_code_UPs.append(single_course_df)
    merged_df = pd.concat(class_code_UPs, axis = 0) # 合併
    
    # 寫檔
    merged_df = merged_df.apply(lambda col: col.apply(lambda x: str(x) if pd.notna(x) else x)) # 將非空白欄位轉換成字串寫入以避免被轉換成科學記號
    merged_df.to_excel('總表.xlsx', index = False, sheet_name = 'Sheet0', engine = 'xlsxwriter')

def main():
    merge(get_total_courses())

if __name__ == '__main__':
    main()