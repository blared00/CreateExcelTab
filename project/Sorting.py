
import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox


def save_excel(Window, df, df2):
    try:
        file, _ = QtWidgets.QFileDialog.getSaveFileName(Window,
                                                        'Выбор Excel-файла',
                                                        '',
                                                        'Excel-файлы (*.xlsx);;Все файлы(*.*)')
        writer = pd.ExcelWriter(file, engine='xlsxwriter')
        df.to_excel(writer, 'Таблички',)
        df2.to_excel(writer, 'Свод', )
        worksheet = writer.sheets['Таблички']
        worksheet.set_column(1, 1, 70,)
        worksheet2 = writer.sheets['Свод']
        worksheet2.set_column(1, 1, 70, )
        writer.save()
        print("Я сохранил")
        return True

    except PermissionError:
        QMessageBox.critical(Window, "Ошибка сохранения", "Фаил открыт другим приложением", QMessageBox.Ok)


def read_df(Window, df_address, n):
        df = pd.read_excel(df_address)
        try:
            df = df[['Тип и\nтехническая характеристика', 'Кол.', 'Код 1 С']].dropna()
        except:
            QMessageBox.critical(Window, "Ошибка чтения", "Ошибка чтения файла", QMessageBox.Ok)

        df.rename(columns={'Тип и\nтехническая характеристика': 'Тип и техническая характеристика'}, inplace=True)
        df['Кол. общ'] = df['Кол.']*n
        df['test'] = df['Тип и техническая характеристика'].str.contains('№ камеры')
        df.loc[df['test']==True, 'Кол. общ'] = df.loc[df['test']==True, 'Кол.']
        df.pop('test')
        return df


def create_itog(Window, files_df):
    try:
        df_r = [read_df(Window,sourse_df['sourse'], sourse_df['num_yach']) for sourse_df in files_df]
        df = df_r[0]

        if len(df_r) > 2:
            for df_n in df_r[1:len(df_r)]:
                df = pd.concat([df, df_n], ignore_index=True)
        else:
            df = pd.concat([df, df_r[1]], ignore_index=True) #reset_index().drop(['index'],axis=1)

        #df = df.sort_values(by=['Код 1 С','Тип и техническая характеристика'], ascending=[False, False],)

        final_df = df.drop(['Кол.'], axis=1).groupby('Тип и техническая характеристика').agg({
            'Кол. общ': 'sum',
            'Код 1 С': 'last'

        }).reset_index()
        final_df = final_df.sort_values(by=['Код 1 С', 'Тип и техническая характеристика'], ascending=[False, True], ).reset_index().drop(['index'],axis=1).reset_index()
        final_df['index'] += 1

        final_df.set_index('index', inplace=True)
        final_df.index.name = None
        svod = final_df.copy(deep=True)
        svod = create_svod_table(svod, df_r, files_df)

        s = save_excel(Window, final_df, svod)
        if s:
            print('Стопудов')
            return True
    except:
        QMessageBox.critical(Window, "Ошибка чтения", "Ошибка чтения файла", QMessageBox.Ok)


def create_svod_table(df, df_one_yach,files_df):

    for n in range(len(df_one_yach)):
        df_one = df_one_yach[n].copy(deep=True)
        df_one.rename(columns={'Кол.': f'{files_df[n]["title_cell"]}','Кол. общ': f'{files_df[n]["title_cell"]} общ'}, inplace=True)
        df_one = df_one.drop(['Код 1 С'], axis=1)
        df = df.merge(df_one, on='Тип и техническая характеристика', how='outer')
    df = df.reset_index()
    df['index'] += 1
    df.set_index('index', inplace=True)
    df.index.name = None
    df['1'] = df['Кол. общ']
    df['2'] = df['Код 1 С']
    df = df.drop(['Кол. общ','Код 1 С'], axis=1)
    df.rename(columns={'1': 'Кол. общ','2': 'Код 1 С'}, inplace=True)
    return df



# files_df = [{'sourse': 'C:\\Excel_клеммник\\4132Н\\Таблички\\таблички ОЛ кД 3с 4132Н(яч.16,17,18).xlsx',
#              'num_yach': 3},
#             {'sourse': 'C:\\Excel_клеммник\\4132Н\\Таблички\\таблички ОЛ кД 4132Н(яч.6,7,8,14).xlsx',
#              'num_yach': 4},
#
#             ]
# files_df = [{'sourse': 'C:\\Excel_клеммник\\4132Н\\Таблички\\таблички ОЛ кД 3с 4132Н(яч.16,17,18).xlsx', 'num_yach': 4},
#             {'sourse': 'C:\\Excel_клеммник\\4132Н\\Таблички\\таблички ОЛ кД 4132Н(яч.6,7,8,14).xlsx', 'num_yach': 3}]
# create_itog(files_df)
