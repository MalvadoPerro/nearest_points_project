# Импортируем все неоходимые библиотека
import pandas as pd
import numpy as np
from geopy.distance import geodesic as gd  # Функция для рассчета расстояния между координатами


def create_dfs(path_a, path_b):  # Создаем функцию по созданию фреймов из файлов с координатами
    global df_stations, df_attractions
    df_stations = pd.read_excel(path_a)
    df_attractions = pd.read_excel(path_b)


def check_dfs():  # Функция для проверки фреймов, которые были выбраны пользователем
    global df_stations, df_attractions
    check_columns = ['Название', 'Координаты']
    if list(df_stations.columns) == check_columns and list(df_attractions.columns) == check_columns:
        return True
    else:
        return False


def get_points_df_result(count):
    global count_points, df_result
    count_points = count  # Получаем количество искомых ближайших точек

    # Создаем результирующий фрейм
    arr_columns = ['Название', 'Координаты']
    for i in range(count_points):
        arr_columns.append(f'Место_{i + 1}')
        arr_columns.append(f'Координаты_{i + 1}')
        arr_columns.append(f'Расстояние_{i + 1}')
    df_result = pd.DataFrame(columns=arr_columns)


# Функция подсчета
def calculate():
    global count_points, df_result
    global df_stations, df_attractions

    for i in df_stations.values:
        data = {'Место': [], 'Координаты': [], 'Расстояние': []}
        df_time = pd.DataFrame(data)  # Создаем временный фрейм для записи n ближайших точек к каждой станции
        for j in df_attractions.values:
            arr = [j[0], j[1], round(gd(i[1], j[1]).km, 3)]  # Список с названием, координатами
            # достопримечательности и расстоянием до нее
            df_time.loc[len(df_time.index)] = arr  # Добавляем во временный фрейм строку, которую создали выше
        df_time.sort_values(by=['Расстояние'], inplace=True)  # Фильтруем фрейм по возрастанию расстояния
        df_time = df_time.head(count_points).reset_index(drop=True)  # Оставляем во фрейме первые n строк
        for j in df_time.values:
            i = np.append(i, j)  # Добавляем к станции и ее координатам строки, полученные ранне и записанные во
            # временный фрейм
        df_result.loc[len(df_result.index)] = i  # Добавляем полученную строку в результирующий фрейм


# Создаем функцию сохранения результирующего фрейма
def save_result(path_save):
    global df_result

    writer = pd.ExcelWriter(path_save, engine='xlsxwriter')  # Создаем драйвер для записи фрейма в эксель
    df_result.to_excel(writer, sheet_name='Sheet1', index_label='№')  # Записываем фрейм в эксель

    # Автоматически настраиваем ширину столбцов по их содержимому
    for column in df_result:
        column_width = max(df_result[column].astype(str).map(len).max(), len(column))
        col_idx = df_result.columns.get_loc(column) + 1
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width + 1.5)

    writer.close()

# create_dfs('Stations.xlsx', 'Attractions.xlsx')
# print(check_dfs())
# get_points_df_result(4)
# calculate()
# save_result('D:/nearest_points/output/nearest_points_111.xlsx')
