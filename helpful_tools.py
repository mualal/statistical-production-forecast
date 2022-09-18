import os
import pandas as pd
import numpy as np
import xlwings as xw


def history_preprocessing(
    history: pd.DataFrame,
    max_delta=float('inf')
) -> pd.DataFrame:
    """
    предварительная обработка МЭР
    @param history: датафрейм со столбцами - № скважины; дата; добыча нефти за последний месяц, т;
    добыча жидкости за последний месяц, т; время работы в добыче, часы; объекты работы;
    координата забоя Y (по траектории); координата забоя X (по траектории)
    @param max_delta: максимальный период остановки, дни
    @return: обрезанный и структурированный по скважинам и дням датафрейм
    """
    # заполнение пустых ячеек нулями
    history.fillna(0, inplace=True)
    # массив уникальных значений дат (массив элементов типа np.datetime64)
    current_date = history['Дата'].unique()
    # сортировка по возрастанию дат
    current_date.sort(axis=0)

    # удаление нулевых строк
    history = history[(history['Добыча нефти за посл.месяц, т'] != 0) &
                      (history['Добыча жидкости за посл.месяц, т'] != 0) &
                      (history['Время работы в добыче, часы'] != 0) &
                      (history['Объекты работы'] != 0)]
    
    # массив уникальных скважин
    unique_wells = history['№ скважины'].unique()
    # сортировка по возрастанию номеров скважин и затем дат
    history = history.sort_values(['№ скважины', 'Дата'])
    
    # создание нового структурированного датафрейма
    history_new = pd.DataFrame()
    for i in unique_wells:
        # слайс данных для рассматриваемой скважины
        slice = history[history['№ скважины'] == i].copy()
        # пласт (или пласты), на которые работала скважина в последний месяц
        object = slice['Объекты работы'].iloc[-1]
        # история работы только на этот пласт (или эти пласты)
        slice = slice[slice['Объекты работы'] == object]
        # расчёт времени работы (в днях) до следующей строки данных с измерениями
        next_date = slice["Дата"].iloc[1:]
        last_date = slice["Дата"].iloc[-1]
        next_date.loc[-100] = last_date + pd.DateOffset(months=1)
        slice["След. дата"] = np.array(next_date)
        slice["Разница дат"] = slice["След. дата"] - slice["Дата"]

        # если скважина работала меньше суток в последний месяц - он удаляется
        if slice["Время работы в добыче, часы"].iloc[-1] < 24:
            slice = slice.iloc[:-1]

        # обрезка истории, если скважина была остановлена больше max_delta (или не проводились новые измерения)
        if not slice[slice["Разница дат"] > np.timedelta64(max_delta, 'D')].empty:
            last_index = slice[slice["Разница дат"] > np.timedelta64(max_delta, 'D')].index.tolist()[-1]
            slice = slice.loc[last_index + 1:]
        
        # конкатенация данных по ранее обработанным скважинам с данными по только что обработанной скважине
        history_new = pd.concat([history_new, slice], ignore_index=True)
    
    # удаление вспомогательных столбцов с разницей дат
    history_new.drop(['Разница дат', 'След. дата'], axis=1, inplace=True)

    return history_new


if __name__ == "__main__":
    # чтение данных из файла
    df_initial = pd.read_excel(os.path.join('data', 'Западно-Чистинное-Ю1(2)-МЭР.xlsx'), sheet_name='МЭР')
    # обработка и структурирование данных
    df_initial = history_preprocessing(df_initial, max_delta=365)
    # отображение обработанных и структурированных данных в MS Excel
    xw.view(df_initial)
