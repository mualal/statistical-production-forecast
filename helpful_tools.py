import os
import pandas as pd
import numpy as np
import xlwings as xw


def history_preprocessing(
    history: pd.DataFrame,
    max_delta=float('inf')
):
    """
    предварительно обработать МЭР
    @param history: датафрейм со столбцами - № скважины; дата; добыча нефти за последний месяц, т;
    добыча жидкости за последний месяц, т; время работы в добыче, часы; объекты работы;
    координата забоя Y (по траектории); координата забоя X (по траектории)
    @param max_delta: максимальный период остановки, дни
    @return: обрезанный датафрейм
    """
    last_data = history['Дата'].unique()
    last_data.sort(axis=0)
    history = history.fillna(0)  # заполнить пустые ячейки нулями

    history = history[(history['Добыча нефти за посл.месяц, т'] != 0) &
                      (history['Добыча жидкости за посл.месяц, т'] != 0) &
                      (history['Время работы в добыче, часы'] != 0) &
                      (history['Объекты работы'] != 0)]  # оставить ненулевые строки

    unique_wells = history['№ скважины'].unique()  # уникальный список скважин (без нулевых значений)
    history = history[history['№ скважины'].isin(unique_wells)]  # выделить историю только этих скважин
    history = history.sort_values(['№ скважины', 'Дата'])

    history_new = pd.DataFrame()
    for i in unique_wells:
        slice = history.loc[history['№ скважины'] == i].copy()
        object = slice['Объекты работы'].iloc[-1]
        slice = slice[slice['Объекты работы'] == object]
        sec_date = slice["Дата"].iloc[1:]
        last_data = slice["Дата"].iloc[-1]
        sec_date.loc[-1] = last_data
        slice["След. дата"] = list(sec_date)
        slice["Разница дат"] = slice["След. дата"] - slice["Дата"]

        # если скважина работала меньше суток в последний месяц - он удаляется
        if slice["Время работы в добыче, часы"].iloc[-1] < 24:
            slice = slice.iloc[:-1]

        # обрезка истории, если скважина была остановлена больше max_delta
        if not slice[slice["Разница дат"] > np.timedelta64(max_delta, 'D')].empty:
            last_index = slice[slice["Разница дат"] > np.timedelta64(max_delta, 'D')].index.tolist()[-1]
            slice = slice.loc[last_index + 1:]
        history_new = pd.concat([history_new, slice], ignore_index=True)

    del history_new["Разница дат"]
    del history_new["След. дата"]
    return history_new


if __name__ == "__main__":
    df_initial = pd.read_excel(os.path.join('data', 'Западно-Чистинное-Ю1(2)-МЭР.xlsx'), sheet_name='МЭР')
    df_initial = history_preprocessing(df_initial, max_delta=365)
    xw.view(df_initial)
