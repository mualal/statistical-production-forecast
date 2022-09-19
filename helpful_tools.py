import os
import pandas as pd
import numpy as np
from scipy import interpolate
from sklearn.linear_model import LinearRegression
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


def interpolate_gur(
    x,
    y,
    table_x,
    table_y,
    table_z
) -> tuple:

    table_x = np.reshape(np.array(table_x, dtype='float64'), (-1,))
    table_y = np.reshape(np.array(table_y, dtype='float64'), (-1,))
    table_z = np.reshape(np.array(table_z, dtype='float64'), (-1,))

    if len(table_x) <= 16:
        gur_1 = interpolate.interp2d(table_x, table_y, table_z, kind='linear')
        gur_1 = gur_1(x, y)[0]
        gur_2 = gur_1
    else:
        gur_1 = interpolate.griddata(
            (table_x, table_y),
            table_z,
            (x, y),
            method='cubic'
        )
        gur_2 = interpolate.interp2d(table_x, table_y, table_z, kind='cubic')
        gur_2 = gur_2(x, y)[0]
    
    return gur_1, gur_2


def linear_model(
    df: pd.DataFrame,
    param
) -> tuple:
    match param:
        case 'Nazarov_Sipachev':
            x = df['qnak_water'].values.reshape((-1, 1))
            y = df['y']
        case 'Sipachev_Pasevich':
            x = df['qnak_liq'].values.reshape((-1, 1))
            y = df['y']
        case 'FNI':
            x = df['qnak_oil'].values.reshape((-1, 1))
            y = df['y']
        case 'Maksimov':
            x = df['qnak_oil'].values.reshape((-1, 1))
            y = df['log_Qw']
        case 'Sazonov':
            x = df['qnak_oil'].values.reshape((-1, 1))
            y = df['log_Ql']
    
    model = LinearRegression().fit(x, y)
    a = model.intercept_
    b = model.coef_
    b = float(b)
    b = np.fabs(b)
    q_0 = df['qnak_oil'].values[-1]
    if b != 0:
        match param:
            case 'Nazarov_Sipachev':
                q_izv = (1 / b) * (1 - ((a - 1) * (1 - 0.99) / 0.99) ** 0.5)
            case 'Sipachev_Pasevich':
                q_izv = (1 / b) - ((0.01 * a) / (b ** 2)) ** 0.5
            case 'FNI':
                q_izv = 1 / (2 * b * (1 - 0.99)) - a / 2 * b
            case 'Maksimov' | 'Sazonov':
                q_izv = (1 / b) * np.log(0.99 / ((1 - 0.99) * b * np.exp(a)))
        oiz = q_izv - q_0  # остаточные извлекаемые запасы нефти
    else:
        q_izv = 0
        oiz = 0
    korrelation = np.fabs(np.corrcoef(df['qnak_water'], df['y'])[1, 0])
    determination = model.score(x, y)

    return q_izv, oiz, korrelation, determination


def calculate_reserves_statistics(
    df_well: pd.DataFrame,
    name_well,
    marker=0
) -> tuple:

    error = ''

    df_well['qnak_oil'] = df_well['Добыча нефти за посл.месяц, т'].cumsum()
    df_well['qnak_liq'] = df_well['Добыча жидкости за посл.месяц, т'].cumsum()
    df_well['qnak_water'] = df_well['qnak_liq'] - df_well['qnak_oil']

    df_well['y'] = df_well['qnak_liq'] / df_well['qnak_oil']

    df_well['log_Ql'] = np.log(df_well['qnak_liq'])
    df_well['log_Qw'] = np.log(df_well['qnak_water'])
    df_well['log_Qo'] = np.log(df_well['qnak_oil'])

    df_well['Год'] = df_well['Дата'].map(lambda x: x.year)

    q_before_the_last = 0

    if marker == 0:
        if len(df_well['qnak_oil']) > 1:
            q_before_the_last = float(df_well['Добыча нефти за посл.месяц, т'][-2:-1])
        else:
            error = 'имеется только одна точка'
    else:
        if len(df_well['qnak_oil']) > 2:
            df_well = df_well.tail(3)
            q_last = float(df_well['Добыча нефти за посл.месяц, т'][-1:])
            q_before_the_last = float(df_well['Добыча нефти за посл.месяц, т'][-2:-1])
            if q_last / q_before_the_last < 0.25:
                df_well = df_well[:-1]
        else:
            error = 'имеется только одна или две точки'
    
    cumulative_oil_production = df_well['qnak_oil'].values[-1]
    well_operation_time = int(df_well['Год'].tail(1)) - int(df_well['Год'].head(1))

    # статистические методы
    models = []  # list of tuples; (reserves, residual_reserves, korrelation, determination)
    methods = ['Nazarov_Sipachev', 'Sipachev_Pasevich', 'FNI', 'Maksimov', 'Sazonov']
    for name in methods:
        models.append(linear_model(df_well, name))

    # формирование итогового датафрейма
    df_well_result = pd.DataFrame()
    df_well_result['НИЗ'] = [model[0] for model in models]
    df_well_result['ОИЗ'] = [model[1] for model in models]
    df_well_result['Метод'] = methods
    df_well_result['Добыча нефти за посл. мес работы скв., т'] = df_well['Добыча нефти за посл.месяц, т'].values[-1]
    df_well_result['Добыча нефти за предпосл. мес работы скв., т'] = q_before_the_last
    df_well_result['Накопленная добыча нефти, т'] = cumulative_oil_production
    df_well_result['Скважина'] = name_well
    df_well_result['Korrelation'] = [model[2] for model in models]
    df_well_result['Sigma'] = [model[3] for model in models]
    df_well_result['Оставшееся время работы, прогноз, лет'] = \
        df_well_result['ОИЗ'] / (df_well_result['Добыча нефти за посл. мес работы скв., т'] * 12)
    df_well_result['Время работы, прошло, лет'] = well_operation_time
    df_well_result['Координата X'] = float(df_well['Координата забоя Х (по траектории)'][-1:])
    df_well_result['Координата Y'] = float(df_well['Координата забоя Y (по траектории)'][-1:])

    df_well_result = df_well_result.loc[df_well_result['ОИЗ'] > 0]
    if df_well_result.empty:
        error = 'остаточные запасы <= 0'
    
    df_up = df_well_result.loc[df_well_result['Korrelation'] > 0.7]
    df_down = df_well_result.loc[df_well_result['Korrelation'] < (-0.7)]
    df_well_result = pd.concat([df_up, df_down]).reset_index()
    if df_well_result.empty:
        error = 'Корреляция <0.7 или >-0.7'
    
    df_well_result = df_well_result.loc[df_well_result['Оставшееся время работы, прогноз, лет'] < 50]
    if df_well_result.empty:
        error = 'Оставшееся время работы превышает 50 лет'
    
    df_well_result = df_well_result.sort_values('ОИЗ')
    df_well_result = df_well_result.tail(1)

    if not df_well_result.empty:
        if marker == 0:
            df_well_result['Метка'] = 'Расчёт по всем точкам'
        else:
            df_well_result['Метка'] = 'Расчёт по последним 3-м точкам'
    
    return df_well_result, error


def calculate_reserves(
    df: pd.DataFrame,
    min_reserves,
    r_max,
    year_min,
    year_max
):
    wells_set = set(df['№ скважины'])
    df_reserves = pd.DataFrame()
    well_error = []

    for well_name in wells_set:
        print(well_name)
        df_well = df.loc[df['№ скважины'] == well_name].reset_index(drop=True)
        df_result = calculate_reserves_statistics(df_well, well_name)[0]
        if df_result.empty:
            df_result = calculate_reserves_statistics(df_well, well_name, marker=1)[0]
            if df_result.empty:
                well_error.append(well_name)
                continue
        
        # проверка ограничений
        new_oiz = df_result['ОИЗ']
        if df_result['Оставшееся время работы, прогноз, лет'].values[0] > year_max:
            new_oiz = (df_result['Добыча нефти за посл. мес работы скв., т'] + \
                df_result['Добыча нефти за предпосл. мес работы скв., т']) * year_max * 6
        elif df_result['Оставшееся время работы, прогноз, лет'].values[0] < year_min:
            new_oiz = (df_result['Добыча нефти за посл. мес работы скв., т'] + \
                df_result['Добыча нефти за предпосл. мес работы скв., т']) * year_min * 6
        if df_result['ОИЗ'].values[0] < min_reserves:
            new_oiz = min_reserves
        
        df_reserves = pd.concat([df_reserves, df_result])

        df_result['ОИЗ'] = new_oiz
        df_result['Оставшееся время работы, прогноз, лет'] = new_oiz / \
            (df_result['Добыча нефти за посл. мес работы скв., т'] * 12)
    
    # расчёт по карте для скважин с ошибкой

    df_coordinates = df[[
        '№ скважины',
        'Координата забоя Х (по траектории)',
        'Координата забоя Y (по траектории)'
    ]]
    df_coordinates = df_coordinates.drop_duplicates(subset=['№ скважины']).reset_index(drop=True)
    df_all = df_reserves.set_index('Скважина')
    df_coordinates.set_index('№ скважины', inplace=True)
    df_field = pd.merge(
        df_coordinates[[
            'Координата забоя Х (по траектории)',
            'Координата забоя Y (по траектории)'
        ]],
        df_all[['НИЗ']], left_index=True, right_index=True
    )

    df_errors = pd.DataFrame({'Скважина': well_error, '№ скважины': well_error})
    df_errors.set_index('№ скважины', inplace=True)
    df_errors = pd.merge(
        df_coordinates[[
            'Координата забоя Х (по траектории)',
            'Координата забоя Y (по траектории)'
        ]],
        df_errors[['Скважина']], left_index=True, right_index=True
    )

    marker = []
    new_oiz_list = []
    new_oiz_df = []
    new_niz = []

    for well_name in well_error:
        x_er = df_errors['Координата забоя Х (по траектории)'][well_name]
        y_er = df_errors['Координата забоя Y (по траектории)'][well_name]

        distance = ((x_er - df_field['Координата забоя Х (по траектории)']) ** 2 + \
            (y_er - df_field['Координата забоя Y (по траектории)']) ** 2) ** 0.5
        
        r_min = distance.min()
        if r_min > r_max:
            marker.append('! Ближайшая скважина на расстоянии ' + str(r_min))
        else:
            marker.append('Скважина в пределах ограничений')
        
        gur = interpolate_gur(
            x=x_er,
            y=y_er,
            table_x=df_field[['Координата забоя Х (по траектории)']],
            table_y=df_field[['Координата забоя Y (по траектории)']],
            table_z=df_field[['НИЗ']]
        )

        df_well = df.loc[df['№ скважины'] == well_name].reset_index(drop=True)
        df_well['qnak_oil'] = df_well['Добыча нефти за посл.месяц, т'].cumsum()

        cumulative_oil_production = df_well['qnak_oil'].values[-1]

        if len(df_well['qnak_oil']) > 1:
            # добыча нефти за предпоследний месяц
            q_before_the_last = float(df_well['Добыча нефти за посл.месяц, т'][-2:-1])
        else:
            q_before_the_last = 0
        
        q_last = df_well['Добыча нефти за посл.месяц, т'].values[-1]

        for k in gur:
            oiz = k - cumulative_oil_production
            if oiz > 0:
                new_oiz_list.append(oiz)
        
        if len(new_oiz_list) == 0:
            new_oiz = (q_before_the_last + q_last) * year_min * 6
        elif len(new_oiz_list) == 1:
            new_oiz = new_oiz_list[0]
            forecast_residual_operation_time = new_oiz / (q_last * 12)
            if forecast_residual_operation_time > year_max:
                new_oiz = (q_before_the_last + q_last) * year_max * 6
            elif forecast_residual_operation_time  < year_min:
                new_oiz = (q_before_the_last + q_last) * year_min * 6
        else:
            forecast_residual_operation_time_1 = new_oiz_list[0] / (q_last * 12)
            forecast_residual_operation_time_2 = new_oiz_list[1] / (q_last * 12)
            if forecast_residual_operation_time_1 < year_max and forecast_residual_operation_time_1 > year_min:
                new_oiz = new_oiz_list[0]
            else:
                if forecast_residual_operation_time_2 < year_max and forecast_residual_operation_time_2 > year_min:
                    new_oiz = new_oiz_list[1]
                else:
                    if forecast_residual_operation_time_1 > year_max:
                        new_oiz = (q_before_the_last + q_last) * year_max * 6
                    elif forecast_residual_operation_time_1 < year_min:
                        new_oiz = (q_before_the_last + q_last) * year_min * 6
        if new_oiz < min_reserves:
            new_oiz = min_reserves
        new_oiz_df.append(int(new_oiz))
        new_niz.append(int(new_oiz + cumulative_oil_production))
    
    df_errors['НИЗ'] = new_niz
    df_errors['ОИЗ'] = new_oiz_df
    df_errors['Метка'] = marker

    df_all_reserves = pd.concat([df_errors[['Скважина', 'ОИЗ']], df_reserves[['Скважина', 'ОИЗ']]])
    df_all_reserves['ОИЗ'] = df_all_reserves['ОИЗ'] / 1000

    df_reserves = df_reserves.set_index('Скважина')
    df_errors = df_errors.set_index('Скважина')

    with pd.ExcelWriter(os.path.join('data', 'Подсчёт ОИЗ.xlsx')) as writer:
        df_reserves.to_excel(
            writer,
            sheet_name='Расчёт по истории',
            startrow=0,
            startcol=0,
            header=True,
            index=True
        )
        df_errors.to_excel(
            writer,
            sheet_name='Расчёт по карте',
            startrow=0,
            startcol=0,
            header=True,
            index=True
        )
    
    return df_all_reserves


if __name__ == "__main__":
    # чтение данных из файла
    df_initial = pd.read_excel(os.path.join('data', 'Западно-Чистинное-Ю1(2)-МЭР.xlsx'), sheet_name='МЭР')
    # обработка и структурирование данных
    df_initial = history_preprocessing(df_initial, max_delta=365)
    # отображение обработанных и структурированных данных в MS Excel
    xw.view(df_initial)
    oiz = calculate_reserves(df_initial, 2000, 1000, 5, 50).set_index('Скважина').T.to_dict('list')
