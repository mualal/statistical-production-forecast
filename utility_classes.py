import numpy as np
import calendar
from datetime import date, timedelta


class FluidProduction:

    def __init__(
        self,
        day_fluid_production,
        considerations,
        well_name
    ):
        self.day_fluid_production = day_fluid_production
        self.considerations = considerations
        self.well_name = well_name
        self.first_month = -1
        self.start_q = -1
        self.ind_max = -1
    
    def adaptation(
        self,
        correlation_coeffs
    ):
        k1, k2 = correlation_coeffs
        max_day_prod = np.amax(self.day_fluid_production)
        index = list(np.where(self.day_fluid_production) == max_day_prod)[0][0]
        if index != (self.day_fluid_production.size - 1) and \
            index > (self.day_fluid_production.size - 4) and \
                self.day_fluid_production.size > 3:
                max_day_prod = np.amax(self.day_fluid_production[:-3])
                index = list(np.where(self.day_fluid_production == np.amax(self.day_fluid_production[0:-3])))[0][0]
        
        indexes = np.arange(start=index, stop=self.day_fluid_production.size, step=1) - index
        day_fluid_production_month = max_day_prod * (1 + k1 * k2 * indexes) ** (-1 / k2)
        deviation = [(self.day_fluid_production[index:] - day_fluid_production_month) ** 2]
        self.first_month = self.day_fluid_production.size - index + 1
        self.start_q = max_day_prod
        self.ind_max = index
        return np.sum(deviation)
    
    def to_conditions(
        self,
        correlation_coeffs
    ):
        k1, k2 = correlation_coeffs
        global base_correction
        point = self.considerations[self.well_name][1]
        if np.isnan(point):
            point = 1
        if point == 1:
            base_correction = self.day_fluid_production[-1]
        elif point == 3:
            if self.day_fluid_production.size >= 3:
                base_correction = np.average(self.day_fluid_production[-3:-1])
            elif self.day_fluid_production.size == 2:
                base_correction = np.average(self.day_fluid_production[-2:-1])
            else:
                base_correction = self.day_fluid_production[-1]
        else:
            print('Неверный формат для условия привязки! Привязка будет осуществляться к последней точке.')
            base_correction = self.day_fluid_production[-1]
        
        max_day_prod = np.amax(self.day_fluid_production)
        index = list(np.where(self.day_fluid_production == max_day_prod))[0][0]

        if index > (self.day_fluid_production.size - 4) and self.day_fluid_production.size > 3:
            max_day_prod = np.amax(self.day_fluid_production[:-3])
            index = list(np.where(self.day_fluid_production == np.amax(self.day_fluid_production[0:-3])))[0][0]
        
        last_prod = max_day_prod * (1 + k1 * k2 * (self.day_fluid_production.size - 1 - index)) ** (-1 / k2)
        binding = base_correction - last_prod
        return binding


class DesaturationCharacteristic:

    def __init__(
        self,
        oil_production,
        liq_production,
        irr,
        considerations,
        well_name,
        mark,
        wc_fact,
        rf_now
    ):
        self.oil_production = oil_production
        self.liq_production = liq_production
        self.irr = irr
        self.considerations = considerations
        self.well_name = well_name
        self.mark = mark
        self.wc_fact = wc_fact
        self.rf_now = rf_now
    
    def solver(
        self,
        correlation_coeffs
    ):
        corey_oil, corey_water, mef = correlation_coeffs

        v1 = np.cumsum(self.oil_production) / self.irr / 1e3
        v1 = np.delete(v1, 0)
        v1 = np.insert(v1, 0, 0)
        k1 = (1 - v1) ** corey_oil / ((1 - v1) ** corey_oil + mef * v1 * corey_water)

        v2 = v1 + self.liq_production * k1 / 2 / self.irr / 1e3
        k2 = (1 - v2) ** corey_oil / ((1 - v2) ** corey_oil + mef * v2 * corey_water)

        v3 = v1 + self.liq_production * k2 / 2 / self.irr / 1e3
        k3 = (1 - v3) ** corey_oil / ((1 - v3) ** corey_oil + mef * v3 * corey_water)

        v4 = v1 + self.liq_production * k3 / 2 / self.irr / 1e3
        k4 = (1 - v4) ** corey_oil / ((1 - v4) ** corey_oil + mef * v4 * corey_water)

        oil_production_model = self.liq_production / 6 * (k1 + 2 * k2 + 2 * k3 + k4)
        oil_production_model[oil_production_model == -np.inf] = 0
        oil_production_model[oil_production_model == np.inf] = 0

        deviation = [(oil_production_model - self.oil_production) ** 2]

        return np.sum(deviation)
    
    def to_conditions(
        self,
        correlation_coeffs
    ):
        corey_oil, corey_water, mef = correlation_coeffs
        point = self.considerations[self.well_name][0]

        if np.isnan(point) or self.mark == False:
            point = 1
        if point == 1:
            if self.wc_fact.size > 1:
                wc_last = self.wc_fact[-1]
            else:
                wc_last = self.wc_fact
        elif point == 3:
            if self.wc_fact.size >= 3:
                wc_last = np.average(self.wc_fact[-3:-1])
            elif self.wc_fact.size == 2:
                wc_last = np.average(self.wc_fact[-2:-1])
            else:
                wc_last = self.wc_fact
        else:
            print('Неверный формат для условия привязки! Привязка будет осуществляться к последней точке')
            if self.wc_fact.size > 1:
                wc_last = self.wc_fact[-1]
            else:
                wc_last = self.wc_fact
        
        wc_model = mef * self.rf_now ** corey_water / \
            ((1 - self.rf_now) ** corey_oil + mef * self.rf_now ** corey_water)
        binding = wc_model - wc_last

        return binding


def fluid_production_profile(
    period,
    desaturation_characteristic,
    liq_production,
    date_last,
    now_rf,
    irr
):
    c_oil, c_water, mef = desaturation_characteristic
    k1, k2, num_m, q_start = liq_production
    k1 = float(k1)
    k2 = float(k2)
    num_m = int(float(num_m))
    q_start = float(q_start)
    q_n_t = [0]
    q_n = []
    wc_model = []
    q_liq = []
    date_last = date(date_last[0], date_last[1], date_last[2])

    for _ in range(period):
        now_rf = now_rf + q_n_t[-1] / irr / 1e3
        if now_rf >= 1:
            now_rf = 0.99999999999
        wc_model.append(mef * now_rf ** c_water / ((1 - now_rf) ** c_oil + mef * now_rf ** c_water))
        q_liq.append(q_start * (1 + k1 * k2 * (num_m - 1)) ** (-1 / k2))
        num_m += 1
        q_n.append(q_liq[-1] * (1 - wc_model[-1]))
        days_in_month = calendar.monthrange(date_last.year, date_last.month)[1]
        q_n_t.append(q_n[-1] * days_in_month)
        date_last += timedelta(days=days_in_month)
    
    return q_n, q_liq
