import os
import shutil
import pandas as pd
import datetime


class LoadBIData:
    def __init__(self):
        self.col_lst = ['store', 'day', 'r_p', 'r_f', 'im_f', 'uss_p', 'uss_f', 'vh_p',
               'vh_f', 'prod_f', 'ch_f','konv_p', 'aks_p', 'aks_f', 'sch_p', 'sch_f',
               'adt_p', 'adt_f', 'kred_p', 'kred_f', 'tn_p', 'tn_f', 'empl_f']
        self.store_lst = ['650', '640', '6502']
        self._store_bi = ['650/000', '640/000', '6502']

        self._files = {'curr_week': 'curr_week.xlsm',
                       'curr_month': 'curr_month.xlsm',
                       'week': 'weeks.xlsm',
                       'months': 'months.xlsm'}

        if os.name == 'posix':
            srv = '/Volumes'
        else:
            srv = r'\\192.168.181.63'
        self.path = os.path.join(srv, 'Public', 'Marchenko_VV')

    def _pars(self, x):
        try:
            g = [int(m) for m in x.split('.')]
            return datetime.date(*g[::-1])
        except:
            return pd.NaT

    def _create_df(self, fil, sheet):
        df = pd.read_excel(fil, sheet_name=sheet, usecols=[x for x in range(23)],
                           skiprows=4, parse_dates=[1], date_parser=self._pars)
        df.columns = self.col_lst
        df.set_index('day', drop=True, inplace=True)
        df['konv_f'] = df['ch_f'] / df['vh_f']
        d = df.columns.to_list()
        for i in d:
            kpi = i[:-1]
            if i[-1] == 'p':
                try:
                    df[kpi + 'fp'] = df[kpi + 'f'] / df[kpi + 'p']
                except:
                    pass
        return df


    def load_files(self, ls=None):
        lst = ls if ls else self._files
        for i in lst:
            print('start copy ' + i)
            shutil.copy2(os.path.join(self.path, i), i)
            print('finish')

    def get_curr_week(self, store=None, results=True):

        lst = []
        sheets = ['curr_week', 'prev_week', 'prev_week_year']
        for i in sheets:
            df = self._create_df(self._files['curr_week'], i)
            if store:
                if results:
                    df = df[df.store.str.startswith(store) & df.store.str.endswith('Итог')]
                else:
                    df = df[df.store.str.startswith(store) & ~df.store.str.endswith('Итог')]
            else:
                df = df[df.store == 'Общий итог']
            lst.append(df)
        return lst


    def get_curr_month(self):
        pass

    def get_month(self):
        pass

    def get_week(self):
        pass


a = LoadBIData()
# a.load_files(['curr_week.xlsm', 'curr_month.xlsm'])

print(a.get_curr_week('650/000', False))




