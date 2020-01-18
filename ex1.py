import dash
import dash_core_components as dcc
import dash_html_components as html

import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import datetime


def pars(x):
    try:
        lst = [int(m) for m in x.split('.')]
        return datetime.date(*lst[::-1])
    except:
        return pd.NaT


df2 = pd.read_excel('BI.xlsm', sheet_name='Лист1',
                    usecols=[0, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23],
                    skiprows=4, index_col=1, parse_dates=[1], date_parser=pars)

df2.columns = ['store','r_p', 'r_f', 'im_f', 'uss_p', 'uss_f', 'vh_p',
               'vh_f', 'prod_f', 'ch_f', 'konv_p', 'aks_p', 'aks_f', 'sch_p', 'sch_f',
              'adt_p', 'adt_f', 'kred_p', 'kred_f', 'tn_p', 'tn_f']

df2.index.name = 'day'
df2 = df2[df2.index.notnull()]

df2['konv_f'] = df2['ch_f']/df2['vh_f']
df2['week'] = df2.index.weekofyear

codes1 = ['650/', '640/', '6502']
codes2 = codes1[:]
codes3 = codes1[:]
names = ('650', '640', '6502')

for i, j in enumerate(codes1):
    codes1[i] = df2[df2.store.str.startswith(j)].drop(columns='store')
    codes2[i] = codes1[i].resample('w', closed='right').sum()
    # codes2[i]['week'] = codes2[i].index.weekofyear
    codes3[i] = codes1[i].resample('M', closed='right').sum()

dfd = dict(zip(names, codes1))  # df по дням
dfw = dict(zip(names, codes2))  # df по неделям
dfm = dict(zip(names, codes3))  # df по месяцам

lst = (('r_', 'vh_', 'konv_'), ('uss_', 'aks_', 'sch_'), ('tn_', 'kred_', 'adt_'))
lst_ = tuple([y for x in lst for y in x])

for k in dfd:
    for i in dfd[k].columns:
        dfd[k][i+'_D-7'] = dfd[k][i].shift(7)
    for i in lst_:
        dfd[k][i+'fp'] = dfd[k][i+'f']/dfd[k][i+'p']
        dfd[k][i+'diff%'] = (dfd[k][i+'f']-dfd[k][i+'f_D-7'])/dfd[k][i+'f_D-7']
        dfd[k][i+'diff'] = dfd[k][i+'f']-dfd[k][i+'f_D-7']

year = '2020'
week = 3
row_ = 3
col_ = 3

for k in dfd:
    figm = make_subplots(rows=row_, cols=col_,
                         # column_width=[0.25, 0.3, 0.45],
                         # row_heights=[0.32, 0.04, 0.32, 0.32],
                         specs=[[{"secondary_y": True}, {"secondary_y": True}, {"secondary_y": True}],
                                # [{"type": 'indicator'}, {"type": 'indicator'}, {"type": 'indicator'}],
                                [{"secondary_y": True}, {"secondary_y": True}, {"secondary_y": True}],
                                [{"secondary_y": True}, {"secondary_y": True}, {"secondary_y": True}]],
                         subplot_titles=("Реал", "Вход", "Конв", "Усс", "Акс", "СЧ", 'ТН', 'Кред', 'АДТ'),
                         horizontal_spacing=0.08, vertical_spacing=0.08)
    ddf = dfd[k][(dfd[k].index >= year) & (dfd[k]['week'] == week)]
    xa = ddf.index.day

    for r_ in range(row_):
        for c_ in range(col_):
            p = lst[int(r_ / 2)][c_] + 'p'
            f = lst[int(r_ / 2)][c_] + 'f'
            d = lst[int(r_ / 2)][c_] + 'diff%'
            d2 = lst[int(r_ / 2)][c_] + 'diff'


            m = max(ddf[p].max(), ddf[f].max()) * 1.1
            figm.add_trace(go.Scatter(x=xa, y=ddf[p], name="План", mode='lines', marker_color='red'), row=r_ + 1,
                           col=c_ + 1)
            figm.add_trace(go.Scatter(x=xa, y=ddf[f], name="Факт", mode='lines', marker_color='green'), row=r_ + 1,
                           col=c_ + 1)
            figm.add_trace(go.Bar(x=xa,
                                    y=ddf[d],
                                    # hoverinfo='x+text',
                                    hovertext=ddf[d2],
                                    hovertemplate='%{y:.2%}, %{hovertext:.3s}',
                                    textposition='auto',
                                    text=ddf[d],
                                    texttemplate="%{y:%}",
                                    marker_color='LightBlue',
                                    name="Откл.",
                                    opacity=0.5),
                            secondary_y=True,
                            row=r_ + 1, col=c_ + 1)
            figm.update_yaxes(range=[0, m], row=r_ + 1, col=c_ + 1, tickfont=dict(size=9), tickformat='s')
            figm.update_yaxes(secondary_y=True, range=[-0.75, 0.75],
                                dtick=0.25, tickfont=dict(size=9), tickformat=' >3%',
                                zeroline=True, zerolinewidth=2, zerolinecolor='LightBlue',
                                linecolor='LightBlue', gridcolor='LightBlue',
                                row=r_ + 1, col=c_ + 1)
            figm.update_xaxes(nticks=len(ddf) + 1, tickangle=0, showgrid=True, tickfont=dict(size=9))

            if (c_ == 0 and r_ == 0):
                figm.update_yaxes(secondary_y=False, tickformat='.4s', dtick=250000, row=r_ + 1, col=c_ + 1)
            if (c_ == 1 and r_ == 0):
                figm.update_yaxes(secondary_y=False, tickformat='.3s', dtick=100, row=r_ + 1, col=c_ + 1)
            if (c_ == 2 and r_ == 0):
                figm.update_yaxes(secondary_y=False, tickformat='%', row=r_ + 1, col=c_ + 1)

            if (c_ == 0 and r_ == 2):
                figm.update_yaxes(secondary_y=False, tickformat='.3s', dtick=10000, row=r_ + 1, col=c_ + 1)
            if (c_ == 1 and r_ == 2):
                figm.update_yaxes(secondary_y=False, tickformat='%', dtick=0.02, row=r_ + 1, col=c_ + 1)
            if (c_ == 2 and r_ == 2):
                figm.update_yaxes(secondary_y=False, tickformat='.2s', dtick=2000, row=r_ + 1, col=c_ + 1)

            if (c_ == 0 and r_ == 3):
                figm.update_yaxes(secondary_y=False, tickformat='.3s', dtick=100000, row=r_ + 1, col=c_ + 1)
            if (c_ == 1 and r_ == 3):
                figm.update_yaxes(secondary_y=False, tickformat='%', dtick=0.04, row=r_ + 1, col=c_ + 1)
            if (c_ == 2 and r_ == 3):
                figm.update_yaxes(secondary_y=False, tickformat='.2s', dtick=0.005, row=r_ + 1, col=c_ + 1)

    figm.update_layout(title_text=f"Неделя {week}, магазин {k}",
                       title_font_size=14,
                       template='presentation',
                       showlegend=False,
                       #height = 800, width = 1000
                       )





'''figm = make_subplots(rows=2, cols=2,
                     specs=[[{"secondary_y": True}, {"secondary_y": True}],
                            [{"secondary_y": True}, {"secondary_y": True}]],
                     subplot_titles=("Реал", "Вход", "Усс", "Акс", 'ТН', 'Кред'),
                     horizontal_spacing=0.08, vertical_spacing=0.2)
figm.add_trace(go.Scatter(x=[1,2,3,4,5], y=[2,6,4,8,3], name="План", mode='lines', marker_color='red'), row=1, col=1)
figm.add_trace(go.Scatter(x=[1,2,3,4,5], y=[3,4,3,4,5], name="План", mode='lines', marker_color='red'), row=1, col=2)
figm.add_trace(go.Scatter(x=[1,2,3,4,5], y=[4,6,4,7,6], name="План", mode='lines', marker_color='red'), row=2, col=1)
figm.add_trace(go.Scatter(x=[1,2,3,4,5], y=[3,5,7,3,4], name="План", mode='lines', marker_color='red'), row=2, col=2)'''




external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

app.layout = html.Div([
    html.Label('Выберите магазин'),
    dcc.Dropdown(
        options=[{'label': x, 'value': x} for x in names]
    ),

    html.Label('Выберите неделю'),
    dcc.Dropdown(
        options=[
            {'label': '3', 'value': 3},
            {'label': '2', 'value': 2},
            {'label': '1', 'value': 1}
        ],
        value=3
    ),
    html.Label('Выберите год'),
    dcc.Dropdown(
        options=[
            {'label': '2020', 'value': 2020},
            {'label': '2019', 'value': 2019},
            {'label': '2018', 'value': 2018}
        ],
        value=2020
    ),
    dcc.Graph(figure=figm)
],
style={'width': '50%'})

if __name__ == '__main__':
    app.run_server(debug=True)