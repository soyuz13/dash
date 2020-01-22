import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output

import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import datetime


def pars(x):
    try:
        g = [int(m) for m in x.split('.')]
        return datetime.date(*g[::-1])
    except:
        return pd.NaT


df2 = pd.read_excel('BI.xlsm', sheet_name='Лист1',
                    usecols=[0] + [x for x in range(3, 24)],
                    skiprows=4, index_col=1, parse_dates=[1], date_parser=pars)

df2.columns = ['store', 'r_p', 'r_f', 'im_f', 'uss_p', 'uss_f', 'vh_p', 'vh_f', 'prod_f', 'ch_f', 'konv_p',
               'aks_p', 'aks_f', 'sch_p', 'sch_f', 'adt_p', 'adt_f', 'kred_p', 'kred_f', 'tn_p', 'tn_f']

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

lst = (('r_', 'vh_', 'konv_'),
       ('uss_', 'aks_', 'sch_'),
       ('tn_', 'kred_', 'adt_'))

lst_ = tuple([y for x in lst for y in x])

for k in dfd:
    for i in dfd[k].columns:
        dfd[k][i+'_D-7'] = dfd[k][i].shift(7)
    for i in lst_:
        dfd[k][i+'fp'] = dfd[k][i+'f']/dfd[k][i+'p']
        dfd[k][i+'diff%'] = (dfd[k][i+'f']-dfd[k][i+'f_D-7'])/dfd[k][i+'f_D-7']
        dfd[k][i+'diff'] = dfd[k][i+'f']-dfd[k][i+'f_D-7']

external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

app.layout = html.Div([
    html.Div([
        html.Div([
            html.Label('Магазин'),
            dcc.Dropdown(
                id='input-store',
                options=[{'label': x, 'value': x} for x in names],
                placeholder='...'
            )
        ],
            style={'display': 'inline-block', 'width': '15%', 'margin-left': '5px'}),

        html.Div([
            html.Label('Отчетный год'),
            dcc.Dropdown(
                id='input-year',
                placeholder='...'
            )
        ],
            style={'display': 'inline-block', 'width': '15%', 'margin-left': '5px'}),

        html.Div([
            html.Label('Отчетная неделя'),
            dcc.Dropdown(
                id='input-week',
                placeholder='...'
            )
        ],
            style={'display': 'inline-block', 'width': '15%', 'margin-left': '5px'})
    ]),

    html.Div([
        dcc.Graph(id='chart')
    ],
        style={'margin-left': '0px', 'float': 'left'}
    )

])

row_ = 3
col_ = 9


@app.callback(Output('input-year', 'options'),
              [Input('input-store', 'value')])
def update_year(value):
    if value:
        y1 = dfd[value].index.year.unique().tolist()
        y1.sort(reverse=True)
        return [{'label': x, 'value': x} for x in y1]
    else:
        raise dash.exceptions.PreventUpdate


@app.callback(Output('input-week', 'options'),
              [Input('input-year', 'value'),
               Input('input-store', 'value')])
def update_week(val1, val2):
    if val1 and val2:
        y2 = dfd[val2][str(val1)]['week'].unique().tolist()
        y2.sort(reverse=True)
        # ls = dfd[val2][str(val1)]['week'].resample('W', closed='right').first().index
        # print(ls)
        return [{'label': x, 'value': x} for x in y2]
    else:
        raise dash.exceptions.PreventUpdate


@app.callback(Output('chart', 'figure'),
              #[Input('submit-button', 'n_clicks')],
              [Input('input-store', 'value'), Input('input-year', 'value'), Input('input-week', 'value')])
def update_chart(input1, input2, input3):
    if not (input3 and input2 and input1):
        figm = go.Figure()
        return figm
        raise dash.exceptions.PreventUpdate

    else:
        ddf = dfd[input1][(dfd[input1].index.year == input2) & (dfd[input1]['week'] == input3)]
        xa = ddf.index.day
        figm = make_subplots(rows=row_, cols=col_, column_width=[0.3, 0.025, 0.00, 0.3, 0.025, 0.00, 0.31, 0.025, 0.00],
                             specs=[[{"secondary_y": True}, {}, {}, {"secondary_y": True}, {}, {}, {"secondary_y": True}, {}, {}],
                                   [{"secondary_y": True}, {}, {}, {"secondary_y": True}, {}, {}, {"secondary_y": True}, {}, {}],
                                   [{"secondary_y": True}, {}, {}, {"secondary_y": True}, {}, {}, {"secondary_y": True}, {}, {}]],
                             subplot_titles=("Реал", '%', '', "Вход", '%', '', "Конв", '%', '',
                                            "Усс", '%', '', "Акс", '%', '', "СЧ", '%', '',
                                            'ТН', '%', '', 'Кред', '%', '', 'АДТ', '%', ''),
                             horizontal_spacing=0.035, vertical_spacing=0.1)

        print(ddf.columns)

        for ro in range(1, row_+1):
            for co in range(1, col_+1, 3):
                co2 = int((co-1)/3)
                p = lst[ro-1][co2] + 'p'
                f = lst[ro-1][co2] + 'f'
                d = lst[ro-1][co2] + 'diff%'
                h = lst[ro-1][co2] + 'diff'

                m = max(ddf[p].max(), ddf[f].max()) * 1.1

                trace_plan = go.Scatter(x=xa, y=ddf[p], name="План", mode='lines', marker_color='red')
                trace_fact = go.Scatter(x=xa, y=ddf[f], name="Факт", mode='lines', marker_color='green')
                trace_diff = go.Bar(x=xa, y=ddf[d], hovertext=ddf[h], hovertemplate='%{y:.2%}, %{hovertext:.3s}',
                                    textposition='auto', text=ddf[d], texttemplate="%{y:%}", name="Откл.",
                                    marker_color='LightBlue', opacity=0.5)

                figm.add_traces([trace_plan, trace_fact, trace_diff], secondary_ys=[False, False, True],
                                rows=[ro]*3, cols=[co]*3)

                t_fmt = [['.4s', '.3s', '%'],
                         ['.3s', '%', '.2s'],
                         ['.3s', '%', '.2s']]

                d_tick = [[250000, 100, None],
                          [10000, 0.02, 2000],
                          [100000, 0.04, 0.005]]

                figm.update_yaxes(range=[0, m], tickformat=t_fmt[ro-1][co2], dtick=d_tick[ro-1][co2],
                                  row=ro, col=co, tickfont=dict(size=9))
                figm.update_yaxes(secondary_y=True, range=[-0.75, 0.75], row=ro, col=co,
                                  dtick=0.25, tickfont=dict(size=9), tickformat=' >3%',
                                  zeroline=True, zerolinewidth=2, zerolinecolor='LightBlue', showgrid=False,
                                  linecolor='LightBlue', gridcolor='LightBlue')
                figm.update_xaxes(nticks=len(ddf) + 1, tickangle=0, showgrid=True, tickfont=dict(size=9))

        for ro in range(1, row_+1):
            for co in range(2, col_+1, 3):
                y1 = 1
                kp = lst[ro-1][int((co-2)/3)]
                y2 = load_cw(input1, kp)

                trace_bar_plan = go.Bar(name='План', y=[y1], marker_color='darkgray', width=20)
                trace_bar_fact = go.Bar(name='Факт', y=[y2], marker_color='lightgreen', width=12)
                trace_bar_diff = go.Bar(name='Факт-1', y=[0], marker_color='blue', width=6)

                figm.add_traces([trace_bar_plan, trace_bar_fact, trace_bar_diff], rows=[ro]*3, cols=[co]*3)

                figm.update_xaxes(row=ro, col=co, visible=False)
                figm.update_yaxes(row=ro, col=co, visible=False)

        figm.update_layout(title_text=f"Неделя {input3}, магазин {input1}", title_font_size=14,
                           template='presentation', showlegend=False, height=700, width=1200)

        return figm


def load_cw(store, kpi):
    columns = ['store', 'day', 'r_p', 'r_f', 'im_f', 'uss_p', 'uss_f', 'vh_p', 'vh_f', 'prod_f', 'ch_f', 'konv_p',
               'aks_p', 'aks_f', 'sch_p', 'sch_f', 'adt_p', 'adt_f', 'kred_p', 'kred_f', 'tn_p', 'tn_f']

    dic_week = {}
    sheets = ['curr_week', 'prev_week', 'prev_week_year']
    names = ['df_cw', 'df_pw', 'df_py']

    p1 = ['650/000', '640/000', '6502']
    p2 = ['650', '640', '6502']
    p3 = dict(zip(p1, p2))

    lst = (('r_', 'vh_', 'konv_'),
           ('uss_', 'aks_', 'sch_'),
           ('tn_', 'kred_', 'adt_'))

    lst_ = tuple([y for x in lst for y in x])

    for i, n in enumerate(sheets):
        m = names[i]
        dic_week[m] = pd.read_excel('curr_week.xlsm', sheet_name=i, usecols=[x for x in range(22)], skiprows=4)
        dic_week[m].columns = columns
        dic_week[m] = dic_week[m][dic_week[m]['store'].str.endswith('Итог')].drop(columns='day')
        dic_week[m].index = dic_week[m].store.str.split(':', expand=True)[0].str.strip().apply(lambda x: p3[x])
        dic_week[m] = dic_week[m].drop(columns='store')
        dic_week[m]['konv_f'] = dic_week[m]['ch_f'] / dic_week[m]['vh_f']

        for k in lst_:
            dic_week[m][k + 'fp'] = dic_week[m][k + 'f'] / dic_week[m][k + 'p']

    return dic_week['df_cw'].loc[store][kpi+'fp']


if __name__ == '__main__':
    app.run_server(debug=True)
