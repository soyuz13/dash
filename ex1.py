import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output, State

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

external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

app.layout = html.Div([

    html.Div([
        html.Div([
            dcc.Dropdown(
                id='input-store',
                options=[{'label': x, 'value': x} for x in names],
                placeholder='Выберите магазин'
            )
        ],
            style={'display': 'inline-block', 'width': '15%', 'margin-left': '5px'}),

        html.Div([
            dcc.Dropdown(
                id='input-week',
                options=[
                    {'label': '3', 'value': 3},
                    {'label': '2', 'value': 2},
                    {'label': '1', 'value': 1}
                ],
                placeholder='Выберите неделю'
            )
            ],
            style={'display': 'inline-block', 'width': '15%', 'margin-left': '5px'}),


        html.Div([
            dcc.Dropdown(
                id='input-year',
                options=[
                    {'label': '2020', 'value': 2020},
                    {'label': '2019', 'value': 2019},
                    {'label': '2018', 'value': 2018}
                ],
                placeholder='Выберите год'
            )
            ],
            style={'display': 'inline-block', 'width': '15%', 'margin-left': '5px'}),

        html.Div([
            html.Button(id='submit-button', n_clicks=0, children='Показать')
            ],
            style={'margin-left': '5px'}),
        ]),

    html.Div([
        dcc.Graph(id='chart')
    ])

])

row_ = 3
col_ = 3


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
        return [{'label': x, 'value': x} for x in y2]
    else:
        raise dash.exceptions.PreventUpdate




@app.callback(Output('chart', 'figure'),
              [Input('submit-button', 'n_clicks')],
              [State('input-store', 'value'), State('input-year', 'value'), State('input-week', 'value')])
def update_chart(n_clicks, input1, input2, input3):
    if not n_clicks:
        raise dash.exceptions.PreventUpdate
    else:
        ddf = dfd[input1][(dfd[input1].index.year == input2) & (dfd[input1]['week'] == input3)]
        xa = ddf.index.day
        figm = make_subplots(rows=row_, cols=col_,
                            # column_width=[0.25, 0.3, 0.45],
                            # row_heights=[0.32, 0.04, 0.32, 0.32],
                            specs=[[{"secondary_y": True}, {"secondary_y": True}, {"secondary_y": True}],
                                   [{"secondary_y": True}, {"secondary_y": True}, {"secondary_y": True}],
                                   [{"secondary_y": True}, {"secondary_y": True}, {"secondary_y": True}]],
                            subplot_titles=("Реал", "Вход", "Конв", "Усс", "Акс", "СЧ", 'ТН', 'Кред', 'АДТ'),
                            horizontal_spacing=0.08, vertical_spacing=0.08)

        for r_ in range(row_):
            for c_ in range(col_):
                p = lst[r_][c_] + 'p'
                f = lst[r_][c_] + 'f'
                d = lst[r_][c_] + 'diff%'
                d2 = lst[r_][c_] + 'diff'

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

                t_fmt = [['.4s', '.3s', '%'],
                         ['.3s', '%', '.2s'],
                         ['.3s', '%', '.2s']]

                d_tick = [[250000, 100, None],
                          [10000, 0.02, 2000],
                          [100000, 0.04, 0.005]]

                figm.update_yaxes(secondary_y=False, tickformat=t_fmt[r_][c_], dtick=d_tick[r_][c_], row=r_ + 1, col=c_ + 1)

        figm.update_layout(title_text=f"Неделя {input3}, магазин {input1}",
                           title_font_size=14,
                           template='presentation',
                           showlegend=False,
                           height=700, width=1000
                           )

        return figm


if __name__ == '__main__':
    app.run_server(debug=True)
