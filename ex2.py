import plotly.graph_objects as go
from plotly.subplots import make_subplots
import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output

import datetime
import loader


names = ('650/000', '640/000', '6502')
getdata = loader.LoadBIData()
#getdata.load_files()


lst = (('r_', 'tn_', 'sch_' ,  'uss_'),
       ('vh_', 'kred_', 'ch_', 'konvss_'),
       ('konv_', 'im_', 'aks_', 'adt_'))

lst_ = tuple([y for x in lst for y in x])

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
col_ = 8


@app.callback(Output('input-year', 'options'),
              [Input('input-store', 'value')])
def update_year(value):
    if value:
        y1 = [2020, 2019]  # dfd[value].index.year.unique().tolist()
        y1.sort(reverse=True)
        return [{'label': x, 'value': x} for x in y1]
    else:
        raise dash.exceptions.PreventUpdate


@app.callback(Output('input-week', 'options'),
              [Input('input-year', 'value'),
               Input('input-store', 'value')])
def update_week(val1, val2):
    if val1 and val2:
        y2 = [5, 4, 3, 2, 1]  # dfd[val2][str(val1)]['week'].unique().tolist()
        y2.sort(reverse=True)

        return [{'label': x, 'value': x} for x in y2]
    else:
        raise dash.exceptions.PreventUpdate


@app.callback(Output('chart', 'figure'),
              [Input('input-store', 'value'), Input('input-year', 'value'), Input('input-week', 'value')])
def update_chart(inp_store, inp_year, inp_week):
    if not (inp_week and inp_year and inp_store):
        figm = go.Figure()
        return figm
        raise dash.exceptions.PreventUpdate

    else:
        current_week = False
        dtn = datetime.datetime.now()
        if (dtn.year, int(dtn.strftime('%V'))) == (inp_year, inp_week):
            current_week = True
        # current_week = True

        df_cw, df_pw, df_py = getdata.get_curr_week(inp_store, results=False)

        lbl_x_axis = df_cw.index.day.tolist()
        empl_cw = df_cw['empl_f']
        empl_pw = df_pw['empl_f']

        x_txt = []
        for i, k in enumerate(lbl_x_axis):
            x_txt.append(str(k) + '<br>' + str(int(empl_cw[i])) + '/' + str(int(empl_pw[i])))
        lbl_x_axis = x_txt

        figm = make_subplots(rows=row_, cols=col_, column_width=[0.23, 0.02, 0.23, 0.02, 0.23, 0.02, 0.23, 0.02],
                             specs=[[{"secondary_y": True}, {}, {"secondary_y": True}, {}, {"secondary_y": True}, {}, {"secondary_y": True}, {}],
                                    [{"secondary_y": True}, {}, {"secondary_y": True}, {}, {"secondary_y": True}, {}, {"secondary_y": True}, {}],
                                    [{"secondary_y": True}, {}, {"secondary_y": True}, {}, {"secondary_y": True}, {}, {"secondary_y": True}, {}]],
                             subplot_titles=('Реал', '1', 'ТН', '2', 'СЧ', '3', 'Усс', '33',
                                             'Вход', '4', 'Кред', '5', 'СП', '6', 'Конв УСС', '66',
                                             'Конв', '7', 'ИМ', '8', 'Акс', '9', 'АДТ', '66'),
                             horizontal_spacing=0.04, vertical_spacing=0.15)

        for ro in range(1, row_+1):
            for co in range(1, col_+1, 2):
                co2 = int((co-1)/2)
                p = lst[ro-1][co2] + 'p'
                f = lst[ro-1][co2] + 'f'

                df_pw.index = df_cw[f].index
                d = (df_cw[f] - df_pw[f])/df_pw[f]
                h = df_cw[f] - df_pw[f]

                m = max(df_cw[p].max(), df_cw[f].max()) * 1.1

                trace_plan = go.Scatter(x=lbl_x_axis, y=df_cw[p], name="План", mode='lines', marker_color='red')
                trace_fact = go.Scatter(x=lbl_x_axis, y=df_cw[f], name="Факт", mode='lines+markers', marker_color='green')
                trace_diff = go.Bar(x=lbl_x_axis, y=d, hovertext=h, hovertemplate='%{y:.2%}, %{hovertext:.3s}',
                                    textposition='auto', text=d, texttemplate="%{y:%}", name="Откл.",
                                    marker_color='LightBlue', opacity=0.5)

                figm.add_traces([trace_plan, trace_fact, trace_diff], secondary_ys=[False, False, True],
                                rows=[ro]*3, cols=[co]*3)

                t_fmt = [['.4s', '.3s', '.2s', '.2s'],
                         ['.3s', '%.2s', '.2s', '%'],
                         ['%', '.3s', '%', '%']]

                d_tick = [[250000, 50000, 2000, 4000],
                          [100, 0.04, 20, 0.1],
                          [0.05, 20000, 0.02, 0.1]]

                figm.update_yaxes(range=[0, m], tickformat=t_fmt[ro-1][co2], dtick=d_tick[ro-1][co2],
                                  row=ro, col=co, tickfont=dict(size=9))
                figm.update_yaxes(secondary_y=True, range=[-0.75, 0.75], row=ro, col=co,
                                  dtick=0.25, tickfont=dict(size=9), tickformat=' >3%',
                                  zeroline=True, zerolinewidth=2, zerolinecolor='LightBlue', showgrid=False,
                                  linecolor='LightBlue', gridcolor='LightBlue')
                figm.update_xaxes(nticks=len(df_cw) + 1, tickangle=0, showgrid=True, tickfont=dict(size=9))

        for ro in range(1, row_+1):
            for co in range(2, col_+1, 2):
                kp = lst[ro-1][int((co-2)/2)]

                if current_week:
                    y_cw, y_pw, y_py = getdata.get_curr_week(inp_store, results=True, kpi=kp+'f')
                    y4 = getdata.get_curr_week(inp_store, results=True, kpi=kp+'fp')[0]
                    diff = y_cw - y_pw

                else:
                    y_pw = 0.7
                    y_cw = 0.6
                    y_py = 0.8

                # TODO: создать матрицу с форматами подписей
                # trace_bar_plan = go.Bar(name='План', y=[0], width=0, marker_color='#7FFF00') # , marker_color='lightgray'
                trace_bar_prev = go.Bar(name='Факт НЕД-1', x=[inp_week], y=[y_pw], width=12, marker_color='#FF8800') # , marker_color='papayawhip'
                trace_bar_curr = go.Bar(name='Факт', x=[inp_week], y=[y_cw], width=6, marker_color='green',
                                        hovertext=[str(diff)], hovertemplate='%{y:f}<br>%{hovertext:+,.2f}') # , marker_color='green'

                figm.add_traces([trace_bar_prev, trace_bar_curr], rows=[ro]*2, cols=[co]*2)

                figm.update_xaxes(row=ro, col=co, visible=False)
                figm.update_yaxes(row=ro, col=co, visible=False)

                map_ = [[1, 3, 5, 7],
                        [9, 11, 13, 15],
                        [15, 17, 19, 21]]
                pos = map_[ro-1][int(co/2)-1]
                figm['layout']['annotations'][pos]['text'] = f'{y4*100:.1f}%'
                # figm['layout']['annotations'][pos]['bgcolor'] = 'rgb(255, 100, 0, 0)'

        figm.update_layout(title_text=f"Неделя {inp_week}, магазин {inp_store}", title_font_size=14,
                           template='presentation', showlegend=False, height=700, width=1280)

        #print(figm)
        return figm


if __name__ == '__main__':
    app.run_server(debug=True)
