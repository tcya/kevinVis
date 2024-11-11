import pandas as pd
import altair as alt
import plotly.express as px
import plotly.graph_objects as go
import shutil


from shiny.express import input, ui, render
from shinywidgets import render_altair, render_plotly

ui.input_file("file", "Choose an Excel to upload:", multiple=True, accept=[".xlsm", ".xls", ".xlsx"]),

df = pd.DataFrame()

def read_uploaded_excel():
    global df
    if not input.file():
        return
    dfs = []
    for file in input.file():
        file_path = file['datapath']
        df = pd.read_excel(file_path, sheet_name='Br RAW Data', header=[3, 4], engine="openpyxl")
        dfs.append(df)
    df = pd.concat(dfs).drop_duplicates(ignore_index=True)
    df.rename(columns=lambda x: '' if 'Unnamed' in x else x, inplace=True)
    df.columns = df.columns.map(' '.join).str.strip()
    df['Sample No'] = df['Sample No'].str.split('-', expand=True).iloc[:, 1]
    df = df[(df['Sample No'].notna()) & (df['Sample No'] != '')]
    df = df.dropna(how='all')
    return df

@render.download()
def download():
    read_uploaded_excel()
    path = '/tmp/merged.xlsx'
    shutil.copy(input.file()[0]['datapath'], '/tmp/merged.xlsm')
    res = construct_output_df()
    res.to_excel(path, sheet_name='Br RAW Data', engine="openpyxl")
    return str(path)

def construct_output_df():
    if not input.file():
        return
    res = []
    three_empty_rows = pd.DataFrame(index=pd.RangeIndex(3), columns=df.columns, dtype='object')
    for s in df['Sample No'].unique():
        if res:
            res.append(three_empty_rows)
        sample_header = pd.DataFrame(index=pd.RangeIndex(1), columns=df.columns, dtype='object')
        sample_header.loc[0, 'Sample No'] = s
        sample_header.loc[0, 'Day'] = 'x'
        res.append(sample_header)
        sample = df[df['Sample No'] == s].copy()
        sample['Sample No'] = 'D' + sample['Day'].astype(str) + '-' + sample['Sample No'].astype(str)
        res.append(sample)
    return pd.concat(res).reset_index(drop=True)
   
def plot_var_by_day(df, var, config=True):
    selection = alt.selection_point(fields=['Sample No'], bind='legend')
    escaped = var
    d = df[['Day', 'Sample No', var]].dropna()
    for c in '.[]':
        escaped = escaped.replace(c, f'\{c}')
    pts = (
        alt.Chart(d)
        .mark_point()
        .encode(
            x=alt.X('Day:O', axis=alt.Axis(labelAngle=-00)),
            y=alt.Y(escaped, title=var),
            color='Sample No:N',
            tooltip=['Sample No', 'Day', escaped],
            opacity=alt.condition(selection, alt.value(1), alt.value(0.1)),
        )
        .add_params(selection).properties(title=var)
    )
    line = (
        alt.Chart(d)
        .mark_line()
        .encode(
            x=alt.X('Day:O', axis=alt.Axis(labelAngle=-00)),
            y=alt.Y(escaped, title=var),
            color='Sample No:N',
            tooltip=['Sample No', 'Day', escaped],
            opacity=alt.condition(selection, alt.value(1), alt.value(0.2)),
        )
        .add_params(selection)
    )
    res = (pts + line).interactive()
    if not config:
        return res
    return res.configure_axis(
        labelFontSize=15,
        titleFontSize=15
    ).configure_legend(
        titleFontSize=15,
        labelFontSize=15
    ).configure_title(fontSize=20)

def create_fn(var):
    def fn():
        return plot_var_by_day(df, var)
    return fn
   
with ui.navset_pill(id="tab"):
    with ui.nav_panel("Raw data"):
        @render.data_frame
        def display_df():
            return read_uploaded_excel()

    vals = [
        'pCO2 [mmHg]',
        'Via. [%]',
        'VCD [106 cells/mL]',
        'Cell Diameter [um]',
        'Gln [mg/L]',
        'Glu [mmol/L]',
        'Gluc. [g/L]',
        'Lact. [g/L]',
        'NH4+ [mmol/L]',
        'Osmo. [mOsm/kg]',
        'Titer [mg/L]',
        'Agitation [rpm]',
        'DO [%]',
        'Temp.  [°C]',
        'pH (int.) [-]',
        'pH Difference [-]',
        'pH (ext.) [-]',
        'pO2 [mmHg]',
        # 'LDH (U/L)',
        # 'Calculated Feed Addition [mL]',
        'pCO2 % [%]',
        'pO2 % [%]',
        'cIVC (106 vc/mL*day)',
        'Specific Productivity (pg/cell/day)',
        'Culture Duration (Days)',
        # 'Base Consumption [mL]',
        # 'log Adjusted VC Growth\nd0-dX µ [day-1]',
        'Doubling Time [hr]',
        'Glucose Consumption [g/L]',
        'Daily Base Consumption [mL/L]',
    ]
    with ui.nav_panel("All"):
        @render_altair
        def all_plots():
            rows = []
            for i, var in enumerate(vals):
                if i % 3== 0:
                    if i != 0:
                        rows.append(res)
                    res = plot_var_by_day(df, var, False)
                else:
                    res = res | plot_var_by_day(df, var, False)
            res = rows[0]
            for r in rows[1:]:
                res &= r
            return res.configure_axis(
                labelFontSize=15,
                titleFontSize=15
            ).configure_legend(
                titleFontSize=15,
                labelFontSize=15
            ).configure_title(fontSize=20)
           
    with ui.nav_panel("3D"):
        ui.input_select('y', 'Y axis', vals)
        ui.input_select('z', 'Z axis', vals)
       
        @render_plotly
        def three_d():
            return px.scatter_3d(df, x='Day', y=input.y(), z=input.z(), color='Sample No', width=800, height=800)
           
    for i, var in enumerate(vals):
        var_base = var.split(' [')[0]
        fn = create_fn(var)
        fn.__name__ = f'fn{i}'
        with ui.nav_panel(var_base):
            render_altair(fn)