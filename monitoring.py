import streamlit as st
import pandas as pd
import datetime as dt
import math
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
from PIL import Image
import matplotlib.pyplot as plt
from colour import Color

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

image = Image.open('invoke_logo.jpg')
st.sidebar.title('Call Centre Performance Tracker 2.0')
st.sidebar.image(image)
option1 = st.sidebar.selectbox('Select option', ('Daily', 'Day-to-Day'))

def clean_names(df, col):
    df[col] = [x.replace('@invokeisdata.com', '') for x in df[col]]
    df[col] = [x.replace('hudahusna', 'huda') for x in df[col]]
    df[col] = [x.replace('amishaa', 'amisha') for x in df[col]]
    df[col] = [x.replace('athiyah', 'tiyah') for x in df[col]]
    df[col] = [x.replace('atiqahliyana', 'atiqah') for x in df[col]]
    return df

def color_kpi(val):
    color = 'red' if val=='X' else 'green'
    return f'background-color: {color}'

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Data')
    workbook = writer.book
    worksheet = writer.sheets['Data']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.save()
    processed_data = output.getvalue()
    return processed_data

if option1 == 'Daily':
    st.image('cc-point-system.png')
    number = st.number_input('Select number of campaign(s)', min_value = 1)
    dfs = []
    for i in range(number):
        st.header('SURVEY ' + str(i + 1))
        cat = st.selectbox('Select survey category', ('', 'A', 'B'), key = str(i) + 'c')
        if cat != '':
            a = st.file_uploader("Upload Call logs (csv)", key = str(i) + 'a')
            if a:
                a = pd.read_csv(a, header=5)
                a = a[a['Dial Leg'] == 'agent']
                a = a[['Agent Username', 'Call Start DT', 'Call Dur Connected', 'Call Clearing Value']]
                #a['Call Start DT'] = [x[:10] for x in a['Call Start DT']]
                a['Call Start DT'] = pd.to_datetime(
                    a['Call Start DT'], format='%Y/%m/%d').apply(lambda x: dt.datetime.strftime(x, '%d/%m/%Y'))
                a = clean_names(a, 'Agent Username')
                agents = list(a['Agent Username'].unique())
                dates = list(a['Call Start DT'].unique())
                if len(dates) > 1:
                    st.error('Please ONLY upload call logs from ONE day')
                else:
                    calls = [len(a[a['Agent Username'] == x]) for x in agents]
                    avg_dur = [sum(a[a['Agent Username'] == x]['Call Dur Connected']) / len(a[a['Agent Username'] == x]) for x in agents]
                    df1 = pd.DataFrame({'Agent': agents,
                                        'Calls Attempted': calls,
                                        'Average Call Dur (s)': avg_dur})

            b = st.file_uploader("Upload Survey responses (csv/xlsx)", key = str(i) + 'b')
            if b:
                if b.name[-3:] == 'csv':
                    b = pd.read_csv(b, na_filter = False)
                else:
                    b = pd.read_excel(b, na_filter = False)
                col1 = st.selectbox('Select column to define CR', ['', 'Row counts'] + list(b.columns), key = str(i) + 'd')
                col2 = st.selectbox('Select column to define CC Agent', [''] + list(b.columns), key = str(i) + 'e')
                col3 = st.selectbox('Select column Date column', [''] + list(b.columns), key = str(i) + 'f')
                if col1 != '' and col2 != '' and col3 != '':
                    b[col3] = pd.to_datetime(
                        b[col3], format='%Y/%m/%d').apply(lambda x: dt.datetime.strftime(x, '%d/%m/%Y'))
                    agents = list(b[col2].unique())
                    dates = list(b[col3].unique())
                    if len(dates) > 1:
                        st.error('Please ONLY upload call logs for ONE day')
                    else:
                        if col1 == 'Row counts':
                            df2 = pd.DataFrame({'Agent': list(dict(b[col2].value_counts()).keys()), 'CR': list((b[col2].value_counts())) })

                        else:
                            b = b[b[col1] != '']
                            df2 = pd.DataFrame({'Agent': list(dict(b[col2].value_counts()).keys()), 'CR': list((b[col2].value_counts())) })
                        
                        df2 = clean_names(df2, 'Agent')

                        if isinstance(a, pd.DataFrame) and isinstance(b, pd.DataFrame):
                            df = pd.merge(df1,df2,on='Agent',how='inner')
                            df['Calls Attempted'] = [int(x) if math.isnan(x) == False else 0 for x in df['Calls Attempted']]
                            df['CR'] = [int(x) if math.isnan(x) == False else 0 for x in df['CR']]
                            df['Calls-CR'] = df['Calls Attempted'] - df['CR']
                            if cat == 'A':
                                #df['Points'] = (df['Calls-CR'] // 50 * 10) + (df['CR'] * 10)
                                df['Points'] = (df['Calls-CR'] * 0.2) + (df['CR'] * 10)
                            else:
                                #df['Points'] = (df['Calls-CR'] // 80 * 20) + (df['CR'] * 20)
                                df['Points'] = (df['Calls-CR'] * 0.25) + (df['CR'] * 20)
                            df = df[['Agent', 'Calls Attempted', 'CR', 'Points', 'Average Call Dur (s)']]
                            dfs.append(df)


    if len(dfs) == number:
        df = pd.concat(dfs).reset_index(drop = True)
        agents = list(df['Agent'].unique())
        calls = [sum(df[df['Agent'] == x]['Calls Attempted']) for x in agents]
        crs = [sum(df[df['Agent'] == x]['CR']) for x in agents]
        points = [sum(df[df['Agent'] == x]['Points']) for x in agents]
        avg_dur = [sum(df[df['Agent'] == x]['Average Call Dur (s)']) / len(df[df['Agent'] == x]) for x in agents]
        df = pd.DataFrame({'Agent': agents,
                            'Calls Attempted': calls,
                            'CR': crs,
                            'Points': points,
                            'Average Call Dur (s)': avg_dur})
        
        option2 = st.multiselect('Any agent(s) on Half Day/Double Duty?', list(df.Agent))
        scores ={}
        for n in range(len(option2)):
            score = st.radio(option2[n], ('Half Day', 'Double Duty (LTS)'), key = option2[n])
            if score == 'Half Day':
                scores[option2[n]] = 50
            else:
                scores[option2[n]] = 85

        option3 = st.button('Generate Daily Report')
        if option3:
            df['Points'] = [int(x) for x in df['Points']]
            df['Average Call Dur (s)'] = [int(x) for x in df['Average Call Dur (s)']]
            df = df.sort_values('Points', ascending= False).reset_index(drop = True)
            df['KPI'] = [scores[x] if x in scores else 100 for x in df['Agent']]
            df['Met KPI'] = ['O' if df['Points'][i] >= df['KPI'][i] else 'X' for i in df.index]
            df = df.drop(columns = 'KPI')
            df.index += 1
            df = df.style.applymap(color_kpi, subset=['Met KPI'])
            st.table(df)

            df_xlsx = to_excel(df)
            date = str(dates[0]).replace('/', '-')
            st.download_button(label='📥 Download Result',
                           data=df_xlsx,
                           file_name='CC-daily-report' + date +'.xlsx')
            


else:
    file1 = st.file_uploader("Upload daily report 1")
    file2 = st.file_uploader("Upload daily report 2")
    file3 = st.file_uploader("Upload daily report 3")
    file4 = st.file_uploader("Upload daily report 4")
    file5 = st.file_uploader("Upload daily report 5")
    files = [file1, file2, file3, file4, file5]
    
    option2 = st.button('Generate Day-to-Day Report')
    if option2:
        files = [pd.read_excel(x) for x in files if x]
        df = pd.concat(files).reset_index(drop = True)
        df['# Met KPI'] = [1 if x == 'O' else 0 for x in df['Met KPI']]
        agents = list(df['Agent'].unique())
        calls = [sum(df[df['Agent'] == x]['Calls Attempted']) for x in agents]
        crs = [sum(df[df['Agent'] == x]['CR']) for x in agents]
        n_kpi = [sum(df[df['Agent'] == x]['# Met KPI']) for x in agents]
        avg_dur = [sum(df[df['Agent'] == x]['Average Call Dur (s)']) / len(df[df['Agent'] == x]) for x in agents]
        avg_dur = [int(x) for x in avg_dur]
        df = pd.DataFrame({'Agent': agents,
                            'Calls Attempted': calls,
                            'CR': crs,
                            '# KPI Met': n_kpi,
                            'Average Call Dur (s)': avg_dur})
        df = df.sort_values(by=['# KPI Met','CR'], ascending= False).reset_index(drop = True)
        df.index += 1
        st.table(df)
        df_xlsx = to_excel(df)
        st.download_button(label='📥 Download Result',
                        data=df_xlsx,
                        file_name='CC-weekly-report' +'.xlsx')
