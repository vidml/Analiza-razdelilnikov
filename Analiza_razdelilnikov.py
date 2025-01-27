headers=['Enota','Objekt','Sklic','Strošek','Opis stroška','Dobavitelj','ID za DDV','Št. računa','Datum računa','Znesek računa','Datum opravljene storitve','St. DDV','Znesek za delitev','Cena za razdelitev','Ključ delitve','Vaša količina','Skupna količina','Vrednost z DDV']
import os
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import plotly.io as pio
import plotly.graph_objects as go
from plotly.subplots import make_subplots


pio.renderers.default='browser'



# Find all files that match the pattern
files = [file for file in os.listdir() if file.startswith('Razdelilnik_') and file.endswith('.xls')]

# Process tables
all_tables = []
datumi_racunov = []
unique_arr=np.array([])
poraba_vode=np.array([])
# costs=np.zeros([])
# tcosts=np.zeros([])

unique_values=pd.DataFrame; #Array of unique cost names. 

for fi in range(len(files)):#Read all files:
    file=files[fi]
    print(file)
    # Read HTML table and clean it up
    df=pd.read_html(file,decimal=',', thousands='.',encoding='ISO-8859-2')#,encoding='utf8',charset='ISO 8859-2'
    df = [df_fracture for df_fracture in df if df_fracture.shape != (1, 4)] #Omit headers
    df = pd.concat(df,ignore_index=True)
    df=df[df[0]!="Enota"] #Omit these rows
    df=df.reset_index(drop=True)
    for il, line in enumerate(df[3]):
        if (line=='OGREVANJE PROSTOROV - TOPLOTNA ENERGIJA'):
            df.loc[il,3] = 'TOPLOTNA ENERGIJA'
    ps_lines=df[3]=="Poraba prostovoljno zbranih sredstev (999 ZBIRANJE SREDSTEV ZA VZDRŽEVANJE-PS)" # Preimenuj OGREVANJE PROSTOROV - TOPLOTNA ENERGIJA v TOPLOTNA ENERGIJA
    ps_lines = ps_lines | ps_lines.shift(-1, fill_value=False)
    df = df[~ps_lines]
    df=df.reset_index(drop=True)
    # Get all unique values in column 3
    unique_arr=np.unique(np.append(unique_arr,df[3])) # Imamo seznam vseh unikatnih
    # Parse and calculate average service dates
    dates=[];
    for i, datestr in enumerate(df[10]):
        if len(datestr)>12:
            d1=pd.to_datetime(datestr[0:10],dayfirst='true')
            d2=pd.to_datetime(datestr[13::],dayfirst='true')
            d=datetime.fromtimestamp(np.mean([d1.timestamp(),d2.timestamp()]))
        elif len(datestr)==8:
            d=datetime_object = datetime.strptime(datestr, '%d%m%Y')
        else:
            d=pd.NaT()
            print("NaT entry")
        df.loc[i,10]=d
        dates.append(d)
    datumi_racunov.append(datetime.fromtimestamp(np.mean([date.timestamp() for date in dates])))
    for il, line in enumerate(df[3]):
        if (line=='HLADNA VODA - PORABA'):
            vodam3=df.loc[il,15][:6]
    poraba_vode=np.append(poraba_vode,float(vodam3[:len(vodam3)-3].replace(",",".")))
    all_tables.append(df.reset_index(drop=True))


    
costs = [[0] * len(all_tables) for _ in range(len(unique_arr))]  # Individual costs (stanovanja, po postavki)
tcosts = [[0] * len(all_tables) for _ in range(len(unique_arr))] # Total costs (polni stroški stavbe, po postavki)
for i, table in enumerate(all_tables): #For each df (i=seq. month)
    for j, cname in enumerate(table[3]): #For each line in table
        a=np.where(unique_arr==cname)[0][0] # a is the ID in unique array
        costs[a][i]=costs[a][i]+table[17][j] # Submit (add) cost to unified table of costs
        tcosts[a][i]=table[12][j] # replace (do not SUM, ker jih upravnik v tabeli ponavlja in ne deli)
        tcosts[a][i]=tcosts[a][i]+table[12][j] # replace (do not SUM, ker jih upravnik v tabeli ponavlja in ne deli)
        
cum_costs=[[0] for _ in range(len(unique_arr))]
for ci, costline in enumerate(costs):
    cum_costs[ci]=sum(costline)  #Cumulative costs so far (za vsako postavko)

cum_tcosts=[[0] for _ in range(len(unique_arr))]
for tci, tcostline in enumerate(tcosts):
    cum_tcosts[tci]=sum(tcostline)  #Cumulative costs so far (za vsako postavko)



#### Imamo (unsorted): unique_arr, costs, cum_costs, tcosts. Sort these by the value of cum_costs, so according to indexes in sortlist
sortlist=np.argsort(cum_costs)[::-1]
sorted_names=unique_arr[sortlist]
sorted_costs = np.array([costs[i] for i in sortlist])
sorted_cum_costs = np.array([cum_costs[i] for i in sortlist])
sorted_cum_tcosts= np.array([cum_tcosts[i] for i in sortlist])
sorted_tcosts = np.array([tcosts[i] for i in sortlist])
avg_costs=np.array([sum(cost) for cost in sorted_costs])/len(datumi_racunov)

# Seštej stroške v mesečne položnice
monthly_costs=np.zeros(np.shape(datumi_racunov))
for m in range(len(datumi_racunov)):
    for strosek in sorted_costs:
        monthly_costs[m]=monthly_costs[m]+strosek[m]

ogrevanje_stanovanja=np.array(sorted_costs[np.where(sorted_names == "TOPLOTNA ENERGIJA")[0][0]])
ogrevanje_stavbe=np.array(sorted_tcosts[np.where(sorted_names == "TOPLOTNA ENERGIJA")[0][0]])
dpr=(ogrevanje_stanovanja)/(ogrevanje_stavbe)

avgs=(sorted_cum_costs)/len(datumi_racunov)
del_pol=sorted_cum_costs/sum(sorted_cum_costs)
del_stavbe=np.divide(sorted_cum_costs,sorted_cum_tcosts)


fig = make_subplots(
    rows=4, cols=3, 
    specs=[
        [{"type": "scatter", "colspan": 2, "rowspan": 2}, None, {"type": "pie","rowspan": 2}], 
        [None,None,None],#[{"type": "scatter"}, {"type": "scatter"}, {"type": "scatter"}, {"type": "scatter"}], 
        [{"type": "scatter"}, {"type": "scatter"}, {"type": "xy", "secondary_y": True}],
        [{"type": "scatter"}, {"type": "scatter"}, {"type": "scatter"}]
    ],
    subplot_titles=[
        "Pregled zgodovine stroškov", "Povprečni stroški",
        "Mesečni stroški", "Poraba vode", "Cene ogrevanja",
        "Skupna elektrika", "Odvoz smeti","Delež cene ogrevanja bloka"
    ]
)


# Add scatter plots
# for i in range(0,len(sorted_names)-1):
#     fig.add_trace(go.Scatter(x=datumi_racunov, y=sorted_costs[i], name=sorted_names[i],legendgroup=sorted_names[i]), row=1, col=1)
fig.update_yaxes(title_text="Strošek (EUR)", row=1, col=1)
fig.add_trace(go.Scatter(x=datumi_racunov, y=monthly_costs, name="Mesečna položnica",showlegend=False), row=3, col=1)
fig.update_yaxes(title_text="Strošek (EUR)", row=3, col=1)
fig.add_trace(go.Scatter(x=datumi_racunov, y=sorted_costs[np.where(sorted_names == "ODVOZ SMETI")[0][0]], legendgroup="ODVOZ SMETI",name="Odvoz smeti",showlegend=False), row=4, col=2)
fig.update_yaxes(title_text="Strošek (EUR)", row=4, col=2)
fig.add_trace(go.Scatter(x=datumi_racunov, y=sorted_costs[np.where(sorted_names == "SKUPNA ELEKTRIKA")[0][0]], legendgroup="SKUPNA ELEKTRIKA",name="Skupna elektrika",showlegend=False), row=4, col=1)
fig.update_yaxes(title_text="Strošek (EUR)", row=4, col=1)
fig.add_trace(go.Scatter(x=datumi_racunov, y=dpr*100,name="DPR", showlegend=False), row=4, col=3)
fig.update_yaxes(title_text="Delež (%)", row=4, col=3)
fig.add_trace(go.Scatter(x=datumi_racunov, y=poraba_vode,name="Poraba vode", showlegend=False), row=3, col=2)  
fig.update_yaxes(title_text="Poraba vode (m³)", row=3, col=2)

fig.add_trace(go.Scatter(x=datumi_racunov, y=ogrevanje_stanovanja,name="Stanovanje"), row=3, col=3,secondary_y=False)
fig.add_trace(go.Scatter(x=datumi_racunov, y=ogrevanje_stavbe, name="Stavba"), row=3, col=3,secondary_y=True)
fig.update_yaxes(title_text="Ogrevanje stanovanja (EUR)", secondary_y=False, row=3, col=3)
fig.update_yaxes(title_text="Ogrevanje stavbe (EUR)", secondary_y=True, row=3, col=3)


# Add a pie chart
fig.add_trace(go.Pie(labels=sorted_names, values=avg_costs, name="Povprečni deleži",legendgroup="Group1",showlegend=False,textinfo='label+percent',textposition='inside'), row=1, col=3)

# Update layout
# fig.update_layout(height=2160/2, width=3840/2, title_text="Pregled obračunov stroškov upravljanja")
fig.update_layout( title_text="Pregled obračunov stroškov z mojUpravnik")

df_summary = pd.DataFrame({
    'Strošek': sorted_names,
    'Povprečno (EUR)': avgs,
    'Dosedanja vsota(EUR)': sorted_cum_costs,
    'Delež položnic': del_pol,
    'Dosedanja vsota stavbe (EUR)(ERROR!)': sorted_cum_tcosts,
    'Delež pol. stavbe (ERROR!)': del_stavbe
})

# Convert datumi_racunov to strings for column headers
columns = [pd.to_datetime(date).strftime('%Y-%m') for date in datumi_racunov]


df_poraba = pd.DataFrame({
    'Poraba vode (m³)': poraba_vode,
    'Ogrevanje (delež oz. DPR)': dpr
}, index=columns)
df_poraba=df_poraba.T

# Create DataFrames for sorted costs and total costs
df_sorted_costs = pd.DataFrame(sorted_costs, columns=columns)
df_sorted_tcosts = pd.DataFrame(sorted_tcosts, columns=columns)

# Add sorted_names as the identifier column
df_sorted_costs.insert(0, 'Strošek', sorted_names)
df_sorted_tcosts.insert(0, 'Strošek', sorted_names)


with pd.ExcelWriter('Rezultati za vse razdelilnike.xlsx') as writer:  
    df_summary.to_excel(writer, sheet_name='Povzetek')
    df_poraba.to_excel(writer, sheet_name='Poraba')
    df_sorted_costs.to_excel(writer, sheet_name='Mesečni stroški stanovanja')
    df_sorted_tcosts.to_excel(writer, sheet_name='Mesečni stroški stavbe (ERROR!)')
    


# Using df_long from before
df_long = df_sorted_costs.melt(
    id_vars="Strošek",
    var_name="Date",
    value_name="Cost"
)
df_long["Date"] = pd.to_datetime(df_long["Date"])

# Create the area plot using px.area (not showing yet)
area_fig = px.area(
    df_long,
    x="Date",
    y="Cost",
    color="Strošek",
    labels={"Cost": "Cost", "Date": "Date"},
)

# Add traces from the area plot to the specified subplot (col=1, row=1)
for trace in area_fig.data:
    trace.hovertemplate = "%{fullData.name}<br>%{x|%Y-%m}<br>%{y:.2f} €<extra></extra>"
    fig.add_trace(trace, row=1, col=1)

# Update layout for the subplot
fig.update_yaxes(title_text="Strošek (EUR)", row=1, col=1)
fig.update_xaxes(title_text="Datum", row=1, col=1)

# Show the updated figure
fig.show()



### Plot % of building cost in your cost for individual cost in sorted_names
# x_values = np.arange(len(del_stavbe))

# # Create a line plot
# afig = go.Figure()

# afig.add_trace(go.Scatter(
#     x=x_values,  # X-axis values
#     y=del_stavbe,      # Y-axis values
#     mode='lines', # Line plot
#     name='Line Plot'  # Legend name
# ))
# afig.show()


# pip install pyinstaller
# pyinstaller --onefile Analiza_razdelilnikov.py                        or      pyinstaller --clean --onefile <your_script.py>
# Edit Analiza_razdelilnikov.spec file: import sys ; sys.setrecursionlimit(sys.getrecursionlimit() * 5)
# pyinstaller Analiza_razdelilnikov.spec


