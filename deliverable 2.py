import pandas as pd
import plotly.express as px

# Load Excel
xls = pd.ExcelFile("february_generator2025.xlsx")
df = xls.parse('Retired', skiprows=2)

# Select relevant columns and rename for clarity
df = df[[
    'Plant Name',
    'Retirement Year',
    'Retirement Month',
    'Nameplate Energy Capacity (MWh)',
    'Latitude',
    'Longitude',
    'Plant State'
]].rename(columns={
    'Plant Name': 'plant_name',
    'Retirement Year': 'retirement_year',
    'Retirement Month': 'retirement_month',
    'Nameplate Energy Capacity (MWh)': 'capacity_mwh',
    'Latitude': 'lat',
    'Longitude': 'lon',
    'Plant State': 'state'
})

# Clean up data
df = df.dropna(subset=['retirement_year', 'retirement_month', 'lat', 'lon'])

df['capacity_mwh'] = pd.to_numeric(df['capacity_mwh'], errors='coerce')

# Convert date columns into a datetime object
df['retirement_date'] = pd.to_datetime(
    df['retirement_year'].astype(int).astype(str) + '-' +
    df['retirement_month'].astype(int).astype(str).str.zfill(2),
    errors='coerce'
)

# Drop bad dates and create a period label
df = df.dropna(subset=['retirement_date'])
df['retirement_period'] = df['retirement_date'].dt.to_period('M')

# Sort the dataframe by retirement_date in descending order
df_clean = df.dropna(subset=['retirement_period', 'lat', 'lon', 'capacity_mwh'])
df_clean = df_clean.sort_values(by='retirement_date', ascending=False)

# Optional: Preview the result
print(df_clean)

fig = px.scatter_geo(
    df_clean,
    lat='lat',
    lon='lon',
    size='capacity_mwh',
    color='state',
    hover_name='plant_name',
    animation_frame='retirement_period',
    projection='natural earth',
    title='U.S. Power Plant Retirements Over Time',
    size_max=20,
    opacity=0.6
)

fig.update_layout(
    margin=dict(l=0, r=0, t=30, b=0)
)

fig.show()
