import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt

# Load Excel for the map data
xls = pd.ExcelFile("february_generator2025.xlsx")
df = xls.parse('Operating', skiprows=2)

# Select relevant columns and rename for clarity
df = df[[
    'Plant Name',
    'Planned Retirement Year',
    'Planned Retirement Month',
    'Technology',
    'Sector',
    'Nameplate Capacity (MW)',
    'Latitude',
    'Longitude',
    'Plant State'
]].rename(columns={
    'Plant Name': 'plant_name',
    'Planned Retirement Year': 'retirement_year',
    'Planned Retirement Month': 'retirement_month',
    'Technology': 'technology',
    'Nameplate Capacity (MW)': 'capacity_mwh',
    'Latitude': 'lat',
    'Longitude': 'lon',
    'Plant State': 'state'
})

# Clean and preprocess
df['retirement_year'] = pd.to_numeric(df['retirement_year'], errors='coerce')
df['retirement_month'] = pd.to_numeric(df['retirement_month'], errors='coerce')
df['capacity_mwh'] = pd.to_numeric(df['capacity_mwh'], errors='coerce')
df = df.dropna(subset=['retirement_year', 'retirement_month', 'lat', 'lon'])

# Convert to integer for date formatting
df['retirement_year'] = df['retirement_year'].astype(int)
df['retirement_month'] = df['retirement_month'].astype(int)

df['retirement_date'] = pd.to_datetime(
    df['retirement_year'].astype(str) + '-' +
    df['retirement_month'].astype(str).str.zfill(2),
    errors='coerce'
)

df = df.dropna(subset=['retirement_date'])
df['retirement_period'] = df['retirement_date'].dt.to_period('M')

df_clean = df.dropna(subset=['retirement_period', 'lat', 'lon', 'capacity_mwh'])
df_clean = df_clean.sort_values(by='retirement_date', ascending=True)

# --- Build animation dataset ---
# Get unique sorted retirement periods
all_frames = sorted(df_clean['retirement_period'].unique())

# Expand data: For each frame, include all plants and assign a status
animation_rows = []

for period in all_frames:
    current_date = period.to_timestamp()
    for _, row in df_clean.iterrows():
        animation_rows.append({
            'lat': row['lat'],
            'lon': row['lon'],
            'capacity_mwh': row['capacity_mwh'],
            'plant_name': row['plant_name'],
            'retirement_date': row['retirement_date'],
            'technology': row['technology'],
            'status': 'Retired' if row['retirement_date'] <= current_date else 'Operating',
            'frame': str(period)
        })

df_anim = pd.DataFrame(animation_rows)

# Plot map!
fig = px.scatter_geo(
    df_anim,
    lat='lat',
    lon='lon',
    size='capacity_mwh',
    color='status',  # Red if retired, green/blue if still operating
    animation_frame='frame',
    hover_name='plant_name',
    hover_data={
        'technology': True,
        'capacity_mwh': True,
        'lat': False,
        'lon': False,
        'retirement_date': True,
        'status': False,
        'frame': False
    },
    scope='usa',
    title='U.S. Power Plant Retirements Over Time',
    size_max=20,
    opacity=0.7,
    color_discrete_map={'Retired': 'red', 'Operating': 'green'}
)

fig.update_layout(
    geo=dict(
        projection_type='albers usa',
        showland=True,
        landcolor='lightgray',
        showlakes=True,
        lakecolor='white',
        lataxis=dict(range=[24, 50]),
        lonaxis=dict(range=[-125, -66]),
        center=dict(lat=37.0902, lon=-95.7129),
    ),
    margin=dict(l=0, r=0, t=30, b=0)
)

# Save the map as HTML
fig.show()
