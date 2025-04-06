import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt

# Load Excel file
xls = pd.ExcelFile("february_generator2025.xlsx")

# ----------------- OPERATING -------------------
df_retire = xls.parse('Operating', skiprows=2)

df_retire = df_retire[[
    'Plant Name', 'Planned Retirement Year', 'Planned Retirement Month',
    'Technology', 'Sector', 'Nameplate Capacity (MW)',
    'Latitude', 'Longitude', 'Plant State'
]].rename(columns={
    'Plant Name': 'plant_name',
    'Planned Retirement Year': 'year',
    'Planned Retirement Month': 'month',
    'Technology': 'technology',
    'Nameplate Capacity (MW)': 'capacity_mwh',
    'Latitude': 'lat',
    'Longitude': 'lon',
    'Plant State': 'state'
})

df_retire['status'] = 'Retired'

# Clean
df_retire['year'] = pd.to_numeric(df_retire['year'], errors='coerce')
df_retire['month'] = pd.to_numeric(df_retire['month'], errors='coerce')
df_retire['capacity_mwh'] = pd.to_numeric(df_retire['capacity_mwh'], errors='coerce')
df_retire = df_retire.dropna(subset=['year', 'month', 'lat', 'lon'])

# Combine to date
df_retire['year'] = df_retire['year'].astype(int)
df_retire['month'] = df_retire['month'].astype(int)
df_retire['event_date'] = pd.to_datetime(
    df_retire['year'].astype(str) + '-' + df_retire['month'].astype(str).str.zfill(2),
    errors='coerce'
)

df_retire['event_period'] = df_retire['event_date'].dt.to_period('M')
df_retire = df_retire.dropna(subset=['event_period'])

# ----------------- PLANNED -------------------
df_plan = xls.parse('Planned', skiprows=2)

df_plan = df_plan[[
    'Plant Name', 'Planned Operation Year', 'Planned Operation Month',
    'Technology', 'Sector', 'Nameplate Capacity (MW)',
    'Latitude', 'Longitude', 'Plant State'
]].rename(columns={
    'Plant Name': 'plant_name',
    'Planned Operation Year': 'year',
    'Planned Operation Month': 'month',
    'Technology': 'technology',
    'Nameplate Capacity (MW)': 'capacity_mwh',
    'Latitude': 'lat',
    'Longitude': 'lon',
    'Plant State': 'state'
})

df_plan['status'] = 'Planned'

# Clean
df_plan['year'] = pd.to_numeric(df_plan['year'], errors='coerce')
df_plan['month'] = pd.to_numeric(df_plan['month'], errors='coerce')
df_plan['capacity_mwh'] = pd.to_numeric(df_plan['capacity_mwh'], errors='coerce')
df_plan = df_plan.dropna(subset=['year', 'month', 'lat', 'lon'])

df_plan['year'] = df_plan['year'].astype(int)
df_plan['month'] = df_plan['month'].astype(int)
df_plan['event_date'] = pd.to_datetime(
    df_plan['year'].astype(str) + '-' + df_plan['month'].astype(str).str.zfill(2),
    errors='coerce'
)

df_plan['event_period'] = df_plan['event_date'].dt.to_period('M')
df_plan = df_plan.dropna(subset=['event_period'])

# --------------------------------------
# --- Build animation dataset for df_retire ---
# Get unique sorted retirement periods
all_frames_retire = sorted(df_retire['event_period'].unique())

# Expand data: For each frame, include all plants and assign a status
animation_rows_retire = []

for period in all_frames_retire:
    current_date = period.to_timestamp()
    for _, row in df_retire.iterrows():
        animation_rows_retire.append({
            'lat': row['lat'],
            'lon': row['lon'],
            'capacity_mwh': row['capacity_mwh'],
            'plant_name': row['plant_name'],
            'event_date': row['event_date'],
            'technology': row['technology'],
            'status': 'Retired' if row['event_date'] <= current_date else 'Operating',
            'frame': str(period),
            'source': 'df_retire'  # Mark this row as coming from df_retire
        })

df_anim_retire = pd.DataFrame(animation_rows_retire)

# --- Build animation dataset for df_plan ---
# Get unique sorted operation periods for df_plan
all_frames_plan = sorted(df_plan['event_period'].unique())

# Expand data: For each frame, include all plants and assign a status
animation_rows_plan = []

for period in all_frames_plan:
    current_date = period.to_timestamp()
    for _, row in df_plan.iterrows():
        # Only show planned plants once their event_date has passed
        if row['event_date'] <= current_date:
            animation_rows_plan.append({
                'lat': row['lat'],
                'lon': row['lon'],
                'capacity_mwh': row['capacity_mwh'],
                'plant_name': row['plant_name'],
                'event_date': row['event_date'],
                'technology': row['technology'],
                'status': 'Planned',  # They are "Planned" once they appear
                'frame': str(period),
                'source': 'df_plan'  # Mark this row as coming from df_plan
            })

df_anim_plan = pd.DataFrame(animation_rows_plan)

# --------------------------------------
# Combine both animations into one DataFrame
df_combined = pd.concat([df_anim_retire, df_anim_plan])

# Plot map with both animations!
fig = px.scatter_geo(
    df_combined,
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
        'event_date': True,
        'status': False,
        'frame': False
    },
    scope='usa',
    title='U.S. Power Plant Retirements and Planned Operations Over Time',
    size_max=20,
    opacity=0.7,
    color_discrete_map={'Retired': 'red', 'Operating': 'green', 'Planned': 'blue'}
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

# Save to HTML
fig.write_html('map.html')

# Load the data for the capacity plot
xls = pd.ExcelFile("Retirement Model.xlsx")
df = xls.parse('Remaining Capacity')

# Ensure the years are the columns starting from index 2 onward
years = df.columns[2:]  # Skip 'Region' and 'Tech Type' columns

# Prepare the plot for capacity per region
fig, axs = plt.subplots(2, 2, figsize=(16, 12))  # Create a 2x2 grid of subplots
axs = axs.flatten()  # Flatten the 2x2 array for easier indexing

df = df.dropna(subset=['Region', 'Tech Type'])  # Drop rows with NaN in 'Region' or 'Tech Type'

# Loop over each region
regions = df['Region'].unique()
for i, region in enumerate(regions):
    region_data = df[df['Region'] == region].dropna(subset=['Tech Type'])

    # Loop over each unique Tech Type in the region
    for tech in region_data['Tech Type'].unique():
        tech_data = region_data[region_data['Tech Type'] == tech]

        # Plot the data for each tech type
        axs[i].plot(years, tech_data.iloc[0, 2:], label=tech)

    # Add labels, title, and legend for each subplot
    axs[i].set_title(f"Remaining Capacity for {region} Over Time")
    axs[i].set_xlabel("Year")
    axs[i].set_ylabel("Capacity (GW)")
    axs[i].legend(title="Technology Type")
    axs[i].tick_params(axis='x', rotation=45)

# Save the Matplotlib figure as an image
plt.tight_layout()
plt.savefig('capacity_plot.png')
plt.close()

# Updated HTML with fancy styling
html_content = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Energy Data Visualization</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f4f4f9;
            margin: 0;
            padding: 0;
            color: #333;
        }

        header {
            background-color: #4CAF50;
            color: white;
            text-align: center;
            padding: 20px;
            font-size: 2em;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }

        h2 {
            color: #4CAF50;
            text-align: center;
            margin-top: 40px;
        }

        .section {
            margin-bottom: 40px;
        }

        iframe {
            width: 100%;
            height: 600px;
            border: none;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }

        img {
            display: block;
            margin: 0 auto;
            width: 100%;
            max-width: 1100px;
            border: 2px solid #ddd;
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }

        footer {
            text-align: center;
            font-size: 0.9em;
            color: #777;
            margin-top: 40px;
            padding: 20px;
            background-color: #f1f1f1;
        }
    </style>
</head>
<body>
    <header>
        Energy Data Visualization
    </header>

    <div class="container">
        <div class="section">
            <h2>Power Plant Retirements Over Time</h2>
            <iframe src="map.html"></iframe>
        </div>

        <div class="section">
            <h2>Remaining Capacity Over Time (by Region)</h2>
            <img src="capacity_plot.png" alt="Capacity Plot">
        </div>
    </div>

    <footer>
        <p>&copy; 2025 Energy Data Visualizations | All Rights Reserved</p>
    </footer>
</body>
</html>
"""

# Save the updated HTML content to a file
with open('index.html', 'w') as f:
    f.write(html_content)

# Now you can open 'index.html' in any browser to view the combined visualization.