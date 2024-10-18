import plotly.graph_objects as go
import pandas as pd

file_path = r"C:\Users\ibrah\OneDrive\Documents\Projects\Experience_Monitoring_Dashboard\data_2.xlsx"
df = pd.read_excel(file_path)

# Mapping categorical data to numbers
df["Most Rider"] = df["Most Rider"].map({"Medical Rider": 2, "CI Rider": 1})

categories = df.iloc[:, 1:].columns
customer = df.iloc[:, 0]

df_plot = df.drop(columns=["Date"])

# Define custom ranges for each axis based on the scale of each variable
custom_ranges = {
    "Age": [0, 100],
    "Sum Assured": [0, 50000],  # Adjust based on your dataset
    "Contribution": [0, 500],  # Adjust based on your dataset
    "Term": [0, 20],  # Adjust based on your dataset
    "Gender": [0, 1],
    "Most Rider": [1, 2],
}

# Create the spider chart using the normalized data
fig = go.Figure()

# Loop through each row to plot each customer profile
for i, row in df_plot.iterrows():
    # Extract actual values for hover data
    actual_values = row.tolist()

    # Normalize the data using the custom ranges for plotting
    normalized_values = []
    for category in categories:
        normalized_value = (row[category] - custom_ranges[category][0]) / (
            custom_ranges[category][1] - custom_ranges[category][0]
        )
        normalized_values.append(normalized_value)

    # Close the loop for both normalized and actual values
    normalized_values += [normalized_values[0]]
    actual_values += [actual_values[0]]

    fig.add_trace(
        go.Scatterpolar(
            r=normalized_values,
            theta=categories.tolist() + [categories[0]],
            fill="toself",
            name=customer[i],
            customdata=[actual_values],  # Pass the actual (non-normalized) values
            hovertemplate="%{theta}: %{customdata[theta]}",  # Show actual values in the hover
        )
    )

# Update layout for better visualization
fig.update_layout(
    polar=dict(
        radialaxis=dict(
            visible=True, range=[0, 1]
        ),  # Normalized range 0-1 for all axes
    ),
    showlegend=True,
)

# Show the figure in Dash or standalone
fig.show()
