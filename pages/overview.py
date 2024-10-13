from dash import Dash, html, dcc, Input, Output, dash_table
import dash
import plotly.graph_objs as go
import pandas as pd
import plotly.express as px
import numpy as np

# register page in directory
dash.register_page(__name__, path="/")

# --------------------------------------------------
# Data Processing
# --------------------------------------------------
# Data read
data_chart = pd.read_excel("data.xlsx", sheet_name="Sheet1")
data_table = pd.read_excel("data.xlsx", sheet_name="Sheet2")

# Convert to percentage
data_chart[["21Q1", "22Q1", "Current"]] = data_chart[["21Q1", "22Q1", "Current"]] * 100


def create_dropdown():
    dropdown = dcc.Dropdown(
        id="overview-selected-product-dropdown",  # Assign an ID to the dropdown
        options=[
            {"label": "All Medical", "value": "All"},
            {"label": "Medi A", "value": "Medi A"},
            {"label": "Medi B", "value": "Medi B"},
            {"label": "Medi C", "value": "Medi C"},
            {"label": "Medi D", "value": "Medi D"},
        ],
        value="Medi A",
        placeholder="Select Product...",
        className="selected-product-dropdown",  # Optional: apply CSS class
        style={
            "background-color": "rgb(35, 36, 72)",  # Custom background color
            "border": "1px solid rgb(35, 36, 72)",  # Border style
            "border-radius": "5px",  # Rounded corners
        },
    )
    return dropdown


def create_fan():
    # Step 1: Define existing data points for the line chart
    x_values = ["21Q1", "22Q1", "Current"]  # Time points for the actual values
    y_values = [60, 85, 98]  # Loss ratio values for actual time points

    # Step 2: Simulate the projected loss ratio for "Projected 24Q1"
    np.random.seed(42)  # Set a seed for reproducibility
    k_values = np.random.normal(
        loc=10, scale=5, size=1000
    )  # Normally distributed k values
    projected_value = 98 + k_values  # Add k to the current value of 98%

    # Step 3: Create the fan chart by simulating confidence intervals from the projections
    x_projected = ["Projected"] * len(
        k_values
    )  # All projections for the next time point

    # Create traces for the fan chart (shaded areas)
    fig = go.Figure()

    # Step 4: Add actual line chart
    fig.add_trace(
        go.Scatter(
            x=x_values + ["Projected"],
            y=[
                60,
                85,
                98,
                np.mean(projected_value),
            ],  # Include the mean of the projected values
            mode="lines+markers",
            name="Mean Loss Ratio",
            line=dict(color="rgb(0, 127, 61)", width=2.5),
            marker=dict(color="rgb(0, 127, 61)", size=5),
            showlegend=False,
        )
    )

    # Step 5: Add confidence intervals for the fan chart
    percentiles = [0, 25, 50, 75, 100]  # Percentiles to create fan chart intervals
    colors = [
        "rgb(255, 83, 83)",
        "rgba(255, 230, 83,0.6)",
        "rgba(255, 230, 83,0.6)",
        "rgb(255, 83, 83)",
    ]

    # Sort projected values to calculate percentiles
    sorted_proj = np.sort(projected_value)

    # Add percentiles as shaded areas
    for i in range(len(percentiles) - 1):
        lower = np.percentile(sorted_proj, percentiles[i])
        upper = np.percentile(sorted_proj, percentiles[i + 1])
        fig.add_trace(
            go.Scatter(
                x=["Current", "Projected", "Projected", "Current"],
                y=[98, lower, upper, 98],
                fill="toself",
                fillcolor=colors[i],
                line=dict(color="rgba(255,255,255,0)"),
                showlegend=False,
            )
        )

    # Step 6: Update layout for better display
    fig.update_layout(
        paper_bgcolor="rgb(35, 36, 72)",
        plot_bgcolor="rgb(35, 36, 72)",
        yaxis=dict(
            range=[50, 130],
            showgrid=False,
        ),
        xaxis=dict(
            showgrid=False,
            tickfont=dict(
                # size=20,
                weight="bold",
            ),
        ),
        font=dict(
            family="Arial, sans-serif",
            color="rgb(213, 215, 224)",
        ),
        margin=dict(
            l=15,  # Left margin
            r=15,  # Right margin
            t=20,  # Top margin
            b=10,  # Bottom margin
        ),
        height=250,
        # width=900,
    )

    # Step 7: Calculate min and max projected values
    min_proj_value = np.min(projected_value)
    max_proj_value = np.max(projected_value)

    return fig


def create_bullet():
    # Current value (83.3% loss ratio)
    current_value = 83.3

    # Create a bullet chart using a bar to represent the current value
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number+delta",
            value=current_value,
            delta={
                "reference": 75,
                "suffix": "%",
                "increasing": {"color": "red"},
                "decreasing": {"color": "green"},
            },
            gauge={
                "shape": "bullet",
                "axis": {
                    "range": [None, 120],
                    "visible": True,
                    "tickcolor": "rgb(213, 215, 224)",
                    "tickfont": {
                        "color": "rgb(213, 215, 224)",
                    },
                },
                "bar": {
                    "color": "rgb(0, 127, 61)",
                    "thickness": 0.5,
                },
                "threshold": {
                    "line": {"color": "orange", "width": 3},
                    "thickness": 0.75,
                    "value": 90,
                },
                "steps": [
                    {
                        "range": [0, 80],
                        "color": "rgba(174, 177, 210, 0.85)",
                    },
                    {
                        "range": [80, 100],
                        "color": "rgb(255, 230, 83)",
                    },
                    {
                        "range": [100, 120],
                        "color": "rgb(255, 83, 83)",
                    },
                ],
            },
            domain={"x": [0, 1], "y": [0.4, 1]},
            number={
                "suffix": "%",
                "font": {
                    "size": 34,
                    "color": "white",
                    "family": "Arial",
                    "weight": "bold",
                },
            },
        )
    )

    # Update layout to add titles and customize appearance
    fig.update_layout(
        paper_bgcolor="rgb(35, 36, 72)",
        plot_bgcolor="rgb(35, 36, 72)",
        margin=dict(l=60, r=60, t=15, b=0),
        height=100,
    )

    return fig


def create_box(val):
    # Generate random data with exponential distribution, mean ~ 100
    np.random.seed(42)
    data = np.random.exponential(scale=val, size=20)

    # Add jitter to y-axis for better visibility
    jitter = np.random.normal(
        0, 0.1, len(data)
    )  # Add small random noise to y-values for jitter

    # Calculate quartiles
    q1 = np.percentile(data, 25)  # 1st quartile (25th percentile)
    median = np.median(data)  # Median (50th percentile)
    q3 = np.percentile(data, 75)  # 3rd quartile (75th percentile)

    # Create the figure
    fig = go.Figure()

    # Add all data points as a scatter plot with jitter
    fig.add_trace(
        go.Scatter(
            x=data,  # Data on the x-axis
            y=jitter,  # Apply jitter to y-values
            mode="markers",
            marker=dict(color="lightgreen", size=5),  # Style for the data points
            name="Data Points",
        )
    )

    # Highlight the specific value 86.4
    fig.add_trace(
        go.Scatter(
            x=[val],  # Highlight the point 86.4
            y=[0],  # Keep y=0 for the highlighted point
            mode="markers",
            marker=dict(color="red", size=10),  # Red marker for the highlight
            name="Highlighted Value",
        )
    )

    # Add vertical lines for 1st quartile, median, and 3rd quartile
    fig.add_shape(
        type="line", x0=q1, y0=-0.5, x1=q1, y1=0.5, line=dict(color="green", width=2)
    )  # 1st Quartile
    fig.add_shape(
        type="line",
        x0=median,
        y0=-0.5,
        x1=median,
        y1=0.5,
        line=dict(color="green", width=2),
    )  # Median
    fig.add_shape(
        type="line", x0=q3, y0=-0.5, x1=q3, y1=0.5, line=dict(color="green", width=2)
    )  # 3rd Quartile

    # Update layout to remove axis titles and values, and adjust appearance
    fig.update_layout(
        xaxis=dict(visible=False),  # Hide x-axis values and title
        yaxis=dict(visible=False),  # Hide y-axis values and title
        showlegend=False,  # Remove legend
        margin=dict(l=100, r=100, t=20, b=0),  # Set margins to 0
        height=60,  # Set height to 60
        # width=200,
        paper_bgcolor="rgb(35, 36, 72)",  # Set background color
        plot_bgcolor="rgb(35, 36, 72)",  # Set plot area color
    )

    return fig


def create_delta_card(val, ref):
    delta_card = go.Figure()

    delta_card.add_trace(
        go.Indicator(
            mode="number+delta",
            value=val,
            number={
                "valueformat": ".1f",
                "font": {
                    "color": "rgb(213, 215, 224)",
                    "weight": "bold",
                    "family": "Arial, sans-serif",
                },
            },
            delta={"position": "right", "reference": ref, "valueformat": ".1f"},
            domain={"x": [0, 1], "y": [0, 1]},
        )
    )

    # Set margins to zero and background color to your preference
    delta_card.update_layout(
        margin=dict(l=0, r=0, t=0, b=10),
        paper_bgcolor="rgb(35, 36, 72)",
        plot_bgcolor="rgb(35, 36, 72)",
        height=60,
    )

    return delta_card


def create_figure():

    # Define the data for the Plotly graph
    line_graph = go.Figure()

    # Use a built-in color palette (Plotly's default color palette)
    color_palette = px.colors.qualitative.Pastel

    # Add the trace for 'All' products to be prominent by default
    line_graph.add_trace(
        go.Scatter(
            x=["21Q1", "22Q1", "Current"],
            y=data_chart[data_chart["Product"] == "All"].iloc[0, 1:],
            mode="lines+markers",
            name="All",
            opacity=1.0,  # Full opacity
            text=["All"] * 3,  # Custom hover text
            hoverinfo="text+y",  # Show custom text and y value on hover
            line=dict(
                width=3, color=color_palette[0]
            ),  # Use the first color from the palette
            marker=dict(color=color_palette[0]),  # Marker color matching the line
        )
    )

    # Add a threshold line at 90%
    line_graph.add_trace(
        go.Scatter(
            x=["21Q1", "22Q1", "Current"],
            y=[90, 90, 90],  # Constant value at 90%
            mode="lines",
            line=dict(color="red", width=2, dash="dash"),  # Dashed line, more prominent
            name="Threshold (90%)",
            opacity=1.0,  # Full opacity
            text=["Threshold (90%)"] * 3,  # Custom hover text
            showlegend=False,  # Hide threshold line from legend
            hoverinfo="skip",  # Disable hover info for this line
        )
    )

    # Add traces for each product, initially semi-transparent
    products = data_chart["Product"].drop(
        data_chart[data_chart["Product"] == "All"].index
    )  # Exclude 'All' product

    for i, product in enumerate(products):
        line_graph.add_trace(
            go.Scatter(
                x=["21Q1", "22Q1", "Current"],
                y=data_chart[data_chart["Product"] == product].iloc[0, 1:],
                mode="lines+markers",
                name=product,
                line=dict(
                    width=2, color=color_palette[(i + 1) % len(color_palette)]
                ),  # Cycle through the palette for each product
                marker=dict(
                    color=color_palette[(i + 1) % len(color_palette)]
                ),  # Marker color matching the line
                opacity=0.8,  # Slight transparency
                text=[product] * 3,  # Custom hover text
                hoverinfo="text+y",  # Show custom text and y value on hover
            )
        )
    # Set layout for the graph
    line_graph.update_layout(
        plot_bgcolor="rgb(35, 36, 72)",
        paper_bgcolor="rgb(35, 36, 72)",
        title=dict(
            text="Current Loss Ratio vs Previous Studies",
            x=0,  # Align the title to the left (0 is fully left, 1 is fully right)
            xanchor="left",  # Anchor the title to the left of the x position
            yanchor="top",  # Align the title at the top of the chart
            font=dict(
                size=16,  # Font size for the title
                color="rgb(174, 177, 210)",  # Light red color for the title
            ),
        ),
        title_x=0.01,  # Ensure the title is aligned with the left side of the chart
        title_y=0.95,  # Adjust this to position the title slightly below the top
        legend=dict(
            orientation="h",  # Horizontal legend
            x=0.5,  # Center the legend horizontally
            y=-0.25,  # Position the legend below the title
            xanchor="center",  # Align legend horizontally
            yanchor="bottom",  # Align legend vertically
        ),
        yaxis=dict(
            showgrid=False,  # Remove horizontal gridlines (marker lines)
        ),
        xaxis=dict(
            showgrid=False,  # Ensure grid lines are visible
            tickfont=dict(
                size=20,
                weight="bold",
            ),
        ),
        margin=dict(
            l=15,  # Left margin
            r=15,  # Right margin
            t=20,  # Top margin
            b=10,  # Bottom margin
        ),
        font=dict(
            family="Arial, sans-serif",  # Apply the custom font to all chart elements
            color="rgb(213, 215, 224)",
            size=16,
        ),
        annotations=[
            dict(
                x="Current",  # The x-coordinate where the annotation will appear (rightmost point)
                y=90,  # The y-coordinate corresponding to the Threshold (90%) line
                xref="x",  # Reference to the x-axis
                yref="y",  # Reference to the y-axis
                text="Repricing Threshold (90%)",  # The text to display
                showarrow=False,  # Disable the arrow for a cleaner look
                font=dict(
                    family="var(--font-family)",  # Use custom font defined in CSS
                    size=14,  # Font size for the annotation
                    color="red",  # Set the text color to contrast with the dark background
                ),
                xanchor="right",  # Anchor the text to the left of the x-position
                yanchor="middle",  # Align the text vertically in the middle
                xshift=50,  # Shift the text x pixels to the right of the x coordinate
                yshift=-15,
            )
        ],
    )

    return line_graph


# --------------------------------------------------
# Dashboard Content Layout
# --------------------------------------------------

layout = html.Div(
    [
        # Header section
        html.Div(
            [
                html.Div(
                    [
                        html.Div("Co.", className="logo"),
                        html.Div(
                            [
                                html.Div(
                                    "Experience Monitoring Dashboard",
                                    className="title-dashboard",
                                ),
                                html.Div(
                                    "Actuarial Department",
                                    className="title-dashboard secondary",
                                ),
                            ],
                            className="title-dashboard-container",
                        ),
                    ],
                    className="header-container-left",
                ),
                html.Div(
                    html.Ul(
                        [
                            html.Li(
                                dcc.Link("Home", href="/product"),
                                className="menu-item",
                            ),
                            html.Li(
                                dcc.Link("New Business", href="/product"),
                                className="menu-item",
                            ),
                            html.Li(
                                dcc.Link("Mortality", href="/product"),
                                className="menu-item",
                            ),
                            html.Li(
                                dcc.Link("Morbidity", href="/"),
                                className="selected-menu-item",
                            ),
                            html.Li(
                                dcc.Link("Lapse", href="/product"),
                                className="menu-item",
                            ),
                        ]
                    ),
                    className="menu",
                ),
            ],
            className="header",
        ),
        # Content section
        html.Div(
            [
                html.Div(
                    [
                        dcc.Graph(id="line-graph", figure=create_figure()),
                    ],
                    className="content-left",
                ),
                html.Div(
                    [
                        create_dropdown(),
                        html.Div(
                            [
                                html.Div(
                                    [
                                        html.Div(
                                            "Current Loss Ratio", className="title"
                                        ),
                                        html.Div(
                                            id="overview-curr-lossRatio",
                                            className="data-value-primary",
                                        ),
                                        html.Div(
                                            [
                                                html.Div(
                                                    "Net Contribution",
                                                    className="title",
                                                ),
                                                html.Div(
                                                    id="overview-curr-cont",
                                                    className="data-value-secondary",
                                                ),
                                            ],
                                            className="data-card-secondary",
                                        ),
                                        html.Div(
                                            [
                                                html.Div(
                                                    "Incurred Claims", className="title"
                                                ),
                                                html.Div(
                                                    id="overview-curr-claim",
                                                    className="data-value-secondary",
                                                ),
                                            ],
                                            className="data-card-secondary",
                                        ),
                                        dcc.Graph(
                                            id="overview-fan-chart",
                                        ),
                                    ],
                                    className="data-card-primary",
                                ),
                                html.Div(
                                    [
                                        html.Div(
                                            "3-Yr Accummulated Loss Ratio",
                                            className="title",
                                        ),
                                        dcc.Graph(
                                            id="overview-bullet-accum-lossRatio",
                                        ),
                                        html.Div(
                                            [
                                                html.Div(
                                                    "Net Contribution",
                                                    className="title",
                                                ),
                                                html.Div(
                                                    id="overview-accum-cont",
                                                    className="data-value-secondary",
                                                ),
                                            ],
                                            className="data-card-secondary",
                                        ),
                                        html.Div(
                                            [
                                                html.Div(
                                                    "Incurred Claims", className="title"
                                                ),
                                                html.Div(
                                                    id="overview-accum-claim",
                                                    className="data-value-secondary",
                                                ),
                                            ],
                                            className="data-card-secondary",
                                        ),
                                    ],
                                    className="data-card-primary",
                                ),
                                html.Div(
                                    [
                                        html.Div(
                                            id="overview-num-lives",
                                            className="data-value-primary",
                                        ),
                                        html.Div("Number of Lives", className="title"),
                                    ],
                                    className="data-card-primary",
                                ),
                                html.Div(
                                    [
                                        dcc.Graph(
                                            id="overview-avg-claim",
                                            config={"displayModeBar": False},
                                        ),
                                        html.Div(
                                            "Average Claim Size", className="title"
                                        ),
                                        dcc.Graph(
                                            id="overview-box-avg-claim",
                                            config={"displayModeBar": False},
                                        ),
                                    ],
                                    className="data-card-primary",
                                ),
                                html.Div(
                                    [
                                        html.Div(
                                            id="overview-reprice-date",
                                            className="data-value-primary",
                                        ),
                                        html.Div(
                                            "Last Repricing Date", className="title"
                                        ),
                                    ],
                                    className="data-card-primary",
                                ),
                                html.Div(
                                    [
                                        html.Div(
                                            id="overview-reprice-mnths",
                                            className="data-value-primary",
                                        ),
                                        html.Div(
                                            "Months Since Last Reprice",
                                            className="title",
                                        ),
                                    ],
                                    className="data-card-primary",
                                ),
                            ],
                            id="data-container",
                        ),
                    ],
                    className="content-right",
                ),
            ],
            className="content",
        ),
    ]
)

# --------------------------------------------------
# Callbacks
# --------------------------------------------------


@dash.callback(
    [
        Output("overview-curr-lossRatio", "children"),
        Output("overview-curr-cont", "children"),
        Output("overview-curr-claim", "children"),
        Output("overview-bullet-accum-lossRatio", "figure"),
        Output("overview-accum-cont", "children"),
        Output("overview-accum-claim", "children"),
        # Output("overview-selected-product-title", "children"),
        Output("overview-num-lives", "children"),
        Output("overview-avg-claim", "figure"),
        Output("overview-reprice-date", "children"),
        Output("overview-reprice-mnths", "children"),
        Output("overview-box-avg-claim", "figure"),
        Output("overview-fan-chart", "figure"),
    ],
    [Input("overview-selected-product-dropdown", "value")],
)
def update_data(selected_product):
    if selected_product is None:
        # Default to "All" product
        product_name = "All"
    else:
        # Get the product name from the clicked point
        product_name = selected_product

    # Filter the data for the selected product
    selected_data = data_table[data_table["Product"] == product_name]

    # Loss ratio data
    current_cont = selected_data.iloc[0]["Net Contribution"]
    accum_cont = selected_data.iloc[0]["3Yr Cum Net Contribution"]
    current_claim = selected_data.iloc[0]["Incurred Claim"]
    accum_claim = selected_data.iloc[0]["3Yr Cum Incurred Claim"]

    current_lossRatio = current_claim / current_cont
    accum_lossRatio = accum_claim / accum_cont

    # Update the title to display the selected product
    product_title = f"Product: {product_name}"

    num_lives = selected_data.iloc[0]["Number of Lives"]
    avg_claim = selected_data.iloc[0]["Average Claim Size"]
    reprice_date = selected_data.iloc[0]["Last Reprice Date"]
    reprice_mnths = selected_data.iloc[0]["Mths Since Reprice"]

    # Average claim indicator
    ind_avg_claim = create_delta_card(avg_claim, 200)

    # Format values
    formatted_num_lives = f"{int(num_lives):,}"
    formatted_current_cont = f"{current_cont/1_000_000:,.1f}m"
    formatted_current_claim = f"{current_claim/1_000_000:,.1f}m"
    formatted_accum_cont = f"{accum_cont/1_000_000:,.1f}m"
    formatted_accum_claim = f"{accum_claim/1_000_000:,.1f}m"
    formatted_curr_lossRatio = f"{current_lossRatio:,.1%}"
    formatted_accum_lossRatio = f"{accum_lossRatio:,.1%}"

    # Check if reprice_date is a string (blank or 'NA')
    if (
        pd.isna(reprice_date)
        or isinstance(reprice_date, str)
        and (reprice_date.strip() == "" or reprice_date.strip().upper() == "NA")
    ):
        formatted_reprice_date = "Not Available"
    else:
        formatted_reprice_date = reprice_date.strftime("%#d %B %Y")

    # Check if reprice_mnths is a string (blank or 'NA')
    if (
        pd.isna(reprice_mnths)
        or isinstance(reprice_mnths, str)
        and (reprice_mnths.strip() == "" or reprice_mnths.strip().upper() == "NA")
    ):
        formatted_reprice_mnths = "Not Available"
    else:
        formatted_reprice_mnths = (
            f"{int(reprice_mnths)}"  # Format as integer with no decimal points
        )

    # box plot
    box_avg_claim = create_box(avg_claim)

    # bullet chart
    bullet_accum_loss_ratio = create_bullet()

    # fan chart
    fan_chart = create_fan()

    return (
        formatted_curr_lossRatio,
        formatted_current_cont,
        formatted_current_claim,
        bullet_accum_loss_ratio,
        formatted_accum_cont,
        formatted_accum_claim,
        # product_title,
        formatted_num_lives,
        ind_avg_claim,
        formatted_reprice_date,
        formatted_reprice_mnths,
        box_avg_claim,
        fan_chart,
    )
