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
data_table_curr = pd.read_excel("data.xlsx", sheet_name="Sheet2")
data_table_prev = pd.read_excel("data.xlsx", sheet_name="Sheet3")

# Convert to percentage
data_chart[["21Q1", "22Q1", "Current"]] = data_chart[["21Q1", "22Q1", "Current"]] * 100


def get_data(products="All"):

    product_name = products

    # Filter the data for all medical product
    data_curr = data_table_curr[data_table_curr["Product"] == product_name]
    data_prev = data_table_prev[data_table_curr["Product"] == product_name]

    # Loss ratio data
    current_cont = data_curr.iloc[0]["Net Contribution"]
    prev_cont = data_prev.iloc[0]["Net Contribution"]
    current_claim = data_curr.iloc[0]["Incurred Claim"]
    prev_claim = data_prev.iloc[0]["Incurred Claim"]

    current_lossRatio = current_claim / current_cont
    prev_lossRatio = prev_claim / prev_cont

    # Other data
    current_num_lives = data_curr.iloc[0]["Number of Lives"]
    current_avg_claim = data_curr.iloc[0]["Average Claim Size"]
    prev_num_lives = data_prev.iloc[0]["Number of Lives"]
    prev_avg_claim = data_prev.iloc[0]["Average Claim Size"]
    reprice_date = data_curr.iloc[0]["Last Reprice Date"]
    reprice_mnths = data_curr.iloc[0]["Mths Since Reprice"]

    return (
        current_lossRatio,
        current_claim,
        current_cont,
        prev_lossRatio,
        prev_cont,
        prev_claim,
        current_num_lives,
        current_avg_claim,
        prev_num_lives,
        prev_avg_claim,
        reprice_mnths,
        reprice_date,
    )


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


def create_bullet(value):
    current_value = value * 100
    max_value = round(current_value, 0) + 15

    # Create a bullet chart using a bar to represent the current value
    fig = go.Figure(
        go.Indicator(
            mode="gauge+delta",
            value=current_value,
            gauge={
                "shape": "bullet",
                "axis": {
                    "range": [None, max_value],
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
                        "range": [100, max_value],
                        "color": "rgb(255, 83, 83)",
                    },
                ],
            },
            domain={"x": [0, 1], "y": [0.4, 1]},  # use y value to adjust height
            number={
                "suffix": "%",
                "font": {
                    "size": 34,
                    "color": "white",
                    "family": "Arial",
                    "weight": "bold",
                },
            },
            delta={
                "reference": 90,
                "suffix": "%",
                "valueformat": ".1f",
                "position": "right",
                "increasing": {"color": "red"},
                "decreasing": {"color": "green"},
                "font": {
                    "size": 16,
                },
            },
        )
    )

    fig.update_layout(
        paper_bgcolor="rgb(35, 36, 72)",
        plot_bgcolor="rgb(35, 36, 72)",
        margin=dict(l=60, r=60, t=15, b=0),
        height=100,
    )

    return fig


def create_box(val):
    # Generate random data with exponential distribution
    # np.random.seed(42)
    data = np.random.exponential(scale=val, size=30)

    # Add jitter to y-axis for better visibility
    jitter = np.random.normal(0, 0.1, len(data))

    # Rank the input value within the generated data
    rank_percentile = (np.sum(data < val) / len(data)) * 100

    # Calculate quartiles
    q1 = np.percentile(data, 25)
    median = np.median(data)
    q3 = np.percentile(data, 75)

    # Create the figure
    fig = go.Figure()

    # Add all data points as a scatter plot with jitter
    fig.add_trace(
        go.Scatter(
            x=data,
            y=jitter,
            mode="markers",
            marker=dict(color="lightgreen", size=5),
            name="Historical Value",
        )
    )

    # Highlight the specific current value
    fig.add_trace(
        go.Scatter(
            x=[val],
            y=[0],
            mode="markers",
            marker=dict(color="red", size=10),
            name="Current Value",
        )
    )

    # Add vertical lines for 1st quartile, median, and 3rd quartile
    fig.add_shape(
        type="line",
        x0=q1,
        y0=-0.5,
        x1=q1,
        y1=0.5,
        line=dict(color="rgb(174, 177, 210)", width=2),
    )  # 1st Quartile
    fig.add_shape(
        type="line",
        x0=median,
        y0=-0.5,
        x1=median,
        y1=0.5,
        line=dict(color="rgb(174, 177, 210)", width=2),
    )  # Median
    fig.add_shape(
        type="line",
        x0=q3,
        y0=-0.5,
        x1=q3,
        y1=0.5,
        line=dict(color="rgb(174, 177, 210)", width=2),
    )  # 3rd Quartile

    # Add annotation for the value point, showing the percentile
    percentile_text = f"{rank_percentile:.0f}th%"
    fig.add_annotation(
        x=val * 1.2,
        y=0.5,  # Place the annotation slightly above the point
        text=percentile_text,
        showarrow=True,
        arrowhead=2,
        ax=0,
        ay=-10,
        font=dict(
            color="rgb(174, 177, 210)",
            size=12,
        ),
    )

    # Update layout to remove axis titles and values, and adjust appearance
    fig.update_layout(
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
        showlegend=False,
        margin=dict(l=100, r=100, t=0, b=0),
        height=60,
        paper_bgcolor="rgb(35, 36, 72)",
        plot_bgcolor="rgb(35, 36, 72)",
    )

    # Return the chart and the rank percentile as a message
    percentile_text = f"{rank_percentile:.0f}th percentile"

    return fig, percentile_text


def create_delta_card(val, ref, format, suffix, direction="income", height="Medium"):

    # Map height labels to numeric values
    height_map = {"small": 40, "medium": 60, "big": 80}

    # Use the height map to get the correct size, default to "medium"
    height_used = height_map.get(height.lower(), 60)

    delta_card = go.Figure()

    if direction == "outgo":
        delta_color = {
            "increasing": {"color": "red"},  # Red for increase in outgo
            "decreasing": {"color": "green"},  # Green for decrease in outgo
        }
    else:  # Default is "income"
        delta_color = {
            "increasing": {"color": "green"},  # Green for increase in income
            "decreasing": {"color": "red"},  # Red for decrease in income
        }

    delta_card.add_trace(
        go.Indicator(
            mode="number+delta",
            value=val,
            number={
                "valueformat": format,
                "suffix": suffix,
                "font": {
                    "color": "rgb(213, 215, 224)",
                    "weight": "bold",
                    "family": "Arial, sans-serif",
                },
            },
            delta={
                "position": "right",
                "reference": ref,
                "valueformat": format,
                "suffix": suffix,
                **delta_color,
            },
            domain={"x": [0, 1], "y": [0, 1]},
        )
    )

    delta_card.update_layout(
        margin=dict(l=0, r=0, t=10, b=10),
        paper_bgcolor="rgb(35, 36, 72)",
        plot_bgcolor="rgb(35, 36, 72)",
        height=height_used,
    )

    return delta_card


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
        html.Div("Medical Loss Ratio", className="title page"),
        # Content section
        html.Div(
            [
                html.Div(
                    [
                        html.Div(
                            [
                                html.Div(
                                    [
                                        html.Div(
                                            "Current Loss Ratio", className="title"
                                        ),
                                        dcc.Graph(
                                            id="curr-lossRatio",
                                            config={"displayModeBar": False},
                                        ),
                                        dcc.Graph(
                                            id="bullet-curr-lossRatio",
                                        ),
                                        dcc.Graph(
                                            id="box-curr-lossRatio",
                                            config={"displayModeBar": False},
                                        ),
                                    ],
                                    className="data-card-primary",
                                ),
                                html.Div(
                                    [
                                        html.Div(
                                            [
                                                html.Div(
                                                    "Incurred Claims", className="title"
                                                ),
                                                dcc.Graph(
                                                    id="curr-claim",
                                                    config={"displayModeBar": False},
                                                ),
                                                html.Div(
                                                    [
                                                        dcc.Graph(
                                                            id="box-curr-claim",
                                                            config={
                                                                "displayModeBar": False
                                                            },
                                                        ),
                                                    ]
                                                ),
                                            ],
                                            className="data-card-primary",
                                        ),
                                        html.Div(
                                            [
                                                html.Div(
                                                    "Net Contribution",
                                                    className="title",
                                                ),
                                                dcc.Graph(
                                                    id="curr-cont",
                                                    config={"displayModeBar": False},
                                                ),
                                                dcc.Graph(
                                                    id="box-curr-cont",
                                                    config={"displayModeBar": False},
                                                ),
                                            ],
                                            className="data-card-primary",
                                        ),
                                    ],
                                    className="container-flex-col",
                                ),
                            ],
                            className="container-flex-row",
                        ),
                        html.Div(
                            [
                                html.Div(
                                    [
                                        html.Div(
                                            "Projected loss ratio", className="title"
                                        ),
                                        dcc.Graph(
                                            id="fan-chart",
                                        ),
                                    ],
                                    className="data-card-primary",
                                ),
                                html.Div(
                                    [
                                        html.Div(
                                            [
                                                html.Div(
                                                    "Number of Claims",
                                                    className="title",
                                                ),
                                                dcc.Graph(
                                                    id="num-claim",
                                                    config={"displayModeBar": False},
                                                ),
                                                dcc.Graph(
                                                    id="box-curr-num-claim",
                                                    config={"displayModeBar": False},
                                                ),
                                            ],
                                            className="data-card-primary",
                                        ),
                                        html.Div(
                                            [
                                                html.Div(
                                                    "Average Claim Size",
                                                    className="title",
                                                ),
                                                dcc.Graph(
                                                    id="avg-curr-claim",
                                                    config={"displayModeBar": False},
                                                ),
                                                dcc.Graph(
                                                    id="box-curr-avg-claim",
                                                    config={"displayModeBar": False},
                                                ),
                                            ],
                                            className="data-card-primary",
                                        ),
                                    ],
                                    className="container-flex-col",
                                ),
                            ],
                            className="container-flex-row",
                        ),
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
                                        html.Div("Number of Lives", className="title"),
                                    ],
                                    className="data-card-primary",
                                ),
                                html.Div(
                                    [
                                        html.Div(
                                            "Average Claim Size", className="title"
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
        # all loss ratios
        Output("curr-lossRatio", "figure"),
        Output("bullet-curr-lossRatio", "figure"),
        Output("box-curr-lossRatio", "figure"),
        # all contribution
        Output("curr-cont", "figure"),
        Output("box-curr-cont", "figure"),
        # all claims
        Output("curr-claim", "figure"),
        Output("box-curr-claim", "figure"),
        # all number of claims
        Output("num-claim", "figure"),
        Output("box-curr-num-claim", "figure"),
        # all average claims
        Output("avg-curr-claim", "figure"),
        Output("box-curr-avg-claim", "figure"),
        # all others
        Output("overview-reprice-date", "children"),
        Output("overview-reprice-mnths", "children"),
        Output("fan-chart", "figure"),
    ],
    [Input("overview-selected-product-dropdown", "value")],
)
def update_data(selected_product):

    # get all medical product data
    (
        current_lossRatio,
        current_claim,
        current_cont,
        prev_lossRatio,
        prev_cont,
        prev_claim,
        current_num_claim,
        current_avg_claim,
        prev_num_claim,
        prev_avg_claim,
        reprice_mnths,
        reprice_date,
    ) = get_data("All")

    # get all medical product data
    (
        selected_current_lossRatio,
        selected_current_claim,
        selected_current_cont,
        selected_prev_lossRatio,
        selected_prev_cont,
        selected_prev_claim,
        selected_current_num_claim,
        selected_current_avg_claim,
        selected_prev_num_claim,
        selected_prev_avg_claim,
        selected_reprice_mnths,
        selected_reprice_date,
    ) = get_data(selected_product)

    # Delta cards
    ind_current_lossRatio = create_delta_card(
        current_lossRatio * 100,
        prev_lossRatio * 100,
        ".1f",
        "%",
        "outgo",
        "big",
    )
    ind_curr_claim = create_delta_card(
        current_claim / 1_000_000, prev_claim / 1_000_000, ".1f", "m", "outgo"
    )
    ind_curr_cont = create_delta_card(
        current_cont / 1_000_000, prev_cont / 1_000_000, ".1f", "m", "income"
    )
    ind_curr_avg_claim = create_delta_card(
        current_avg_claim, prev_avg_claim, ".1f", "", "outgo"
    )
    ind_curr_num_claim = create_delta_card(
        current_num_claim, prev_num_claim, ",.0f", "", "outgo"
    )

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
    box_curr_lossRatio, _ = create_box((current_lossRatio - prev_lossRatio) * 100)
    box_curr_claim, _ = create_box((current_claim - prev_claim) / 1_000_000)
    box_curr_cont, _ = create_box((current_cont - prev_cont) / 1_000_000)
    box_curr_num_claim, _ = create_box((current_num_claim - prev_num_claim) / 1_000_000)
    box_curr_avg_claim, _ = create_box((current_avg_claim - prev_avg_claim) / 1_000_000)

    # bullet chart
    bullet_curr_loss_ratio = create_bullet(current_lossRatio)

    # fan chart
    fan_chart = create_fan()

    return (
        # all loss ratios
        ind_current_lossRatio,
        bullet_curr_loss_ratio,
        box_curr_lossRatio,
        # all contributions
        ind_curr_cont,
        box_curr_cont,
        # all claims
        ind_curr_claim,
        box_curr_claim,
        # all num of claims
        ind_curr_num_claim,
        box_curr_num_claim,
        # all average claim
        ind_curr_avg_claim,
        box_curr_avg_claim,
        # all others
        formatted_reprice_date,
        formatted_reprice_mnths,
        fan_chart,
    )
