from dash import Dash, html, dcc, Input, Output, dash_table
import dash
import plotly.graph_objs as go
import pandas as pd

# register page in directory
dash.register_page(__name__, path="/product")

# --------------------------------------------------
# Data Processing
# --------------------------------------------------
# Data read
data_chart = pd.read_excel("data.xlsx", sheet_name="Sheet1")
data_table = pd.read_excel("data.xlsx", sheet_name="Sheet2")

# Convert to percentage
data_chart[["21Q1", "22Q1", "Current"]] = data_chart[["21Q1", "22Q1", "Current"]] * 100


def create_figure(selected_products=None):

    # Define the data for the Plotly graph
    line_graph = go.Figure()

    # Add the trace for 'All' products to be prominent by default
    line_graph.add_trace(
        go.Scatter(
            x=["21Q1", "22Q1", "Current"],
            y=data_chart[data_chart["Product"] == "All"].iloc[0, 1:],
            mode="lines+markers",
            name="All",
            line=dict(width=3),  # Make this line thicker to be more prominent
            opacity=1.0,  # Full opacity
            text=["All"] * 3,  # Custom hover text
            hoverinfo="text+y",  # Show custom text and y value on hover
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

    for product in products:
        line_graph.add_trace(
            go.Scatter(
                x=["21Q1", "22Q1", "Current"],
                y=data_chart[data_chart["Product"] == product].iloc[0, 1:],
                mode="lines+markers",
                name=product,
                line=dict(width=2),  # Normal width
                opacity=0.3,
                text=[product] * 3,  # Custom hover text
                hoverinfo="text+y",  # Show custom text and y value on hover
            )
        )

    # Set layout for the graph
    line_graph.update_layout(
        title=dict(
            text="Loss Ratio Development vs Previous Studies",
            font=dict(
                family="var(--font-family)",  # Use the custom font defined in CSS
                weight="bold",  # Bold the title
            ),
        ),
        xaxis=dict(title="Study Period"),
        yaxis=dict(title="Loss Ratio (%)"),
        legend=dict(
            orientation="h",  # Horizontal legend
            x=-0.05,  # Center the legend horizontally
            y=1.05,  # Position the legend below the title
            xanchor="left",  # Align legend horizontally
            yanchor="bottom",  # Align legend vertically
        ),
        font=dict(
            family="var(--font-family)",  # Apply the custom font to all chart elements
        ),
        template="plotly_white",
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
                html.Div("Co.", className="logo"),
                html.Div(
                    [
                        html.Div(
                            "Claim Experience Study - Medical Products",
                            className="title-text",
                        ),
                        html.Div("Q1 2024", className="title-date"),
                    ],
                    className="title",
                ),
            ],
            className="header",
        ),
        # Menu section
        html.Div(
            html.Ul(
                [
                    html.Li(dcc.Link("Overview", href="/")),
                    html.Li(
                        dcc.Link("Product", href="/product"),
                        className="selected-menu-item",
                    ),
                ]
            ),
            className="menu",
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
                        html.Div(
                            id="selected-product-title",
                            className="selected-product-title",
                        ),
                        html.Br(),
                        dash_table.DataTable(
                            id="data-table",
                            columns=[
                                {
                                    "name": "Loss Ratio Component",
                                    "id": "Loss Ratio Component",
                                },
                                {
                                    "name": "Current",
                                    "id": "Current",
                                    "type": "numeric",
                                    "format": {"specifier": ",.2f"},
                                },
                                {
                                    "name": "3-Yr Cumulative",
                                    "id": "3-Yr Cumulative",
                                    "type": "numeric",
                                    "format": {"specifier": ",.2f"},
                                },
                            ],
                            data=[],  # Will be filled by callback
                            style_table={
                                "width": "80%",
                            },
                            style_cell={
                                "font-family": "var(--font-family)",
                                "minWidth": "0px",  # Set minimum width to 0
                                "width": "auto",  # Auto width to fit the content
                                "maxWidth": "150px",  # Set a reasonable maximum width
                                "whiteSpace": "normal",  # Allow content to wrap
                            },
                            style_header={
                                "margin": "auto",
                                "padding-left": "5px",
                                "background-color": "var(--theme-color)",
                                "font-weight": "bold",
                                "color": "white",
                                "textAlign": "left",
                            },
                            style_data_conditional=[
                                # Style for Net Contribution and Incurred Claim (first two rows)
                                {
                                    "if": {
                                        "filter_query": '{Loss Ratio Component} = "Net Contribution" || {Loss Ratio Component} = "Incurred Claim"'
                                    },
                                    "textAlign": "right",
                                    "padding": "5px 10px",
                                },
                                # Style for Loss Ratio (last row)
                                {
                                    "if": {
                                        "filter_query": '{Loss Ratio Component} = "Loss Ratio"'
                                    },
                                    "fontWeight": "bold",
                                    "border-top": "2px solid black",
                                    "border-bottom": "2px solid black",
                                    "padding": "5px 10px",
                                },
                                # Style for Loss Ratio Component column (first column)
                                {
                                    "if": {"column_id": "Loss Ratio Component"},
                                    "textAlign": "left",
                                    "fontWeight": "bold",
                                },
                            ],
                        ),
                        html.Br(),
                        html.Div(
                            [
                                html.Div(
                                    [
                                        html.Div("Number of Lives"),
                                        html.Div(
                                            id="num-lives",
                                            className="data-value-child",
                                        ),
                                    ],
                                    className="data-child",
                                ),
                                html.Div(
                                    [
                                        html.Div("Average Claim Size"),
                                        html.Div(
                                            id="avg-claim",
                                            className="data-value-child",
                                        ),
                                    ],
                                    className="data-child",
                                ),
                                html.Div(
                                    [
                                        html.Div("Last Repricing Date"),
                                        html.Div(
                                            id="reprice-date",
                                            className="data-value-child",
                                        ),
                                    ],
                                    className="data-child",
                                ),
                                html.Div(
                                    [
                                        html.Div("Months Since Last Reprice"),
                                        html.Div(
                                            id="reprice-mnths",
                                            className="data-value-child",
                                        ),
                                    ],
                                    className="data-child",
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
        Output("data-table", "data"),
        Output("selected-product-title", "children"),
        Output("num-lives", "children"),
        Output("avg-claim", "children"),
        Output("reprice-date", "children"),
        Output("reprice-mnths", "children"),
    ],
    [Input("line-graph", "clickData")],
)
def update_table(clickData):
    if clickData is None:
        # Default to "All" product
        product_name = "All"
    else:
        # Get the product name from the clicked point
        product_name = clickData["points"][0]["text"]

    # Filter the data for the selected product
    selected_data = data_table[data_table["Product"] == product_name]

    # Prepare data for the table
    table_data = [
        {
            "Loss Ratio Component": "Net Contribution",
            "Current": selected_data.iloc[0]["Net Contribution"],
            "3-Yr Cumulative": selected_data.iloc[0]["3Yr Cum Net Contribution"],
        },
        {
            "Loss Ratio Component": "Incurred Claim",
            "Current": selected_data.iloc[0]["Incurred Claim"],
            "3-Yr Cumulative": selected_data.iloc[0]["3Yr Cum Incurred Claim"],
        },
        {
            "Loss Ratio Component": "Loss Ratio",
            "Current": f'{selected_data.iloc[0]["Loss Ratio"]:.2%}',
            "3-Yr Cumulative": f'{selected_data.iloc[0]["Loss Ratio.1"]:.2%}',
        },
    ]

    # Update the title to display the selected product
    product_title = f"Product: {product_name}"

    num_lives = selected_data.iloc[0]["Number of Lives"]  # Number of lives
    avg_claim = selected_data.iloc[0]["Average Claim Size"]  # Average Claim Size
    reprice_date = selected_data.iloc[0]["Last Reprice Date"]  # Repricing Date
    reprice_mnths = selected_data.iloc[0][
        "Mths Since Reprice"
    ]  # Months from Repricing Date

    # Format values
    formatted_num_lives = (
        f"{int(num_lives):,}"  # Format as integer with comma separator
    )
    formatted_avg_claim = (
        f"{avg_claim:,.2f}"  # Format as float with 2 decimal points and comma separator
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

    return (
        table_data,
        product_title,
        formatted_num_lives,
        formatted_avg_claim,
        formatted_reprice_date,
        formatted_reprice_mnths,
    )
