from dash import Dash, dcc, html
import dash

# Initialize the Dash app (only once, in the main app file)
app = Dash(__name__, use_pages=True)

# Main layout for the app
app.layout = html.Div(
    [dash.page_container]  # This will render the correct page based on the URL
)

# Run the app
if __name__ == "__main__":
    app.run_server(debug=True)
