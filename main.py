from datetime import datetime
import os
import argparse
from collections import defaultdict
import pandas as pd
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def setup_jinja_environment():
    current_dir = os.path.dirname(__file__)
    return Environment(
        loader=FileSystemLoader(current_dir),
        autoescape=select_autoescape(['html', 'xml'])
    )


def format_year(num: int, first: str = "год", second: str = "года", third: str = "лет") -> str:
    """Returns the correct year form based on the given number."""
    if 4 < num < 21:
        return third
    num %= 10
    if num == 1:
        return first
    if 1 < num < 5:
        return second
    return third


def load_wines_data(file_path: str) -> list:
    """Loads data from an Excel file and returns it as a list of dictionaries."""
    return pd.read_excel(
        io=file_path,
        sheet_name="Лист1",
        na_values=['N/A', 'NA'], keep_default_na=False
    ).to_dict(orient="records")


def categorize_wines(drinks_data: list) -> tuple:
    """Processes drinks data into categories and finds the lowest price."""
    all_drinks = defaultdict(list)
    lowest_price = None

    for drink in drinks_data:
        category = drink['Категория']
        all_drinks[category].append(drink)

        if lowest_price is None or drink["Цена"] < lowest_price:
            lowest_price = drink["Цена"]

    return all_drinks, lowest_price


def render_html(template, difference_in_years, formatted_year, drinks, lowest_price):
    """Renders HTML page using the template and provided data."""
    return template.render(
        date_of_establishment=difference_in_years,
        formatted_year=formatted_year,
        drinks=drinks,
        lowest_price=lowest_price
    )


def save_rendered_page(content: str, file_path: str):
    """Saves rendered HTML content to a file."""
    with open(file_path, 'w', encoding="utf8") as file:
        file.write(content)


def start_server():
    """Starts the HTTP server."""
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


def main():
    # Command-line argument parsing
    parser = argparse.ArgumentParser(description="Wine Data Processing")
    parser.add_argument("--path", default=os.getenv("WINE_DATA_PATH", "wine.xlsx"),
                        help="Path to the Excel file containing wine data.")
    args = parser.parse_args()

    wine_data_path = args.path

    # Setup and template rendering
    env = setup_jinja_environment()
    template = env.get_template("template.html")

    today = datetime.now().year
    year_of_establishment = 1920
    difference_in_years = today - year_of_establishment
    formatted_year = format_year(difference_in_years)

    all_wines_data = load_wines_data(wine_data_path)
    drinks, lowest_price = categorize_wines(all_wines_data)

    rendered_page = render_html(
        template, difference_in_years, formatted_year, drinks, lowest_price)
    save_rendered_page(rendered_page, os.path.join(
        os.path.dirname(__file__), "index.html"))

    start_server()


if __name__ == "__main__":
    main()
