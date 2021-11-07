import datetime
import pandas
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    current_date = datetime.datetime.today()
    past_date = datetime.datetime(year=1920, month=11, day=22)
    delta = (current_date - past_date).days
    years = delta // 365

    df = pandas.read_excel(
        io='wine3.xlsx',
        sheet_name="Лист1",
        na_values=['nan', None], keep_default_na=False
    )

    items = defaultdict(list)

    for item in df.to_dict(orient='records'):
        items[item['Категория']].append(item)

    rendered_page = template.render(
        years=years, items=dict(sorted(items.items(), key=lambda x: x[0]))
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
