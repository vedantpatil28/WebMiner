from flask import Flask, render_template, request, send_file
import requests
from bs4 import BeautifulSoup
import pandas as pd
import os

app = Flask(__name__)

# scrape tables and store in excel
def scrape_tables_from_web(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")
    
    # get all tables
    tables = soup.find_all("table")
    
    if not tables:
        return None
    
    # save multiple tables in sheets in one excel file
    excel_path = "scraped_data.xlsx"
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for index, table in enumerate(tables):
            rows = table.find_all("tr")
            
            # get table headers 
            headers = [th.text.strip() for th in rows[0].find_all("th")] if rows else []

            # get table rows
            data = []
            for row in rows[1:]:  # skip the header row
                columns = row.find_all(["td", "th"])
                row_data = [col.text.strip() for col in columns]
                data.append(row_data)

            # data -> frames
            df = pd.DataFrame(data, columns=headers if headers else None)
            
            # new sheet for each table
            sheet_name = f"Table_{index + 1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    return excel_path

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        url = request.form["url"]
        if url:
            excel_path = scrape_tables_from_web(url)
            if excel_path:
                return send_file(excel_path, as_attachment=True)
            else:
                return "No tabular data found on the provided URL."
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)