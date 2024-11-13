import pandas as pd
import numpy as np
from flask import Flask, render_template_string

app = Flask(__name__)

file_path = 'combined_courses_file_d.xlsx'
data_frames = pd.read_excel(file_path, sheet_name=None)

stevens_color_primary = "#98002E"
stevens_color_secondary = "#555555"

html = '''
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>{{ heading }}</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f8f9fa;
            color: {{ stevens_color_secondary }};
        }
        .container {
            max-width: 1100px;
            margin: 20px auto;
            padding: 20px;
        }
        .heading {
            text-align: center;
            color: {{ stevens_color_primary }};
        }
        .tabs {
            display: flex;
            justify-content: space-around;
            background-color: {{ stevens_color_primary }};
            padding: 10px;
            border-radius: 5px;
            margin-bottom: 20px;
        }
        .tabs button {
            background: none;
            color: white;
            border: none;
            cursor: pointer;
            font-size: 16px;
            padding: 10px 15px;
        }
        .tabs button:hover {
            background-color: {{ stevens_color_secondary }};
        }
        .tabcontent {
            display: none;
            padding: 20px;
            border: 1px solid #ddd;
            border-radius: 5px;
            background-color: white;
        }
        .table-container {
            overflow-x: auto;
            max-height: 600px;
            overflow-y: auto;
            margin-top: 20px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th, td {
            padding: 8px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: {{ stevens_color_primary }};
            color: white;
        }
    </style>
    <script>
        function openTab(evt, sheetName) {
            var i, tabcontent, tablinks;
            tabcontent = document.getElementsByClassName("tabcontent");
            for (i = 0; i < tabcontent.length; i++) {
                tabcontent[i].style.display = "none";
            }
            tablinks = document.getElementsByClassName("tablinks");
            for (i = 0; i < tablinks.length; i++) {
                tablinks[i].className = tablinks[i].className.replace(" active", "");
            }
            document.getElementById(sheetName).style.display = "block";
            evt.currentTarget.className += " active";
        }
        document.addEventListener("DOMContentLoaded", function() {
            document.getElementsByClassName("tablinks")[0].click();
        });
    </script>
</head>
<body>
    <div class="container">
        <h1 class="heading">{{ heading }}</h1>
        <div class="tabs">
            {% for sheet_name in data.keys() %}
                <button class="tablinks" onclick="openTab(event, '{{ sheet_name }}')">{{ sheet_name }}</button>
            {% endfor %}
        </div>
        {% for sheet_name, rows in data.items() %}
        <div id="{{ sheet_name }}" class="tabcontent">
            <h2>{{ sheet_name }}</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            {% for col in columns[sheet_name] %}
                                <th>{{ col }}</th>
                            {% endfor %}
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in rows %}
                            <tr>
                                {% for cell in row %}
                                    <td>{{ cell if cell != 'nan' else '' }}</td>
                                {% endfor %}
                            </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
        {% endfor %}
    </div>
</body>
</html>
'''

@app.route("/")
def index():
    heading = "Department of Systems and Enterprises Load on Instructors"
    processed_data = {}
    columns = {}

    for sheet_name, df in data_frames.items():
        df_cleaned = df.fillna('')
        processed_data[sheet_name] = df_cleaned.astype(str).replace('nan', '').values.tolist()
        columns[sheet_name] = df_cleaned.columns.tolist()

    return render_template_string(
        html,
        heading=heading,
        data=processed_data,
        columns=columns,
        stevens_color_primary=stevens_color_primary,
        stevens_color_secondary=stevens_color_secondary
    )

if __name__ == "__main__":
    app.run(port=5001)
