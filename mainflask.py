import pandas as pd
from flask import Flask, render_template_string

app = Flask(__name__)

# Sample file path (replace with your actual file path)
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
        .tabs button.active {
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
            position: sticky;
            top: 0;
            z-index: 1;
        }
        select {
            padding: 10px;
            font-size: 16px;
            border: 1px solid #ddd;
            border-radius: 5px;
            margin-bottom: 20px;
            min-width: 200px;
        }
        .filter-container {
            margin-bottom: 20px;
        }
        .no-results {
            text-align: center;
            padding: 20px;
            color: {{ stevens_color_secondary }};
            font-style: italic;
            display: none;
        }
    </style>
    <script>
        function openTab(evt, sheetName) {
            var tabcontent = document.getElementsByClassName("tabcontent");
            for (var i = 0; i < tabcontent.length; i++) {
                tabcontent[i].style.display = "none";
            }
            var tablinks = document.getElementsByClassName("tablinks");
            for (var i = 0; i < tablinks.length; i++) {
                tablinks[i].classList.remove("active");
            }
            document.getElementById(sheetName).style.display = "block";
            evt.currentTarget.classList.add("active");
            
            // Reset the filter when switching tabs
            var select = document.getElementById(sheetName + "-instructor");
            if (select) {
                select.value = "";
                filterByInstructor(sheetName);
            }
        }

        function filterByInstructor(sheetName) {
            var select = document.getElementById(sheetName + "-instructor");
            var instructor = select.value.toLowerCase();
            var tableContainer = document.getElementById(sheetName);
            var rows = tableContainer.getElementsByClassName("table-row");
            var visibleCount = 0;
            
            for (var i = 0; i < rows.length; i++) {
                var row = rows[i];
                var instructorCell = row.querySelector(".instructor-cell");
                
                if (instructorCell) {
                    var instructorText = instructorCell.textContent.toLowerCase();
                    if (instructor === "" || instructorText.includes(instructor)) {
                        row.style.display = "";
                        visibleCount++;
                    } else {
                        row.style.display = "none";
                    }
                }
            }
            
            // Show/hide no results message
            var noResults = tableContainer.querySelector(".no-results");
            if (noResults) {
                noResults.style.display = visibleCount === 0 ? "block" : "none";
            }
        }

        document.addEventListener("DOMContentLoaded", function() {
            // Open first tab by default
            var firstTab = document.querySelector(".tablinks");
            if (firstTab) {
                firstTab.click();
            }
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
            <div class="filter-container">
                <label for="{{ sheet_name }}-instructor">Filter by Instructor(s)/TA: </label>
                <select id="{{ sheet_name }}-instructor" onchange="filterByInstructor('{{ sheet_name }}')">
                    <option value="">All Instructors</option>
                    {% for instructor in instructors[sheet_name] %}
                        {% if instructor %}
                            <option value="{{ instructor }}">{{ instructor }}</option>
                        {% endif %}
                    {% endfor %}
                </select>
            </div>
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
                            <tr class="table-row">
                                {% for cell in row %}
                                    <td class="{% if loop.index0 == instructor_index[sheet_name] %}instructor-cell{% endif %}">
                                        {{ cell }}
                                    </td>
                                {% endfor %}
                            </tr>
                        {% endfor %}
                    </tbody>
                </table>
                <div class="no-results">No results found for the selected instructor.</div>
            </div>
        </div>
        {% endfor %}
    </div>
</body>
</html>
'''

@app.route("/", methods=["GET", "POST"])
def index():
    heading = "Department of Systems and Enterprises Load on Instructors"
    processed_data = {}
    columns = {}
    instructors = {}
    instructor_index = {}

    for sheet_name, df in data_frames.items():
        # Clean the dataframe
        df_cleaned = df.fillna('')
        
        # Convert all columns to string and clean them
        for col in df_cleaned.columns:
            df_cleaned[col] = df_cleaned[col].astype(str).str.strip()
        
        processed_data[sheet_name] = df_cleaned.values.tolist()
        columns[sheet_name] = df_cleaned.columns.tolist()

        # Handle instructors
        instructor_col = 'Instructor(s)/Teaching Assistant'
        if instructor_col in df_cleaned.columns:
            instructor_index[sheet_name] = df_cleaned.columns.get_loc(instructor_col)
            # Get unique instructors, remove empty strings and sort
            unique_instructors = df_cleaned[instructor_col].unique()
            instructors[sheet_name] = sorted([instr for instr in unique_instructors if instr and instr != 'nan'])

    return render_template_string(
        html,
        heading=heading,
        data=processed_data,
        columns=columns,
        instructors=instructors,
        instructor_index=instructor_index,
        stevens_color_primary=stevens_color_primary,
        stevens_color_secondary=stevens_color_secondary
    )

if __name__ == "__main__":
    app.run(debug=True, port=5001)