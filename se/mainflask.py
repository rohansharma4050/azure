import pandas as pd
from flask import Flask, render_template_string

app = Flask(__name__)

# Sample file path (replace with your actual file path)
file_path = 'combined_courses_file_d.xlsx'
# Read Excel file with all string columns to prevent data type conversion issues
data_frames = pd.read_excel(file_path, sheet_name=None, keep_default_na=False, dtype=str)

stevens_color_primary = "#98002E"
stevens_color_secondary = "#555555"

@app.route("/", methods=["GET", "POST"])
def index():
    heading = "Department of Systems and Enterprises Load on Instructors"
    processed_data = {}
    columns = {}
    instructors = {}
    instructor_index = {}
    combined_index = {}
    building_index = {}

    for sheet_name, df in data_frames.items():
        # Ensure consistent column names and clean up whitespace
        df.columns = [col.strip() for col in df.columns]
        
        # Standardize the Combined Courses column name
        combined_cols = ['Combined Courses', 'Combined_Course']
        for col in combined_cols:
            if col in df.columns:
                df = df.rename(columns={col: 'Combined_Course'})
                break
        
        # Clean the dataframe
        df_cleaned = df.copy()
        
        # Convert all columns to string and clean them
        for col in df_cleaned.columns:
            df_cleaned[col] = df_cleaned[col].astype(str).str.strip()
            
        # Replace empty strings and 'nan' with None for better template handling
        df_cleaned = df_cleaned.replace({'': None, 'nan': None})
        
        # Ensure Combined_Course values are properly formatted
        if 'Combined_Course' in df_cleaned.columns:
            df_cleaned['Combined_Course'] = df_cleaned['Combined_Course'].apply(
                lambda x: x if x and x.lower() != 'nan' else None
            )
        
        # Store processed data
        processed_data[sheet_name] = df_cleaned.replace({None: ''}).values.tolist()
        columns[sheet_name] = df_cleaned.columns.tolist()

        # Handle instructors column
        instructor_col = 'Instructor(s)/Teaching Assistant'
        instructor_index[sheet_name] = df_cleaned.columns.get_loc(instructor_col) if instructor_col in df_cleaned.columns else -1
        if instructor_col in df_cleaned.columns:
            unique_instructors = sorted([
                instr for instr in df_cleaned[instructor_col].unique()
                if instr and str(instr).lower() != 'nan'
            ])
            instructors[sheet_name] = unique_instructors
        
        # Set column indices
        combined_index[sheet_name] = df_cleaned.columns.get_loc('Combined_Course') if 'Combined_Course' in df_cleaned.columns else -1
        building_index[sheet_name] = df_cleaned.columns.get_loc('Building/Room') if 'Building/Room' in df_cleaned.columns else -1

    return render_template_string(
        html,  # Using the same HTML template as before
        heading=heading,
        data=processed_data,
        columns=columns,
        instructors=instructors,
        instructor_index=instructor_index,
        combined_index=combined_index,
        building_index=building_index,
        stevens_color_primary=stevens_color_primary,
        stevens_color_secondary=stevens_color_secondary
    )

if __name__ == "__main__":
    app.run(debug=True, port=5001)