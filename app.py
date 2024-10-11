from flask import Flask, request, render_template, redirect, url_for, send_file
import pandas as pd
import matplotlib.pyplot as plt

app = Flask(__name__)

# Load the Excel file and define the global DataFrame
df = pd.read_excel('student_data.xlsx')

def delete_student(enrollment_id):
    global df
    # Check if the Enrollment ID exists
    if enrollment_id in df['Enrollment ID'].values:
        # Drop the row with the given Enrollment ID
        df.drop(df[df['Enrollment ID'] == enrollment_id].index, inplace=True)
        print(f"Student with Enrollment ID {enrollment_id} has been deleted.")
        save_data()
    else:
        print(f"Student with Enrollment ID {enrollment_id} not found.")

### Generate a Pie Chart with Three Ranges ###
def pie_chart(column):
    global df
    # Validate the column name
    if column not in ['CW marks', 'SW marks', 'Attendance (%)']:
        print(f"Invalid column name '{column}'. Please use 'CW marks', 'SW marks', or 'Attendance (%)'.")
        return

    # Define the ranges with only three bins
    if column in ['CW marks', 'SW marks']:
        bins = [0, 10, 20, 30]  # Three ranges for CW and SW marks
        labels = ['Low (0-10)', 'Medium (11-20)', 'High (21-30)']
    elif column == 'Attendance (%)':
        bins = [0, 60, 75, 100]  # Three ranges for Attendance
        labels = ['Low (0-60%)', 'Medium (61-75%)', 'High (76-100%)']

    # Cut the data into categories based on the defined ranges
    df['Range'] = pd.cut(df[column], bins=bins, labels=labels, include_lowest=True)
    
    # Count the number of students in each range
    range_counts = df['Range'].value_counts()

    # Create a pie chart
    plt.figure(figsize=(10, 6))
    plt.pie(range_counts, labels=range_counts.index, autopct='%1.1f%%', startangle=90)
    plt.title(f'Pie Chart of {column} (Grouped into 3 Ranges)')
    plt.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

    # Save the pie chart as an image file
    chart_file = f'pie_chart_{column}.png'
    plt.savefig(chart_file)
    plt.close()  # Close the plot
    print(f"Pie chart saved as '{chart_file}'.")

    # Save the data used for the pie chart to a new Excel file
    pie_data_file = f'pie_chart_data_{column}.xlsx'
    range_counts_df = range_counts.reset_index()
    range_counts_df.columns = [column, 'Count']  # Rename columns for clarity
    range_counts_df.to_excel(pie_data_file, index=False)
    print(f"Pie chart data saved to '{pie_data_file}'.")


### Flask Routes ###
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/add', methods=['GET', 'POST'])
def add_student_page():
    if request.method == 'POST':
        global df
        enrollment_id = request.form['enrollment_id']
        name = request.form['name']
        cw_marks = int(request.form['cw_marks'])
        sw_marks = int(request.form['sw_marks'])
        attendance = int(request.form['attendance'])
        
        # Check if the Enrollment ID already exists
        if enrollment_id in df['Enrollment ID'].values:
            return f"Enrollment ID {enrollment_id} already exists. Cannot add student."

        new_student = pd.DataFrame({
            'Enrollment ID': [enrollment_id],
            'Name': [name],
            'CW marks': [cw_marks],
            'SW marks': [sw_marks],
            'Attendance (%)': [attendance]
        })

        df = pd.concat([df, new_student], ignore_index=True)
        df.sort_values(by='Enrollment ID', inplace=True)
        save_data()
        
        return redirect(url_for('index'))

    return render_template('add_student.html')

@app.route('/modify', methods=['GET', 'POST'])
def modify_student_page():
    if request.method == 'POST':
        global df
        enrollment_id = request.form['enrollment_id']
        column = request.form['column']
        new_value = request.form['new_value']

        global df
        # Check if the Enrollment ID exists
        if enrollment_id in df['Enrollment ID'].values:
            # Validate the column name
            if column in ['CW marks', 'SW marks', 'Attendance (%)']:
                # Update the specified column for the given Enrollment ID
                df.loc[df['Enrollment ID'] == enrollment_id, column] = new_value
                print(f"Student with Enrollment ID {enrollment_id} updated: {column} set to {new_value}.")
                save_data()
            else:
                print(f"Invalid column name '{column}'. Please use 'CW marks', 'SW marks', or 'Attendance (%)'.")
        else:
            print(f"Student with Enrollment ID {enrollment_id} not found.")
        return redirect(url_for('index'))

        
        # modify_student(enrollment_id, column, new_value)
        # return redirect(url_for('index'))

    return render_template('modify_student.html')

@app.route('/delete', methods=['GET', 'POST'])
def delete_student_page():
    if request.method == 'POST':
        global df
        enrollment_id = request.form['enrollment_id'] 
        # delete_student(enrollment_id)
        global df
        # Check if the Enrollment ID exists
        if enrollment_id in df['Enrollment ID'].values:
            # Drop the row with the given Enrollment ID
            df.drop(df[df['Enrollment ID'] == enrollment_id].index, inplace=True)
            print(f"Student with Enrollment ID {enrollment_id} has been deleted.")
            save_data()
        else:
            print(f"Student with Enrollment ID {enrollment_id} not found.")
            return redirect(url_for('index'))
        
    return render_template('delete_student.html')

@app.route('/extract', methods=['GET', 'POST'])
def extract_data_page():
    if request.method == 'POST':
        global df
        start_id = request.form['start_id']
        end_id = request.form['end_id']
        
        global df
        # Ensure start_id and end_id are valid and in the correct format
        if start_id > end_id:
            print("Invalid range: start_id should be less than or equal to end_id.")
            return
        
        # Filter the DataFrame for Enrollment IDs in the specified range
        filtered_data = df[(df['Enrollment ID'] >= start_id) & (df['Enrollment ID'] <= end_id)]
        
        if not filtered_data.empty:
            # Save extracted data to a new Excel file
            output_file = f'extracted_data_{start_id}_to_{end_id}.xlsx'
            filtered_data.to_excel(output_file, index=False)
            print(f"Extracted data saved to '{output_file}'.")
        else:
            print(f"No students found in the range {start_id} to {end_id}.")

        return redirect(url_for('index'))

    return render_template('extract_data.html')

@app.route('/filter', methods=['GET', 'POST'])
def filter_data_page():
    if request.method == 'POST':
        global df
        column = request.form['column']
        min_value = int(request.form['min_value'])
        max_value = int(request.form['max_value'])
        
        global df
        # Validate the column name
        if column not in ['CW marks', 'SW marks', 'Attendance (%)']:
            print(f"Invalid column name '{column}'. Please use 'CW marks', 'SW marks', or 'Attendance (%)'.")
            return
        
        # Filter the DataFrame based on the column and range
        filtered_data = df[(df[column] >= min_value) & (df[column] <= max_value)]
        
        if not filtered_data.empty:
            # Save filtered data to a new Excel file
            output_file = f'filtered_data_{column}_{min_value}_to_{max_value}.xlsx'
            filtered_data.to_excel(output_file, index=False)
            print(f"Filtered data based on {column} ({min_value} to {max_value}) saved to '{output_file}'.")
        else:
            print(f"No students found in the range {min_value} to {max_value} for {column}.")

        return redirect(url_for('index'))

    return render_template('filter_data.html')

@app.route('/pie_chart', methods=['GET', 'POST'])
def generate_pie_chart_page():
    if request.method == 'POST':
        column = request.form['column']
        
        global df
        # Validate the column name
        if column not in ['CW marks', 'SW marks', 'Attendance (%)']:
            print(f"Invalid column name '{column}'. Please use 'CW marks', 'SW marks', or 'Attendance (%)'.")
            return render_template('pie_chart.html')  # Return the form page again with the error message

        # Convert the column to numeric, coercing invalid values to NaN
        df[column] = pd.to_numeric(df[column], errors='coerce')

        # Drop rows where the column is NaN (invalid values)
        df = df.dropna(subset=[column])

        # Define the ranges
        if column in ['CW marks', 'SW marks']:
            bins = [0, 10, 20, 30]  # Define custom ranges for CW and SW marks (e.g., Low, Medium, High)
            labels = ['Low (0-10)', 'Medium (11-20)', 'High (21-30)']
        elif column == 'Attendance (%)':
            bins = [0, 50, 75, 100]  # Define custom ranges for Attendance (e.g., Low, Medium, High)
            labels = ['Low (0-50%)', 'Medium (51-75%)', 'High (76-100%)']

        # Cut the data into categories based on the defined ranges
        df['Range'] = pd.cut(df[column], bins=bins, labels=labels, include_lowest=True)

        # Count the number of students in each range
        range_counts = df['Range'].value_counts()

        # Create a pie chart
        plt.figure(figsize=(10, 6))
        plt.pie(range_counts, labels=range_counts.index, autopct='%1.1f%%', startangle=90)
        plt.title(f'Pie Chart of {column} (Grouped into 3 Ranges)')
        plt.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

        # Save the pie chart as an image file
        chart_file = f'pie_chart_{column}.png'
        plt.savefig(chart_file)
        plt.close()  # Close the plot
        print(f"Pie chart saved as '{chart_file}'.")

        # Save the data used for the pie chart to a new Excel file
        pie_data_file = f'pie_chart_data_{column}.xlsx'
        range_counts_df = range_counts.reset_index()
        range_counts_df.columns = [column, 'Count']  # Rename columns for clarity
        range_counts_df.to_excel(pie_data_file, index=False)
        print(f"Pie chart data saved to '{pie_data_file}'.")

        return redirect(url_for('index'))

    # Render the pie chart form if it's a GET request
    return render_template('pie_chart.html')



@app.route('/display', methods=['GET'])
def display_data():
    global df
    # Assuming you have a function to retrieve all student data
    students = df.to_dict(orient='records')  # Convert DataFrame to dictionary
    return render_template('display_data.html', students=students)

def save_data():
    df.to_excel('student_data_updated.xlsx', index=False)
    print("Data saved to 'student_data_updated.xlsx'.")

if __name__ == '__main__':
    app.run(debug=True)
