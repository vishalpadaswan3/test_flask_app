from flask import Flask, render_template, request
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta
import calendar
import pandas as pd

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        project_type = request.form['project_type']
        project_name = request.form['project_name']
        start_date_str = request.form['start_date']
        end_date_str = request.form['end_date']

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d")

            if start_date >= end_date:
                raise ValueError("End date should be after start date.")
        except ValueError as e:
            return render_template('index.html', error=str(e))

        stages = get_stages(project_type)
        project_timeline, excel_filename = create_project_timeline(stages, project_name, start_date, end_date)

        # Pass project details to the result page
        return render_template('result.html', project_type=project_type, project_name=project_name, start_date=start_date.strftime("%Y-%m-%d"), end_date=end_date.strftime("%Y-%m-%d"), project_timeline=project_timeline, excel_filename=excel_filename)

    return render_template('index.html')


def get_stages(project_type):
    minor_change_stages = [
        "Project kickoff",
        "Drawing and BOM",
        "QGCO",
        "Cable Sourcing",
        "Air gap analysis",
        "Customer C sample",
        "Customer approval for drawing",
        "D sample Production Release",
        "Cable PPAP",
        "QGC4",
        "SOP"
    ]
    adapt_project_stages = [
        "Project kickoff",
        "Design and Specification",
        "Prototype",
        "Production",
        "Testing",
        "Implementation",
        "Finalization"
    ]
    return minor_change_stages if project_type == '1' else adapt_project_stages

def create_project_timeline(stages, project_name, start_date, end_date):
    project_timeline = []
    current_stage_start = start_date

    for stage_name in stages:
        stage_duration = request.form[stage_name.lower().replace(" ", "_")]
        start_weeks, end_weeks = map(int, stage_duration.split('-'))

        current_stage_end = current_stage_start + timedelta(weeks=end_weeks)

        project_timeline.append({
            'Stage Name': stage_name,
            'Start Date': current_stage_start,  # Convert to datetime object
            'End Date': current_stage_end,  # Convert to datetime object
            'Duration (weeks)': f"{start_weeks}-{end_weeks}"
        })

        current_stage_start = current_stage_end

    df = pd.DataFrame(project_timeline)
    excel_filename = "project_timeline.xlsx"
    df.to_excel(excel_filename, index=False)

    display_calendar(start_date, end_date, project_timeline)

    return project_timeline, excel_filename


def display_calendar(start_date, end_date, project_timeline):
    fig, ax = plt.subplots(figsize=(12, 6))
    cal = calendar.Calendar()

    start_year, start_month = start_date.year, start_date.month
    end_year, end_month = end_date.year, end_date.month

    stage_months = {stage['Stage Name']: (stage['Start Date'].month, stage['End Date'].month) for stage in project_timeline}

    for stage_name, months in stage_months.items():
        start_month_num, end_month_num = months
        for month_num in range(start_month_num, end_month_num + 1):
            ax.plot(stage_name, month_num, marker='o', color='blue')

    ax.set_xticks(range(len(project_timeline)))
    ax.set_xticklabels([stage['Stage Name'] for stage in project_timeline], rotation=45, ha='right')
    ax.set_yticks(range(1, 13))
    ax.set_yticklabels(calendar.month_abbr[1:], rotation=45, ha='right')

    ax.set_title('Project Timeline')
    ax.set_xlabel('Stage Name')
    ax.set_ylabel('Month')
    plt.grid(True)

    plt.tight_layout()
    plt.savefig('static/calendar.png')

if __name__ == '__main__':
    app.run(debug=True)
