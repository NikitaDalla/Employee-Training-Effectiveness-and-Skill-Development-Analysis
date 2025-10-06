import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html

# ✅ 1. Load Excel Data
excel_path = r"C:\Users\Dell\OneDrive\Desktop\coding\Python\project\Employee_Training_Data_v2.xlsx"

employees = pd.read_excel(excel_path, sheet_name="Employees")
trainings = pd.read_excel(excel_path, sheet_name="Trainings")
performance = pd.read_excel(excel_path, sheet_name="Performance")

# ✅ 2. Clean Column Names
employees.columns = employees.columns.str.strip()
trainings.columns = trainings.columns.str.strip()
performance.columns = performance.columns.str.strip()

# ✅ 3. Merge Data
# Merge performance with employees on Employee ID
merged_data = performance.merge(employees, on="Employee ID", how="left")

# Merge with trainings on Employee ID
merged_data = merged_data.merge(trainings, on="Employee ID", how="left")

# ✅ 4. Create Skill Improvement Column
merged_data["Skill Improvement"] = merged_data["Skill Level After"] - merged_data["Skill Level Before"]

# ✅ 5. Create Charts

# Bar Chart – Average Skill Improvement by Department
fig_bar = px.bar(
    merged_data.groupby("Department", as_index=False)["Skill Improvement"].mean(),
    x="Department",
    y="Skill Improvement",
    title="Average Skill Improvement by Department"
)

# Pie Chart – Training Completion Status
fig_pie = px.pie(
    trainings,
    names="Completion Status",  # correct column name
    title="Training Completion Status"
)

# Line Chart – Average Skill Improvement Over Time
fig_line = px.line(
    merged_data.groupby("Evaluation Date", as_index=False)["Skill Improvement"].mean(),
    x="Evaluation Date",
    y="Skill Improvement",
    title="Skill Improvement Trend Over Time"
)

# Heatmap – Department vs Skill Name Skill Improvement
fig_heatmap = px.density_heatmap(
    merged_data,
    x="Department",
    y="Skill Name",
    z="Skill Improvement",
    color_continuous_scale="Blues",
    title="Heatmap of Skill Improvement"
)

# Stacked Bar – Count of Trainings by Department & Completion Status
fig_stack = px.bar(
    merged_data,  # Use merged_data to include Department
    x="Department",
    color="Completion Status",
    title="Trainings by Department and Completion Status",
    barmode="stack"
)

# ✅ 6. Build Dash App
app = Dash(__name__)

app.layout = html.Div([
    html.H1("Employee Training Effectiveness Dashboard", style={"textAlign": "center"}),

    dcc.Graph(figure=fig_bar),
    dcc.Graph(figure=fig_pie),
    dcc.Graph(figure=fig_line),
    dcc.Graph(figure=fig_heatmap),
    dcc.Graph(figure=fig_stack),
])

# ✅ 7. Run App
if __name__ == "__main__":
    app.run(debug=True)
