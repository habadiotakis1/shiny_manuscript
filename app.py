# to deploy in terminal
# rsconnect add --account habadio --name habadio --token 8D654CD9CCCEBE179F6DB7BE6764D040 --secret No1MZ0K+IHtQPDcwd7fNbJliPUVvvlcL5sJuuR3v
# rsconnect deploy shiny /Users/helenabadiotakis/Downloads/ChanLab/shiny_manuscript --name habadio --title shiny-manuscript

# imports
from shiny import App, reactive, render, ui
import pandas as pd
import shinywidgets as sw
import os
from io import StringIO
import re
from scipy import stats
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pickle

# set default and alternative statistical tests
default_tests = {
    "Omit": "Omit",
    "Categorical (Y/N)": "Fisher's Exact Test",
    "Categorical (Dichotomous)": "Fisher's Exact Test",
    "Categorical (Multinomial)": "Fisher's Exact Test",
    "Ratio Continuous": "T-Test",
    "Ordinal Discrete": "Wilcoxon Rank Sum Test",
}

default_aggregation = {
    "Omit": "Omit",
    "Categorical (Y/N)": "sum",
    "Categorical (Dichotomous)": "sum",
    "Categorical (Multinomial)": "sum",
    "Ratio Continuous": "mean",
    "Ordinal Discrete": "average",
}

alternative_tests = {
    "Categorical (Y/N)": "Chi-Square Test",
    "Categorical (Dichotomous)": "Chi-Square Test",
    "Categorical (Multinomial)": "Chi-Square Test",
    "Ratio Continuous": "Mann-Whitney Test",
    "Ordinal Discrete": "T-Test",
}

variable_types = list(default_tests.keys())

# get p-values from statistical test
################################################################################
### ONLY SUPPORTS 2 GROUPS AT THE MOMENT, NEED TO UPDATE TO MULTIPLE GROUPS ####
################################################################################
def run_statistical_test(df, group_var, test_type, var_name):
    groups = df[group_var].unique()
    if len(groups) != 2:
        return None  # Only supports two-group comparisons
    
    group1 = df[df[group_var] == groups[0]][var_name].dropna()
    group2 = df[df[group_var] == groups[1]][var_name].dropna()
    
    if test_type == "fisher":
        contingency_table = pd.crosstab(df[var_name], df[group_var])
        _, p_value = stats.fisher_exact(contingency_table)
    elif test_type == "chi2":
        contingency_table = pd.crosstab(df[var_name], df[group_var])
        _, p_value, _, _ = stats.chi2_contingency(contingency_table)
    elif test_type == "t-test":
        _, p_value = stats.ttest_ind(group1, group2, equal_var=False)
    elif test_type == "mannwhitney":
        _, p_value = stats.mannwhitneyu(group1, group2)
    elif test_type == "wilcoxon":
        _, p_value = stats.ranksums(group1, group2)
    else:
        p_value = None
    
    return p_value

# Function to perform aggregation analysis based on the variable type
def perform_aggregate_analysis(df, var_type, column_name, decimal_places):
    if var_type == "Omit":
        return None
    
    # Default result structure
    result = {}

    if var_type == "Categorical (Y/N)" or var_type == "Categorical (Dichotomous)":
        # Aggregate: Sum the binary outcomes (e.g., 1 for "Yes", 0 for "No")
        result["aggregate"] = df[column_name].sum()  # Count the 'Yes' values (or 1 values)
        result["percentage"] = round(df[column_name].mean() * 100, decimal_places)  # Percentage of 'Yes' (or 1) values

    elif var_type == "Categorical (Multinomial)":
        # Aggregate: Count occurrences of each category
        counts = df[column_name].value_counts()
        result["aggregate"] = counts.to_dict()
        result["by_group"] = {category: (count, round(count / len(df) * 100,decimal_places)) for category, count in counts.items()}

    elif var_type == "Ratio Continuous":
        # Aggregate: Mean and Standard Deviation
        result["aggregate"] = round(df[column_name].mean(),decimal_places)
        result["standard_dev"] = round(df[column_name].std(),decimal_places)

    elif var_type == "Ordinal Discrete":
        # Aggregate: Median and Interquartile Range (IQR)
        result["aggregate"] = df[column_name].median()
        result["interquartile_range"] = [
            df[column_name].quantile(0.25),  # 1st quartile (Q1)
            df[column_name].quantile(0.75),  # 3rd quartile (Q3)
        ]
    
    return result

# create stylized microsoft word table
def create_scientific_table(title, headers, data: pd.DataFrame, group_var, var_config):
    doc = Document()
    
    # Add title
    title_paragraph = doc.add_paragraph()
    title_run = title_paragraph.add_run(title)
    title_run.bold = True
    title_run.font.size = Pt(12)
    title_paragraph.alignment = 0  # Left alignment, 1=center, 2=right
    
    doc.add_paragraph()  # Add space
    
    # Create table
    num_rows = len(data) + 1  # Headers + data rows
    num_cols = len(headers)
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = 'Plain Table 3' #'Table Grid'
    
    # Format headers
    for col_idx, header in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = header
        run = cell.paragraphs[0].runs[0]
        run.bold = True
    
    # Add data rows
    for row_idx, row_data in enumerate(data.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row_data):
            table.cell(row_idx, col_idx).text = str(value)
    
    clean_title = re.sub(r'\W+', '', title)
    doc.save(clean_title + ".docx")
    print(f"Table saved as {clean_title}.docx")

# Save configuration to a .pkl file
def save_config(config, filename="config.pkl"):
    with open(filename, "wb") as f:
        pickle.dump(config, f)
    print(f"Configuration saved as {filename}")

# Load configuration from a .pkl file
def load_config(filename="config.pkl"):
    try:
        with open(filename, "rb") as f:
            config = pickle.load(f)
        return config
    except FileNotFoundError:
        print("No saved configuration found.")
        return None
    


################################################################################
######################### Shiny App Layout #####################################
################################################################################
app_ui = ui.page_fluid(
    ui.panel_title("✨ Shiny Manuscript Table Generator ✨"),
    
    # File Upload
    ui.input_file("data_file", "Step 1: Upload CSV or Excel file", accept=[".csv", ".xlsx"]),
    
    # Table Name
    ui.input_text("table_name", "Step 2: Input Table Name", placeholder="Enter table name"),
    
    # Subheadings
    ui.input_text("subheading_1", "Subheading 1", placeholder="Enter subheading 1 name"),
    ui.input_text("subheading_2", "Subheading 2", placeholder="Enter subheading 2 name"),
    ui.input_text("subheading_3", "Subheading 3", placeholder="Enter subheading 3 name"),
    ui.input_text("subheading_4", "Subheading 4", placeholder="Enter subheading 4 name"),
    
    # Variable Selection UI (dynamically generated)
    ui.output_ui("var_settings"),
    
    # Grouping Variable (dynamically generated)
    ui.output_ui("group_variable"),

    # Formatting Options
    ui.input_numeric("decimals", "Number of Decimal Places", 2, min=0, max=5),
    ui.input_radio_buttons("output_format", "Output Format", ["n (%)", "% (n)"]),
    
    # Calculate
    ui.input_action_button("calculate", "Calculate "),
    
    # Download Button
    ui.download_button("download_table", "Download Formatted Table"),

    # Save Configuration
    ui.input_action_button("save_config", "Save Configuration"),
    ui.input_action_button("load_config", "Load Configuration"),
    
)

################################################################################
######################### Shiny App Server #####################################
################################################################################
def server(input, output, session):
    data = reactive.Value({})  # Store uploaded data
    var_config = reactive.Value({})  # Store variable settings dynamically
    subheadings = reactive.Value({})  # Store subheadings
    group_variable = reactive.Value(None)  # Store grouping variable
    decimal_places = reactive.Value(None)
    output_format = reactive.Value(None)

    @reactive.effect
    def save_configurations():
        subheadings.set({
            1: input.subheading_1(),
            2: input.subheading_2(),
            3: input.subheading_3(),
            4: input.subheading_4()
        })
        decimal_places.set(input.decimals())
        output_format.set(input.output_format())
        
    @output
    @render.ui # @reactive.event()# @reactive.event(input.data_file)
    def var_settings():
        if input.data_file():
            file_info = input.data_file()[0]
            ext = os.path.splitext(file_info["name"])[-1]
            
            if ext == ".csv":
                df = pd.read_csv(file_info["datapath"])  # Reads header row by default
            elif ext == ".xlsx":
                df = pd.read_excel(file_info["datapath"])

            data.set(df)  # Store data in reactive value
            columns = df.columns.tolist()  # Get column names
            columns = [re.sub(r'\W+', '', col) for col in columns]

            subheading_options = ["None"] + [s for s in subheadings.get().values() if s]
            default_type = "Omit"
            default_position = 100

            # Store variable settings in a dictionary
            if not var_config.get():
                var_config.set({col: {
                    "type": default_type, 
                    "rename": col, 
                    "subheading": "None",
                    "position": default_position,
                    "p_value": None,
                } for col in columns})

            return ui.layout_column_wrap(
            *[
                ui.card(
                    ui.h5(col),  # Column name title
                    ui.input_select(
                        f"var_type_{col}",
                        "Variable Type",
                        variable_types,
                        # selected=var_config.get()[col]["type"],
                    ),
                    ui.input_text(
                        f"rename_{col}",
                        "Rename Column",
                        value=var_config.get()[col]["rename"],
                    ),
                    ui.input_select(
                        f"subheading_{col}",
                        "Assign Subheading", 
                        subheading_options, 
                        # selected=var_config.get()[col]["subheading"]
                    ),
                    ui.input_select(
                        f"position_{col}",
                        "Assign Position under Subheading", 
                        [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15],
                        selected=100,
                    ),
                )
                for col in columns
            ],
            width=1 / 2, # Each card takes up half the row
        )
    
    # Update variable settings dynamically when inputs change
    @reactive.effect
    def update_var_config():
        df = data.get()
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            return
        
        updated_config = var_config.get()

        for col in df.columns:
            updated_config[col]["type"] = input[f"var_type_{col}"]() or "Omit"
            updated_config[col]["rename"] = input[f"rename_{col}"]() or col
            updated_config[col]["subheading"] = input[f"subheading_{col}"]() or "None"
            updated_config[col]["position"] = input[f"position_{col}"]() or 100

        var_config.set(updated_config)  # Update stored config
        group_variable.set(input.grouping_var())  # Update grouping variable

    # Set Grouping Variable for analysis
    @output
    @render.ui
    def group_variable():
        df = data.get()
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            return
        return ui.input_select("grouping_var", "Select Grouping Variable", df.columns)

    # Perform statistical analysis when the "Calculate" button is clicked
    @reactive.event(input.calculate)
    def calculate_statistical_analysis():
        df = data.get()
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            return
        
        try:
            group_var = input.group_var()  # Get the selected grouping column
            decimal_places = input.decimals()
            output_format = input.output_format()

            # Check if grouping column is selected
            if group_var and decimal_places and output_format:
                updated_config = var_config.get()
                
                # Perform statistical analysis using the grouping variable
                for col in df.columns:
                    test_type = var_config.get()[col]["type"]
                    
                    if test_type != "Omit":
                        p_value = run_statistical_test(df, group_var, test_type, col, decimal_places, output_format)
                        # Store the p-value in the var_config dictionary
                        updated_config[col]["p_value"] = p_value
                        print(f"Column: {col}, Grouping Variable: {group_var}, p-value: {p_value}")

                        # Perform aggregate analysis and update var_config with the results
                        aggregate_result = perform_aggregate_analysis(df, test_type, col, decimal_places.get())
                        if aggregate_result:
                            updated_config[col].update(aggregate_result)

                var_config.set(updated_config)
                    
        except:
            return

    # Download Button - Trigger to save table in .docx format
    @reactive.event(input.download_table)
    def download_table():
        df = data.get()
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            return
        
        create_scientific_table(input.table_name, subheadings.get(), data.get(), group_var.get(), var_config.get())
        clean_title = re.sub(r'\W+', '', input.table_name)
        df.to_csv(f"{clean_title}.csv", index=False)
        return f"{clean_title}.csv"
    
    # Save Configuration Button - Trigger to save settings
    @reactive.event(input.save_config)
    def save_configuration():
        config_to_save = {
            "var_config": var_config.get(),
            "subheadings": subheadings.get(),
            "group_var": group_var.get()
        }
        save_config(config_to_save)  # Save the config to a file
        return "Configuration saved!"

    # Load Configuration Button - Trigger to load saved settings
    @reactive.event(input.load_config)
    def load_configuration():
        ui.input_file("config_file", "Upload pkl file", accept=[".pkl"])
        config = load_config(input.config_file)  # Load the config from a file
        if config:
            var_config.set(config["var_config"])
            subheadings.set(config["subheadings"])
            group_var.set(config["group_var"])
            return "Configuration loaded!"
        return "No saved configuration found."

app = App(app_ui, server)