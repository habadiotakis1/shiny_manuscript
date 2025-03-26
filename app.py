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
from docx.shared import Pt, Cm, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pickle

# set default and alternative statistical tests
default_tests = {
    "Omit": "Omit",
    "Categorical (Y/N)": "fisher",
    "Categorical (Dichotomous)": "fisher",
    "Categorical (Multinomial)": "fisher-freeman-halton",
    "Ratio Continuous": "ttest",
    "Ordinal Discrete": "wilcoxon",
}

alternative_tests = {
    "Categorical (Y/N)": "chi2",
    "Categorical (Dichotomous)": "chi2",
    "Categorical (Multinomial)": "chi2",
    "Ratio Continuous": "mannwhitney",
    "Ordinal Discrete": "ttest",
}

variable_types = list(default_tests.keys())

# get p-values from statistical test
################################################################################
### ONLY SUPPORTS 2 GROUPS AT THE MOMENT, NEED TO UPDATE TO MULTIPLE GROUPS ####
################################################################################
def run_statistical_test(df, group_var, var_type, var_name, decimal_places):
    groups = df[group_var].unique()
    if len(groups) != 2:
        return None  # Only supports two-group comparisons
    
    group1 = df[df[group_var] == groups[0]][var_name].dropna()
    group2 = df[df[group_var] == groups[1]][var_name].dropna()
    
    test_type = default_tests[var_type]

    if test_type == "fisher":
        contingency_table = pd.crosstab(df[var_name], df[group_var])
        _, p_value = stats.fisher_exact(contingency_table)
    elif test_type == 'fisher-freeman-halton': # UPDATE
        contingency_table = pd.crosstab(df[var_name], df[group_var])
        _, p_value, _, _ = stats.chi2_contingency(contingency_table, lambda_="log-likelihood")        
        # rpy2.robjects.numpy2ri.activate()
        # stats = importr('stats')
        # m = np.array([[4,4],[4,5],[10,6]])
        # res = stats.fisher_test(m)
        # print 'p-value: {}'.format(res[0][0])
    elif test_type == "chi2":
        contingency_table = pd.crosstab(df[var_name], df[group_var])
        _, p_value, _, _ = stats.chi2_contingency(contingency_table)
    elif test_type == "ttest":
        _, p_value = stats.ttest_ind(group1, group2, equal_var=False)
    elif test_type == "mannwhitney":
        _, p_value = stats.mannwhitneyu(group1, group2)
    elif test_type == "wilcoxon":
        _, p_value = stats.ranksums(group1, group2)
    else:
        print("Invalid test type:", test_type, " for variable:", var_name)
        p_value = None
    
    try:
        p_value = round(p_value, decimal_places)
    except:
        pass

    return p_value

# Function to perform aggregation analysis based on the variable type
def perform_aggregate_analysis(df, group_var, var_type, var_name, decimal_places, output_format, col_var_config):
    groups = df[group_var].unique()
    if len(groups) != 2:
        return None  # Only supports two-group comparisons
    # print(group_var, test_type, var_name, decimal_places, output_format, col_var_config)
    
    # Check if the variable has a "Yes" option
    yes_values = ['Yes', 'Y', 'y', 'yes', 1]
    yn_var = None

    var_options = df[var_name].unique()        
    for val in yes_values:
        if val in var_options:
            yn_var=val

    group1 = df[df[group_var] == groups[0]][var_name].dropna()
    group2 = df[df[group_var] == groups[1]][var_name].dropna()

    group1_total = len(group1)
    group2_total = len(group2)

    # Default result structure
    result = {}
    
    if var_type == "Omit":
        return None
    
    elif var_type == "Categorical (Y/N)":
        # count occurrences of yn_var in group 1 and group 2
        group1_sum = (group1 == yn_var).sum()
        group2_sum = (group2 == yn_var).sum()

        # Store the aggregate values in var_config
        if output_format == "n (%)":
            col_var_config['group1'] = str(group2_sum)  + " (" + str(round(group2_sum / group2_total * 100, decimal_places)) + "%)"
            col_var_config['group2'] = str(group1_sum) + " (" + str(round(group1_sum / group1_total * 100, decimal_places)) + "%)"
        else:
            col_var_config['group1'] = str(round(group1_sum / group1_total * 100, decimal_places)) + "% (" + str(group1_sum) + ")"
            col_var_config['group2'] = str(round(group2_sum / group2_total * 100, decimal_places)) + "% (" + str(group2_sum) + ")"
        
    elif var_type == "Categorical (Dichotomous)" or var_type == "Categorical (Multinomial)":
        for i in range(len(var_options)):
            group1_sum = (group1 == var_options[i]).sum()
            group2_sum = (group2 == var_options[i]).sum()

            if output_format == "n (%)":
                col_var_config[f'group1_subgroup{i}'] = str(group1_sum) + " (" + str(round(group1_sum / group1_total * 100, decimal_places)) + "%)"
                col_var_config[f'group2_subgroup{i}'] = str(group2_sum)  + " (" + str(round(group2_sum / group2_total * 100, decimal_places)) + "%)"
            else:
                col_var_config[f'group1_subgroup{i}'] = str(round(group1_sum / group1_total * 100, decimal_places)) + "% (" + str(group1_sum) + ")"
                col_var_config[f'group2_subgroup{i}'] = str(round(group2_sum / group2_total * 100, decimal_places)) + "% (" + str(group2_sum) + ")"

    elif var_type == "Ratio Continuous":
        # Aggregate: Mean and Standard Deviation
        group1_mean = round(group1.mean(), decimal_places)
        group2_mean = round(group2.mean(), decimal_places)

        group1_std = round(group1.std(), decimal_places)
        group2_std = round(group2.std(), decimal_places)

        col_var_config['group1'] = str(group1_mean) + " \u00B1 " + str(group1_std)
        col_var_config['group2'] = str(group2_mean) + " \u00B1 " + str(group2_std)

    elif var_type == "Ordinal Discrete":
        # Aggregate: Median and Interquartile Range (IQR)
        group1_median = group1.median()
        group2_median = group2.median()
        group1_iqr = [group1.quantile(0.25), group1.quantile(0.75)]
        group2_iqr = [group2.quantile(0.25), group2.quantile(0.75)]

        col_var_config['group1'] = str(group1_median) + " [" + str(group1_iqr[0]) + "-" + str(group1_iqr[1]) + "]"
        col_var_config['group2'] = str(group2_median) + " [" + str(group2_iqr[0]) + "-" + str(group2_iqr[1]) + "]"
    
    return col_var_config

# Function to create Word table from var_config
def create_word_table(df,var_config, group_var, subheadings):
    # Create a new Word Document
    doc = Document()

    # Create the table with columns for Variable, Group 1, Group 2, P-Value
    table = doc.add_table(rows=1, cols=4)
    table.columns[0].width=Inches(3.5)
    table.columns[1].width=Inches(1.5)
    table.columns[2].width=Inches(1.5)
    table.columns[3].width=Inches(.5)

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Variable'
    hdr_cells[1].text = 'Group 1'
    hdr_cells[2].text = 'Group 2'
    hdr_cells[3].text = 'P-Value'
    
    for row in hdr_cells:
        row.paragraphs[0].runs[0].font.bold = True  # Bold formatting for the subheading row
        
    group1_size = len(df[df[group_var] == df[group_var].unique()[0]])
    group2_size = len(df[df[group_var] == df[group_var].unique()[1]])
    grp_cells = table.add_row().cells
    grp_cells[0].text = ''
    grp_cells[1].text = '(n= ' + str(group1_size) + ")"
    grp_cells[2].text = '(n= ' + str(group2_size) + ")"
    grp_cells[3].text = ''

    # Loop through subheadings
    for subheading_name in subheadings:
        # Add a row for the subheading (this is the row header)
        row_cells = table.add_row().cells
        row_cells[0].text = f"{subheading_name}"  # Subheading name in the first column
        row_cells[1].text = ''  # Leave empty for Group 1
        row_cells[2].text = ''  # Leave empty for Group 2
        row_cells[3].text = ''  # Leave empty for P-Value

        row_cells[0].paragraphs[0].runs[0].font.bold = True  # Bold formatting for the subheading row
        
        # Get and sort all variables for the current subheading
        subheading_vars = [col for col, config in var_config.items() if config['subheading'] == subheading_name]
        sorted_subheading_vars = sorted(subheading_vars, key=lambda x: var_config[x]["position"])
        
        # Add a row for each variable under the current subheading
        for var in sorted_subheading_vars:
            var = var_config[var]["name"]
            var_type = var_config[var]["type"]

            if var_type == "Omit":
                continue
            elif var_type == "Categorical (Y/N)":
                row_cells = table.add_row().cells
                row_cells[0].text = f"{var}"  
                row_cells[1].text = str(var_config[var]["group1"])
                row_cells[2].text = str(var_config[var]["group2"])
                row_cells[3].text = str(var_config[var]["p_value"])

            elif var_type == "Categorical (Dichotomous)":
                row_cells = table.add_row().cells
                row_cells[0].text = f"{var}"  
                row_cells[1].text = ""
                row_cells[2].text = ""
                row_cells[3].text = ""

                row_cells[0].paragraphs[0].runs[0].font.underline = True

                var_options = df[var].unique()        
                for i in range(len(var_options)):
                    row_cells = table.add_row().cells
                    row_cells[0].text = f"   {var_options[i]}"  
                    row_cells[1].text = str(var_config[var][f"group1_subgroup{i}"])
                    row_cells[2].text = str(var_config[var][f"group2_subgroup{i}"])
                    if i == 0:
                        row_cells[3].text = str(var_config[var]["p_value"])
                    else:
                        row_cells[3].text = "-"

                    row_cells[0].paragraphs[0].runs[0].font.italic = True

            elif var_type == "Categorical (Multinomial)":
                row_cells = table.add_row().cells
                row_cells[0].text = f"{var}"  
                row_cells[1].text = ""
                row_cells[2].text = ""
                row_cells[3].text = ""

                row_cells[0].paragraphs[0].runs[0].font.underline = True

                var_options = df[var].unique()        
                for i in range(len(var_options)):
                    row_cells = table.add_row().cells
                    row_cells[0].text = f"   {var_options[i]}"  
                    row_cells[1].text = str(var_config[var][f"group1_subgroup{i}"])
                    row_cells[2].text = str(var_config[var][f"group2_subgroup{i}"])
                    if i == 0:
                        row_cells[3].text = str(var_config[var]["p_value"])
                    else:
                        row_cells[3].text = "-"

                    row_cells[0].paragraphs[0].runs[0].font.italic = True

            elif var_type == "Ratio Continuous" or var_type == "Ordinal Discrete":
                row_cells = table.add_row().cells
                row_cells[0].text = f"{var}"  
                row_cells[1].text = str(var_config[var]["group1"])
                row_cells[2].text = str(var_config[var]["group2"])
                row_cells[3].text = str(var_config[var]["p_value"])
                
                # Apply formatting to the variable name cell (indentation and smaller font)
                # para = row_cells[0].paragraphs[0]
                # run = para.add_run(row_cells[0].text)
                # run.font.size = Pt(8)  # Smaller font size
                # para.paragraph_format.left_indent = Pt(12)  # Indentation for the variable name            

    # Save the document to a file
    doc_filename = "statistical_analysis_results.docx"
    doc.save(doc_filename)
    return doc_filename


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
    
    ui.layout_columns(
        ui.h5("Step 1: Upload File"),
        ui.layout_columns(
            ui.card(ui.input_file("data_file", "Only .csv or .xlsx files will be accepted", accept=[".csv", ".xlsx"])),
            ui.card("Example Output File: ", ui.download_button("download_example", "Download Example")),
            col_widths=(8, 4),
            ),
        col_widths= 12,
        ),
    
    
    ui.h5("Step 2: Select Columns"),
    ui.output_ui('select_columns'),
    
    ui.h5("Step 3: Table Options"),

    ui.layout_columns(    
        # Table Name
        ui.card(ui.input_text("table_name", "Input Table Name", placeholder="Enter table name",width="100%")),

        # Grouping Variable
        ui.card(ui.output_ui("grouping_variable")),
        # ui.input_select("group_var", "Select Grouping Variable", choices=[], selected=None),
        
        # Formatting Options
        ui.card(ui.input_numeric("decimals_table", "Table - # Decimals", 2, min=0, max=5)),
        ui.card(ui.input_numeric("decimals_pvalue", "P-Val - # Decimals", 2, min=0, max=5)),
        ui.card(ui.input_radio_buttons("output_format", "Output Format", ["n (%)", "% (n)"])),
        col_widths= (4,2,2,2,2)
        ),

    ui.h5("Step 4: Customize Table"),
    # Subheadings
    ui.input_text("subheading_1", "Subheading 1", placeholder="Enter subheading 1 name"),
    ui.output_ui("var_settings_1"),

    ui.input_text("subheading_2", "Subheading 2", placeholder="Enter subheading 2 name"),
    ui.output_ui("var_settings_2"),
    
    ui.input_text("subheading_3", "Subheading 3", placeholder="Enter subheading 3 name"),
    ui.output_ui("var_settings_3"),
    
    ui.input_text("subheading_4", "Subheading 4", placeholder="Enter subheading 4 name"),
    ui.output_ui("var_settings_4"),
    
    
    # Variable Selection UI (dynamically generated)
    ui.output_ui("var_settings"),
    
    
    # Calculate
    ui.input_action_button("calculate", "Calculate"),
    
    # Download Button
    ui.download_button("download_table", "Download Table"),

    # Save Configuration
    ui.input_action_button("save_config", "Save Configuration"),
    ui.input_action_button("load_config", "Load Configuration"),
    
)

################################################################################
######################### Shiny App Server #####################################
################################################################################
def server(input, output, session):
    data = reactive.Value({})  # Store uploaded data
    selected_columns = reactive.Value([])  # Store selected columns
    var_config = reactive.Value({})  # Store variable settings dynamically
    subheadings = reactive.Value({0:"",1:None,2:None,3:None})  # Store subheadings
    group_var = reactive.Value(None)  # Store grouping variable
    decimal_places = reactive.Value(None)
    output_format = reactive.Value(None)


    @reactive.effect
    def save_configurations():
        subheadings.set({
            0: input.subheading_1(),
            1: input.subheading_2(),
            2: input.subheading_3(),
            3: input.subheading_4()
        })
        decimal_places.set(input.decimals_table())
        output_format.set(input.output_format())

    @output
    @render.ui 
    def select_columns():
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
            column_dict = {}
            for col in columns:
                column_dict[col] = col
            
            default_type = "Omit"
            default_position = 100
            
            # Store variable settings in a dictionary
            if not var_config.get():
                var_config.set({col: {
                    "type": default_type, 
                    "name": col, 
                    "subheading": 0,
                    "position": default_position,
                } for col in columns})

            return ui.input_selectize(  
                "column_selectize",  
                "Select desired columns below:",  
                {  
                    "": column_dict,  
                },  
                multiple=True,  
                width="100%",
            )  

    @reactive.calc
    def column_selectize():
        selected_columns.set(input.column_selectize())
    
    # Set Grouping Variable for analysis
    @output
    @render.ui
    def grouping_variable():
        return ui.input_select("grouping_var", "Grouping Variable", choices=[])

    @reactive.calc
    def _():
        select_columns = input.column_selectize()
        if len(select_columns) > 0:
            ui.update_select("grouping_var", choices=select_columns, selected=select_columns[0])

    @reactive.effect
    def grouping_var():
        group_var.set(input.grouping_var())
        

    @output
    @render.ui # @reactive.event()# @reactive.event(input.data_file)
    @reactive.event(input.select_columns)
    @reactive.event(input.subheading_2)
    @reactive.event(input.subheading_3)
    @reactive.event(input.subheading_4)
    def var_settings_1():
        if var_config.get():
            columns = selected_columns.get()
            subheading_cols = []
            for col in columns:
                if var_config.get()[col]["subheading"] == 0:
                    subheading_cols.append(col)
                
            return ui.layout_column_wrap(
            *[
                ui.card(
                    ui.h5(col),  # Column name title
                    ui.input_select(
                        f"var_type_{col}",
                        "Variable Type",
                        variable_types,
                        selected=var_config.get()[col]["type"],
                    ),
                    ui.input_text(
                        f"name_{col}",
                        "Column Name",
                        value=var_config.get()[col]["name"],
                    ),
                    ui.input_select(
                        f"subheading_{col}",
                        "Assign Subheading", 
                        subheadings.get().values(), 
                        selected=var_config.get()[col]["subheading"]
                    ),
                    ui.input_select(
                        f"position_{col}",
                        "Assign Position under Subheading", 
                        [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15],
                        selected=100,
                    ),
                    # col_widths=(4, 3, 3, 3, 12),
                    draggable=True,
                )
                for col in subheading_cols
            ],
            width='100%', # Each card takes up half the row
            )
    @output
    @render.ui # @reactive.event()# @reactive.event(input.data_file)
    @reactive.event(input.subheading_1)
    @reactive.event(input.subheading_3)
    @reactive.event(input.subheading_4)
    def var_settings_2():
        if var_config.get():
            columns = selected_columns.get()
            subheading_cols = []
            for col in columns:
                if var_config.get()[col]["subheading"] == 1:
                    subheading_cols.append(col)
            return ui.layout_column_wrap(
            *[
                ui.card(
                    ui.h5(col),  # Column name title
                    ui.input_select(
                        f"var_type_{col}",
                        "Variable Type",
                        variable_types,
                        selected=var_config.get()[col]["type"],
                    ),
                    ui.input_text(
                        f"name_{col}",
                        "Column Name",
                        value=var_config.get()[col]["name"],
                    ),
                    ui.input_select(
                        f"subheading_{col}",
                        "Assign Subheading", 
                        subheadings.get().values(), 
                        selected=var_config.get()[col]["subheading"]
                    ),
                    ui.input_select(
                        f"position_{col}",
                        "Assign Position under Subheading", 
                        [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15],
                        selected=100,
                    ),
                    # col_widths=(4, 3, 3, 3, 12),
                    draggable=True,
                )
                for col in subheading_cols
            ],
            width='100%', # Each card takes up half the row
            )
            # file_info = input.data_file()[0]
            # ext = os.path.splitext(file_info["name"])[-1]
            
            # if ext == ".csv":
            #     df = pd.read_csv(file_info["datapath"])  # Reads header row by default
            # elif ext == ".xlsx":
            #     df = pd.read_excel(file_info["datapath"])

            # data.set(df)  # Store data in reactive value
            # columns = df.columns.tolist()  # Get column names
            # columns = [re.sub(r'\W+', '', col) for col in columns]

            # subheading_options = [""] + [s for s in subheadings.get().values() if s]
            # default_type = "Omit"
            # default_position = 100

            # # Store variable settings in a dictionary
            # if not var_config.get():
            #     var_config.set({col: {
            #         "type": default_type, 
            #         "name": col, 
            #         "subheading": "None",
            #         "position": default_position,
            #         "p_value": None,
            #     } for col in columns})

            # Output grouping variable selection UI dynamically
            # ui.update_select(
            #     "group_var", 
            #     choices=columns
            # )

        #     return ui.layout_column_wrap(
        #     *[
        #         ui.card(
        #             ui.h5(col),  # Column name title
        #             ui.input_select(
        #                 f"var_type_{col}",
        #                 "Variable Type",
        #                 variable_types,
        #                 # selected=var_config.get()[col]["type"],
        #             ),
        #             ui.input_text(
        #                 f"name_{col}",
        #                 "Column Name",
        #                 value=var_config.get()[col]["name"],
        #             ),
        #             ui.input_select(
        #                 f"subheading_{col}",
        #                 "Assign Subheading", 
        #                 subheading_options, 
        #                 # selected=var_config.get()[col]["subheading"]
        #             ),
        #             ui.input_select(
        #                 f"position_{col}",
        #                 "Assign Position under Subheading", 
        #                 [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15],
        #                 selected=100,
        #             ),
        #         )
        #         for col in columns
        #     ],
        #     width=1 / 2, # Each card takes up half the row
        # )
    
    # Update variable settings dynamically when inputs change
    @reactive.effect
    def update_var_config():
        df = data.get()
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            return
        
        updated_config = var_config.get()

        for col in df.columns:
            updated_config[col]["type"] = input[f"var_type_{col}"]() or "Omit"
            updated_config[col]["name"] = input[f"name_{col}"]() or col
            updated_config[col]["subheading"] = input[f"subheading_{col}"]() or "None"
            updated_config[col]["position"] = input[f"position_{col}"]() or 100

        var_config.set(updated_config)  # Update stored config

    

    # Perform statistical analysis when the "Calculate" button is clicked
    @reactive.event(input.calculate)
    def calculate_statistical_analysis():
        df = data.get()
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            return
        
        try:
            group_var = group_var.get()  # Get the selected grouping column
            decimals_pval = input.decimals_pvalue()
            decimals_tab = input.decimals_table()
            output_format = input.output_format()
            
            # Check if grouping column is selected
            if group_var and decimal_places and output_format:
                updated_config = var_config.get()
                
                # Perform statistical analysis using the grouping variable
                for col in df.columns:
                    var_type = var_config.get()[col]["type"]
                    
                    if var_type != "Omit":
                        p_value = run_statistical_test(df, group_var, var_type, col, decimals_pval)
                        
                        # Store the p-value in the var_config dictionary
                        updated_config[col]["p_value"] = p_value
                        print(f"Column: {col}, Grouping Variable: {group_var}, p-value: {p_value}")

                        # Perform aggregate analysis and update var_config with the results
                        aggregate_result = perform_aggregate_analysis(df, group_var, var_type, col, decimals_tab, output_format, updated_config[col])
                        if aggregate_result:
                            updated_config[col].update(aggregate_result)

                var_config.set(updated_config)
                    
        except:
            return

    # Download Button - Trigger to save table in .docx format
    # Updated download_table function
    @session.download()
    def download_table():
        # Retrieve the data and var_config
        df = data.get()
        updated_config = var_config.get()
        
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            return None  # Return None if no data is available
        
        # Generate the Word table document
        doc_filename = create_word_table(data.get(), updated_config, group_var.get(), subheadings.get())
        
        return doc_filename  # Return the Word document file for download

    # @reactive.event(input.download_table)
    # def download_table():
    #     df = data.get()
    #     if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
    #         return
        
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