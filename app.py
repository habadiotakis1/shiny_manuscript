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
import json

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
        subheading_vars = [col for col, config in var_config.items() if config['name'] in subheadings[subheading_name]()]
        subheading_vars = [col for col in subheading_vars if col != group_var]
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
        ui.card(ui.input_numeric("decimals_pvalue", "P-Val - # Decimals", 3, min=0, max=5)),
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
    # subheadings = reactive.Value({0:"",1:None,2:None,3:None})  # Store subheadings
    group_var = reactive.Value(None)  # Store grouping variable
    previous_group_var = reactive.Value(None)
    decimal_places = reactive.Value(None)
    output_format = reactive.Value(None)

    # Reactive values to track column assignments per subheading
    subheadings = {
        "subheading_1": reactive.Value([]),
        "subheading_2": reactive.Value([]),
        "subheading_3": reactive.Value([]),
        "subheading_4": reactive.Value([])
    }

    @output
    @render.ui
    def select_columns():
        return ui.input_selectize("column_selectize", "Select desired variables below:",  
                {  "": {"":""} },  
                multiple=True,  
                width="100%",
            )  
    
    @reactive.effect
    def _():
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
            default_position = 15
            
            # Store variable settings in a dictionary
            if not var_config.get():
                var_config.set({col: {
                    "type": default_type, 
                    "name": col, 
                    "position": default_position,
                } for col in columns})

            ui.update_selectize(  
                "column_selectize",  
                choices={"":column_dict}
            )  

    @reactive.effect
    def column_selectize():
        available_columns = set(input.column_selectize())

        selected_columns.set(available_columns)

        all_subheading_values = set()
        for subheading in subheadings:
            all_subheading_values = all_subheading_values.union(set(subheadings[subheading]()))
            
        for col in available_columns:
            if col not in all_subheading_values:
                subheadings["subheading_1"].set(subheadings["subheading_1"]() + [col])
        
        @reactive.effect
        def sync_column_selection_with_subheadings():
            for subheading in subheadings:
                # current_cols = set(subheadings[subheading]())

                # # Add new columns to subheading
                # new_cols = available_columns - current_cols
                # if new_cols:
                #     updated = list(current_cols.union(new_cols))
                #     subheadings[subheading].set(updated)

                # # Remove deselected columns from subheading_1
                # removed_cols = current_cols - available_columns
                # if removed_cols:
                #     updated = [col for col in subheadings[subheading]() if col not in removed_cols]
                #     subheadings[subheading].set(updated)
                
                generate_subheading_ui(subheading)

    # Set Grouping Variable for analysis
    @output
    @render.ui
    def grouping_variable():
        return ui.input_select("grouping_var", "Grouping Variable", choices=[])

    @reactive.effect
    def _():
        choices = input.column_selectize()
        if choices:
            ui.update_select("grouping_var", choices=choices, selected=choices[0])
            group_var.set(choices[0])  # Set the initial grouping variable
            print("First group var: ", group_var.get())

    @reactive.effect
    def update_group_var():
        new_group = input.grouping_var()
        old_group = group_var()
        
        if not new_group or new_group == old_group:
            return

        # # If there's a previous group_var, ask the user where to put it
        # if old_group:
        #     ui.modal_show(
        #         ui.modal(
        #             ui.h5("Move Previous Grouping Variable"),
        #             ui.p(f"Where should '{old_group}' be moved?"),
        #             ui.input_select(
        #                 "subheading_choice",
        #                 "Select Subheading",
        #                 choices=list(subheadings.keys()),
        #             ),
        #             ui.input_action_button("confirm_subheading", "Confirm"),
        #             easy_close=False,
        #         )
        #     )

        #     # Define an observer for when the user confirms their choice
        #     @reactive.effect
        #     def move_old_group_var():
        #         if input.confirm_subheading() > 0:  # Button clicked
        #             chosen_subheading = input.subheading_choice()
        #             if chosen_subheading:
        #                 subheadings[chosen_subheading].set(
        #                     subheadings[chosen_subheading]() + [old_group]
        #                 )
        #             ui.modal_remove()  # Close modal after selection

        # # Update tracking variables
        # previous_group_var.set(old_group)
        # group_var.set(new_group)
        
        # 1. Remove new group var from subheadings
        for subheading in subheadings:
            updated_cols = [col for col in subheadings[subheading]() if col != new_group]
            subheadings[subheading].set(updated_cols)
            generate_subheading_ui(subheading)

        # 2. Add previous group var back into first available subheading
        # if old_group:
        #     for subheading in subheadings:
        #         cols = subheadings[subheading]()
        #         if old_group not in cols:
        #             subheadings[subheading].set(cols + [old_group])
        #             generate_subheading_ui(subheading)
        #             break

        # 3. Save the new group_var
        previous_group_var.set(old_group)
        group_var.set(new_group)
        print("Set new group var: ", group_var.get())

    # Update columns under subheadings
    def generate_subheading_ui(subheading_key):
        columns = subheadings[subheading_key]()
        if not columns:
            return ui.p("No variables assigned yet.")

        return ui.layout_columns(
        *[
            ui.card(
                ui.h5(col),
                ui.input_text(
                    f"name_{col}",
                    "Column Name",
                    value=var_config.get()[col]["name"],
                ),
                ui.input_select(
                    f"var_type_{col}",
                    "Variable Type",
                    variable_types,
                    selected=var_config.get()[col]["type"],
                ),
                ui.input_select(
                    f"position_{col}",
                    "Position",
                    list(range(1,31)),
                    selected=var_config.get()[col]["position"],
                ),
                # col_widths=(3, 3, 3, 3),
                class_="draggable-item",
                id=f"{subheading_key}_{col}"
            )
            for col in columns
        ],
        col_widths=(12),
        # width=1, 
        class_="droppable-area",
        )
        
    @output
    @render.ui
    def var_settings_1():
        return generate_subheading_ui("subheading_1")

    @output
    @render.ui
    def var_settings_2():
        return generate_subheading_ui("subheading_2")

    @output
    @render.ui
    def var_settings_3():
        return generate_subheading_ui("subheading_3")

    @output
    @render.ui
    def var_settings_4():
        return generate_subheading_ui("subheading_4")

    # JavaScript to enable drag-and-drop using SortableJS
    ui.tags.script(
        """
        document.addEventListener("DOMContentLoaded", function() {
            document.querySelectorAll('.draggable-list').forEach(list => {
                new Sortable(list, {
                    group: 'shared',
                    animation: 150,
                    onEnd: function(evt) {
                        let movedVar = evt.item.dataset.var;
                        let newGroup = evt.to.id.replace('list-', '');
                        
                        // Update the server-side reactive variable
                        Shiny.setInputValue("dragged_var", JSON.stringify({movedVar, newGroup}));
                    }
                });
            });
        });
        """
    )
    # Handle drag-and-drop updates in the server
    @reactive.effect
    def update_subheadings():
        drag_event = input.dragged_var()
        if drag_event:
            drag_data = json.loads(drag_event)
            moved_var = drag_data["movedVar"]
            new_group = drag_data["newGroup"]

            # Remove from old subheading
            for key in subheadings:
                if moved_var in subheadings[key]():
                    subheadings[key].set([v for v in subheadings[key]() if v != moved_var])

            # Add to new subheading
            subheadings[new_group].set(subheadings[new_group]() + [moved_var])
            
    
    # Update variable settings dynamically when inputs change
    @reactive.effect
    def update_var_config():
        df = data.get()
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            return
        
        if selected_columns.get() is None or len(selected_columns.get()) == 0:
            return
        
        updated_config = var_config.get().copy()
        
        print("SELECTED COLUMNS", type(selected_columns.get()),selected_columns.get())
        for col in df.columns:
            # if col in set(selected_columns.get()):
            print("❗️ Updating variable configurations...", updated_config[col])
            updated_config[col]["type"] = input[f"var_type_{col}"]() or "Omit"
            updated_config[col]["name"] = input[f"name_{col}"]() or col
            updated_config[col]["position"] = input[f"position_{col}"]() or 15
            print("to...", updated_config[col])
        var_config.set(updated_config)  # Update stored config

    # Perform statistical analysis when the "Calculate" button is clicked
    @reactive.effect
    @reactive.event(input.calculate)
    def calculate_statistical_analysis():
        print("🔄 Calculate button pressed. Updating variable configurations...")
        df = data.get()
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            print(df)
            return
        
        try:
            curr_group_var = group_var.get()  # Get the selected grouping column
            decimals_pval = input.decimals_pvalue()
            decimals_tab = input.decimals_table()
            output_format = input.output_format()
            
            updated_config = var_config.get()
            
            # Perform statistical analysis using the grouping variable
            if len(selected_columns.get()) > 0:
                for col in df.columns:
                    if col != curr_group_var and col in selected_columns.get():
                        print(f"\n📂 Processing Variable: {col}", updated_config[col])
                        
                        var_type = updated_config[col]["type"]
                        
                        if var_type != "Omit":
                            p_value = run_statistical_test(df, curr_group_var, var_type, col, decimals_pval)
                            
                            # Store the p-value in the var_config dictionary
                            updated_config[col]["p_value"] = p_value
                            print(f"Column: {col}, Grouping Variable: {curr_group_var}, p-value: {p_value}")

                            # Perform aggregate analysis and update var_config with the results
                            aggregate_result = perform_aggregate_analysis(df, curr_group_var, var_type, col, decimals_tab, output_format, updated_config[col])
                            if aggregate_result:
                                updated_config[col].update(aggregate_result)
                            
                            print("After: ", updated_config[col])

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
        doc_filename = create_word_table(data.get(), updated_config, group_var.get(), subheadings)
        
        return doc_filename  # Return the Word document file for download

app = App(app_ui, server)