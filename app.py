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

# set default and alternative statistical tests
default_tests = {
    "Omit": "Omit",
    "Categorical (Dichotomous)": "Fisher's Exact Test",
    "Categorical (Multinomial)": "Fisher's Exact Test",
    "Ratio Continuous": "T-Test",
    "Ordinal Discrete": "Wilcoxon Rank Sum Test",
}

alternative_tests = {
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

# create stylized microsoft word table
def create_scientific_table(title, headers, data: pd.DataFrame, filename):
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
    
    doc.save(filename)
    print(f"Table saved as {filename}")


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
    
    # Grouping Variable
    ui.output_ui("group_variable"),
    
    # Formatting Options
    ui.input_numeric("decimals", "Decimals for values", 2, min=0, max=5),
    ui.input_radio_buttons("output_format", "Output Format", ["n (%)", "% (n)"]),
    
    # Calculate
    ui.input_action_button("calculate", "Calculate "),
    
    # Save Configuration
    ui.input_action_button("save_config", "Save Configuration"),
    ui.input_action_button("load_config", "Load Configuration"),
    
    # Download Button
    ui.download_button("download_table", "Download Formatted Table")
)

def server(input, output, session):
    data = reactive.Value({})  # Store uploaded data
    var_config = reactive.Value({})  # Store variable settings dynamically
    subheadings = reactive.Value({})  # Store subheadings
    # data = {}  # Store uploaded data
    # config = {}  # Store bookmarked configurations

    @reactive.effect
    def save_subheadings():
        subheadings.set({
            1: input.subheading_1(),
            2: input.subheading_2(),
            3: input.subheading_3(),
            4: input.subheading_4()
        })

    @output
    @render.ui
    @reactive.event(input.data_file)
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
                } for col in columns})

            return ui.layout_column_wrap(
            2,  # Display two cards per row
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
                        f"rename_{col}",
                        "Rename Column",
                        value=var_config.get()[col]["rename"],
                    ),
                    ui.input_select(
                        f"subheading_{col}",
                        "Assign Subheading", 
                        subheading_options, 
                        selected=var_config.get()[col]["subheading"]
                    ),
                    ui.input_select(
                        f"position_{col}",
                        "Assign Position under Subheading", 
                        [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15],
                        selected=100,
                    ),
                )
                for col in columns
            ]
        )
    
    # Update variable settings dynamically when inputs change
    @reactive.effect
    def update_var_config():
        df = data.get()
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            return
        
        updated_config = var_config.get()

        for col in df.columns:
            updated_config[col]["type"] = input.get(f"var_type_{col}") or "Omit"
            updated_config[col]["rename"] = input.get(f"rename_{col}") or col
            updated_config[col]["subheading"] = input.get(f"subheading_{col}") or "None"
            updated_config[col]["position"] = input.get(f"position_{col}") or 100

        var_config.set(updated_config)  # Update stored config

    @output
    @render.ui
    # @reactive.event(input.data_file)
    def group_variable():
        df = data.get()
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            return
        
        return ui.input_select("group_var", "Select Grouping Variable", df.columns)

    @session.download()
    def download_table():
        df = data.get()
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:  
            return
        
        create_scientific_table(input.table_name, input.subheadings, data['df'], input.table_name+".docx")
        df.to_csv("formatted_table_separate.csv", index=False)
        return "formatted_table.csv"

app = App(app_ui, server)




# # define app
# app_ui = ui.navset_tab(
#     ui.nav("Step 1: Upload Data", 
#         ui.input_file("data_file", "Upload CSV or Excel file", accept=[".csv", ".xlsx"]),
#         ui.input_action_button("next_1", "Next")
#     )
#     # ,
#     # ui.nav("Step 2: Define Table Name", 
#     #     ui.input_text("table_name", "Table Name", placeholder="Enter table name"),
#     #     ui.input_action_button("next_2", "Next")
#     # ),
#     # ui.nav("Step 3: Select Variables", 
#     #     ui.output_ui("var_settings"),
#     #     ui.input_action_button("next_3", "Next")
#     # ),
#     # ui.nav("Step 4: Select Grouping Variable", 
#     #     ui.output_ui("group_variable"),
#     #     ui.input_action_button("next_4", "Next")
#     # ),
#     # ui.nav("Step 5: Define Subheadings", 
#     #     ui.output_ui("subheading_ui"),
#     #     ui.input_action_button("next_5", "Next")
#     # ),
#     # ui.nav("Step 6: Formatting Options", 
#     #     ui.input_numeric("decimals", "Decimals for values", 2, min=0, max=5),
#     #     ui.input_radio_buttons("output_format", "Output Format", ["n (%)", "% (n)"]),
#     #     ui.input_action_button("next_6", "Next")
#     # ),
#     # ui.nav("Step 7: Download & Save", 
#     #     ui.input_action_button("save_config", "Save Configuration"),
#     #     ui.input_action_button("load_config", "Load Configuration"),
#     #     ui.download_button("download_table", "Download Formatted Table")
#     # )
# )

# def server(input, output, session):
#     data = {}  # Store uploaded data
#     config = {}  # Store bookmarked configurations

#     @output
#     @render.ui
#     def var_settings():
#         if input.data_file():
#             file_info = input.data_file()[0]
#             ext = os.path.splitext(file_info["name"])[-1]
#             if ext == ".csv":
#                 df = pd.read_csv(file_info["datapath"])
#             elif ext == ".xlsx":
#                 df = pd.read_excel(file_info["datapath"])
            
#             data["df"] = df
#             columns = df.columns
            
#             return ui.tag_list(
#                 *[ui.panel_well(
#                     ui.input_select(f"var_type_{col}", f"{col} Type", variable_types, selected="Ratio Continuous"),
#                     ui.input_text(f"rename_{col}", f"Rename {col}", value=col)
#                 ) for col in columns]
#             )

#     @output
#     @render.ui
#     def group_variable():
#         if "df" in data:
#             return ui.input_select("group_var", "Select Grouping Variable", data["df"].columns)

#     @output
#     @render.ui
#     def subheading_ui():
#         if "df" in data:
#             return sw.dragula("subheadings", data["df"].columns.tolist())

#     @session.download()
#     def download_table():
#         if "df" not in data:
#             return "No data available."
#         df = data["df"]
#         df.to_csv("formatted_table.csv", index=False)
#         return "formatted_table.csv"

# app = App(app_ui, server)

# # def server(input, output, session):
# #     data = {}  # Store uploaded data

# #     # @output
# #     # @render.ui
# #     # def var_settings():
# #     #     if input.data_file():
# #     #         file_info = input.data_file()[0]
# #     #         ext = os.path.splitext(file_info["name"])[-1]
# #     #         if ext == ".csv":
# #     #             df = pd.read_csv(file_info["datapath"])
# #     #         elif ext == ".xlsx":
# #     #             df = pd.read_excel(file_info["datapath"])
            
# #     #         data["df"] = df
# #     #         columns = df.columns
            
# #     #         return ui.tag_list(
# #     #             *[ui.panel_well(
# #     #                 ui.input_select(f"var_type_{col}", f"{col} Type", variable_types, selected="Ratio Continuous"),
# #     #                 ui.input_text(f"rename_{col}", f"Rename {col}", value=col)
# #     #             ) for col in columns]
# #     #         )

# #     @output
# #     @render.ui
# #     def group_variable():
# #         if "df" in data:
# #             return ui.input_select("group_var", "Select Grouping Variable", data["df"].columns)

# #     @output
# #     @render.ui
# #     def subheading_ui():
# #         if "df" in data:
# #             return sw.dragula("subheadings", data["df"].columns.tolist())

# #     @session.download()
# #     def download_table():
# #         if "df" not in data:
# #             return "No data available."
# #         df = data["df"]
# #         results = {}
        
# #         for col in df.columns:
# #             test_type = default_tests.get(input.get(f"var_type_{col}"), None)
# #             if test_type and input.group_var():
# #                 p_value = run_statistical_test(df, input.group_var(), test_type, col)
# #                 results[col] = p_value
        
# #         df_pvals = pd.DataFrame.from_dict(results, orient='index', columns=['P-Value'])
# #         df_pvals.to_csv("formatted_table.csv", index=True)
# #         return "formatted_table.csv"

# # app = App(app_ui, server)
