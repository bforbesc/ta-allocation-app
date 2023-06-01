# Import relevant libraries
import streamlit as st
import pandas as pd
import numpy as np

# define general random seed and plotly template
np.random.seed(2023)

# Function to upload excel files
def upload_excel_file(label):
    uploaded_file = st.file_uploader(label, type=['xlsx'])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            return df
        except Exception as e:
            # st.error(f"Please upload a valid file!")
            st.error(f"An error occurred while reading the file: {e}")
    return None

def upload_preferences_excel(label):
    uploaded_file = st.file_uploader(label, type=['xlsx'])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, header=1) # only difference is the "header"
            return df
        except Exception as e:
            # st.error(f"Please upload a valid file!")
            st.error(f"An error occurred while reading the file: {e}")
    return None


#########################################################################################################################################

# INTRO: write introduction
st.title("🧩 Teaching Assistants Allocation")

st.markdown("""
## Table of contents
1. [Application overview](#overview)
1. [Required inputs](#inputs)
1. [Analysis & outputs](#outputs)

## Aplication overview
The purpose of this application is to assign Teaching Assistants (TAs) to the faculty courses, taking into account various factors such as their contract percentage, individual preferences, and the specific needs of each course. 
By analyzing this information, the app streamlines the allocation process, ensuring a fair and optimized distribution of TAs across different courses. 
""", unsafe_allow_html=True)

st.markdown("""
## Required inputs
The app requires the following pieces of information, which are accepted as Excel files:
1. [Faculty courses](#courses): list of courses for (bachelor's and master's) with respective number of classes and expected number of students
    - A file ```bs_courses_weights_EMPTY.xlsx``` is generated from the list of courses, which can be obtained by clicking on the download button
    - The file should be filled in with the courses' respective weigths in the column ```weight```
1. [BS course weights](#weights): Bachelor's courses weights (conversion from needs to TAs contract percentage)
    - After filling in ```bs_courses_weights_EMPTY.xlsx```, the file should be then (re)uploaded to the platform
1. [TAs capacity](#capacity): TAs current contract percentage (from previous semester)
1. [TAs preferences](#preferences): TAs course and contract preferences

Finally, the current semester should be selected.

""", unsafe_allow_html=True)

# Create a dropdown list to select the semester
term = st.selectbox("Select Semester", ["S1", "S2"])

# Use the selected semester in your script
if term == "S1":
    st.write("Selected Semester: S1")
else:
    st.write("Selected Semester: S2")


# PART 1: LIST OF COURSES (DSD)
#########################################################################################################################################
st.markdown("""### Faculty courses""")

dsd_df = upload_excel_file("Please upload the course list")
if dsd_df is not None:
    dsd_df["course"] = dsd_df["COURSE CODE"].astype(str) + " || " + dsd_df["COURSE NAME"].astype(str) + " || " + dsd_df["TERM"].astype(str) + " || " + dsd_df["LANGUAGE"].astype(str)

    if term == "S1":
        selected_terms = ["S1", "T1", "T2"]
        dsd_df = dsd_df[dsd_df['TERM'].isin(selected_terms)]
    else:
        selected_terms = ["S2", "T3", "T4"]
        dsd_df = dsd_df[dsd_df['TERM'].isin(selected_terms)]
    # Only perform the data cleaning and transformation once
    agg_functions = {
        'CLASS': 'count',
        'SLOTS': np.sum
    }
    
    # OUTPUT #1: LIST OF COURSES IMPORTED
    output_1 = dsd_df.groupby(['COURSE NAME', 'TERM', 'COURSE CODE', 'LANGUAGE']).agg(agg_functions).reset_index()
    output_1 = output_1.rename(columns={'CLASS': 'Nº CLASSES', 'SLOTS': 'Nº STUDENTS'})

    # Table actually used for computations (different from OUPUT #1)
    course_demand = dsd_df.groupby(['course']).agg(agg_functions).reset_index()
    course_demand = course_demand.rename(columns={'CLASS': 'number_classes', 'SLOTS': 'number_students'})

    course_demand_extended = course_demand.copy()
    course_demand_extended[["course_code", "course_name", "period", "language"]] = course_demand["course"].str.split(" \|\| ", expand=True)
    course_demand_extended = course_demand_extended[["course", "course_code", "course_name", "period", "language"]]

    # INPUT #3 Get file course list to manually input the weights
    course_demand_extended['masters_course'] = course_demand_extended['course'].apply(lambda x: 0 if x.split(' ')[0].startswith('1') else 1)
    course_demand_extended_bs = course_demand_extended[course_demand_extended.masters_course==0]
    course_demand_extended_bs = course_demand_extended_bs.drop(columns=["masters_course"])
    course_demand_extended_bs["weight"] = ""

    # Download button    
    course_demand_extended_bs.to_excel("bs_courses_weights_EMPTY.xlsx", index=False)
    # Provide download button for the Excel file
    with open("bs_courses_weights_EMPTY.xlsx", "rb") as file:
        file_data = file.read()
        st.download_button(
            label="Please download bachelor's courses to fill in the weights",
            data=file_data,
            file_name="bs_courses_weights_EMPTY.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown("""### BS course weights""") 
bs_weights_df = upload_excel_file("Please upload bachelor's courses weights")
if bs_weights_df is not None:
    bs_weights_df = bs_weights_df[["course", "weight"]]
    bs_weights_df["weight"] = bs_weights_df["weight"] * 0.125


# PART 2: TAs CURRENT CONTRACT
#########################################################################################################################################
# INPUT #2
st.markdown("""### TAs capacity""") 
contract = upload_excel_file("Please upload the TAs contract file")
if contract is not None:
    contract["TA"] = contract["TA"].str.lower()
    # Drop contracts with zero percentage
    zero_contracts = contract[contract['contract'] == 0]["TA"].unique()
    contract = contract[contract['contract'] != 0]
    contract_emails = contract.TA.unique()


# PART 3: TAs PREFERENCES (QUALTRICS SURVEY)
#########################################################################################################################################

# PART 3.1: Cleaning the data
###############################################################
# INPUT #3
st.markdown("""### TAs preferences""") 
preferences_df = upload_preferences_excel("Please upload the TAs preferences")
if preferences_df is not None:
    # Sort the DataFrame by "End Date" column in descending order
    preferences_df = preferences_df.sort_values(by='End Date', ascending=False)

    # Rename the column to "TA"
    preferences_df.rename(columns={'Please write your E-mail @novasbe.pt': 'TA'}, inplace=True)

    # Convert the values in the "TA" column to lowercase
    preferences_df['TA'] = preferences_df['TA'].str.lower()

    # Remove TAs with zero_contracts
    preferences_df = preferences_df[~preferences_df["TA"].isin(zero_contracts)]

    # Create a mask to identify duplicates in the "TA" column
    duplicates_mask = preferences_df.duplicated(subset='TA', keep=False)
    preferences_duplicates = preferences_df[duplicates_mask]
    preferences_duplicates = preferences_duplicates.sort_values(by='End Date', ascending=False)

    preferences_duplicates_last = preferences_duplicates.drop_duplicates(subset='TA', keep='first').copy()

    # Create a mask to check if columns 31:81 or 83:347 have values
    value_mask = preferences_duplicates.iloc[:, 31:81].notnull().any(axis=1) | preferences_duplicates.iloc[:, 83:347].notnull().any(axis=1)
    preferences_duplicates_values = preferences_duplicates[value_mask]
    preferences_duplicates_values = preferences_duplicates_values.drop_duplicates(subset='TA', keep='first').copy()

    # Drop duplicates based on the "TA" column
    preferences_df = preferences_df[~duplicates_mask]

    # Drop duplicates based on the "Full Name" column while keeping the row with the most recent "End Date" (ex. Franziska wrong)
    preferences_df = preferences_df.drop_duplicates(subset='Full Name', keep='first')

    # Create a new DataFrame with columns from preferences_duplicates_last
    preferences_df_final = preferences_duplicates_last.copy()

    # Get the relevant columns from preferences_duplicates_values
    preference_columns = preferences_duplicates_values.columns[31:81].tolist() + preferences_duplicates_values.columns[83:347].tolist()

    # Update the values in preferences_df_final using values from preferences_duplicates_values for preference_columns
    preferences_df_final.set_index('TA', inplace=True, drop=False)
    preferences_duplicates_values.set_index('TA', inplace=True, drop=False)
    preferences_df_final.loc[preferences_duplicates_values.index, preference_columns] = preferences_duplicates_values[preference_columns].values

    # Concatenate the remaining columns from preferences_df to preferences_df_final
    preferences_df_final = pd.concat([preferences_df_final, preferences_df])

    # Sort the final DataFrame by "End Date" column in descending order
    preferences_df_final.sort_values(by='End Date', ascending=False, inplace=True)

    # Reset the index of the final DataFrame
    preferences_df_final.reset_index(drop=True, inplace=True)

    # Drop duplicates based on the "Full Name" column while keeping the row with the most recent "End Date" (ex. Franziska wrong )
    preferences_df_final.drop_duplicates(subset='Full Name', keep='first', inplace=True)

    # Create a mapping of original column names to new column names (course ID as integer)
    mapping = {}

    for column_name in preference_columns:
        # Extract the course ID number from the column name
        course_id = column_name.split(' || ')[0].split(' - ')[3] + " || "  + column_name.split(' || ')[0].split(' - ')[4]+ " || " + column_name.split(' || ')[1] + ' || ' +  column_name.split(' || ')[2].split(' - ')[0]    
        # Map the original column name to the course ID
        mapping[column_name] = course_id


    # Rename the columns using the mapping
    preferences_df_final.rename(columns=mapping, inplace=True)

    # Drop columns with list of courses (redundant) [30, 81, and 82]
    preferences_df_final.drop(columns=preferences_df_final.iloc[:,[30, 81, 82]], inplace=True)

    # OUTPUT #2: TAs LEAVING THIS SEMESTER
    output_2 = preferences_df_final[preferences_df_final.iloc[:, 19] == "No"].iloc[:, [17, 18, 20]]
    output_2 = output_2.rename(columns={output_2.columns[-1]: "Comments"}).sort_values("Full Name")
    
    ta_exits_list = output_2.TA.unique()

    # Filter the DataFrame for rows where "Do you intend to continue your collaboration with Nova SBE next semester as Teaching Assistant?" (column 20) is not equal to "No"
    preferences_df_final = preferences_df_final[preferences_df_final.iloc[:, 19] != "No"]

    # OUTPUT #3: TAs COMMENTS
    output_3 = preferences_df_final[~preferences_df_final.iloc[:,-1].isna()].iloc[:, [17, 18, -1]]

    # OUTPUT #4: TAs EMAILS FROM SURVEY WHICH ARE NOT IN THE TA CONTRACT DATABASE
    output_4 = preferences_df_final[~preferences_df_final["TA"].isin(contract_emails)][["TA", "Full Name"]]

    # Get the course columns
    course_columns = preferences_df_final.columns[30:-1]

    # Create a new DataFrame for the adapted format
    adapted_df = pd.DataFrame(columns=["TA", "course", "preference", "preference_type"])

    # Define the translation mapping for column 22 values
    translation_mapping = {
        "Masters' Courses": 2,
        "Bachelors' Courses": 0,
        "Indifferent": 1,
        pd.NaT: 1  # Assuming NaN values should also be considered "Indifferent"
    }

    # Iterate over the course columns
    for course in course_columns:
        # Check if the course has already been processed
        if course in adapted_df["course"].unique():
            continue

        # Get the duplicate columns for the current course
        duplicate_columns = [col for col in course_columns if col != course and col.endswith(course)]

        # Combine the duplicate columns into a single column
        combined_column = preferences_df_final[[course] + duplicate_columns].ffill(axis=1).iloc[:, -1]

        # Filter the DataFrame for non-null values in the combined column
        non_null_mask = combined_column.notnull()
        non_null_df = preferences_df_final[non_null_mask]

        # Get the teacher names and their corresponding preference rankings for the current course
        teacher_names = non_null_df["TA"]
        preference_rankings = combined_column[non_null_mask]

        # Get the corresponding preference types based on the translation mapping
        preference_types = non_null_df.iloc[:, 22].map(translation_mapping)

        # Create a DataFrame for the current course, preference rankings, and preference types
        course_df = pd.DataFrame({"TA": teacher_names, "course": [course] * len(teacher_names),
                                "preference": preference_rankings, "preference_type": preference_types})

        # Concatenate course_df with adapted_df
        adapted_df = pd.concat([adapted_df, course_df], ignore_index=True)
        
        # Create the 'masters_course' column based on the condition
        adapted_df['masters_course'] = adapted_df['course'].apply(lambda x: 0 if x.split(' ')[0].startswith('1') else 1)

        # Convert "preference" column to integers
        adapted_df['preference'] = adapted_df['preference'].astype(np.int8)

        # Remove preferences above 5
        adapted_df = adapted_df[adapted_df['preference']<=5]

        completed_preferences = adapted_df["TA"].unique()

    # OUTPUT #5: TAs COURSE PREFERENCES
    output_5 = adapted_df.iloc[:,:-1]

    # OUTPUT #6: TAs TO CONTACT (WHO DID NOT FILL-IN THE SURVEY AND ARE NOT LEAVING)
    output_6 = contract[(~contract.TA.isin(completed_preferences)) & (~contract.TA.isin(ta_exits_list))]


    # PART 3.2: Checking contract changes requested
    ###############################################################
    def round_to_closest(value):
        if pd.isnull(value):
            return np.nan
        else:
            capped_value = min(value, 0.5)  # Cap the value at 0.5
            capped_value = max(value, 0.1)  # Cap the value at 0.1
            return capped_value

    def clean_percentage(value): # Clean the "load_requested" column
        if pd.isnull(value):
            return value
        elif isinstance(value, str):
            # Check if the value contains only text characters
            if value.isalpha():
                return np.nan

            # Extract numeric values from string
            numeric_value = ''.join(filter(str.isdigit, value))

            if numeric_value == '':
                return np.nan

            if numeric_value == '100':
                return 100

            if len(numeric_value) >= 2:
                integer_part = numeric_value[:2]
                decimal_part = numeric_value[2:]
                return float(integer_part + '.' + decimal_part)

            return np.nan

        elif isinstance(value, (int, float)):
            return float(value) / 100

        return value

    def decrease_contract_level(value):
        return value - 0.125


    mapping = {
        "I want to increase the contract percentage/workload in the next semester (please specify the desired contract percentage level)": 1,
        "I want to keep the same contract percentage/workload as this semester": 0,
        "I want to reduce the contract percentage/workload in the next semester (please specify the desired contract percentage level)": -1,
        pd.NaT: 0
    }

    mapping_21 = {
        "Yes, I am a PhD student": 0,
        "Yes, I will be a Masters student but not doing any courses, only the Work Project": 0,
        "Yes, I will be a Masters student and I will be doing at least one more course": 1,
        "No": 0,
        pd.NaT: 0
    }

    mapping_23 = {
        "Yes, I have some other constraints that limit my teaching hours/workload (please specify the reason and the limit)": 1,
        "Yes, I have a FCT scholarship that limits my weekly teaching hours to 4h per week": 1,
        "No": 0,
        pd.NaT: 0
    }

    mask = preferences_df_final.iloc[:, 27].notna()
    new_contract = preferences_df_final[mask].iloc[:, [18, 21, 23, 27, 28, 29]]

    new_contract.columns = ['TA', 'master_student', 'PhD_restrictions', 'change_load', 'new_contract_decreased_load', 'new_contract_increased_load']
    new_contract['change_load'] = new_contract['change_load'].map(mapping)
    new_contract['master_student'] = new_contract['master_student'].map(mapping_21).fillna(0).astype(int)
    new_contract['PhD_restrictions'] = new_contract['PhD_restrictions'].map(mapping_23).fillna(0).astype(int)

    # Convert TA column to lowercase
    new_contract['TA'] = new_contract['TA'].str.lower()

    new_contract['new_contract_decreased_load'] = new_contract['new_contract_decreased_load'].apply(clean_percentage) / 100
    new_contract['new_contract_increased_load'] = new_contract['new_contract_increased_load'].apply(clean_percentage) / 100

    # Merge "new_contract_decreased_load" and "new_contract_increased_load" into "load_requested"
    new_contract['load_requested'] = new_contract[['new_contract_decreased_load', 'new_contract_increased_load']].mean(axis=1)
    new_contract['load_requested'] = new_contract['load_requested'].apply(round_to_closest)

    # Drop "new_contract_decreased_load" and "new_contract_increased_load" columns
    new_contract.drop(columns=['new_contract_decreased_load', 'new_contract_increased_load'], inplace=True)

    # OUTPUT #7: TAs WHO WANT TO CHANGE THEIR CONTRACT
    output_7 = new_contract[new_contract.change_load !=0].sort_values(by=["change_load", "TA"])


    all_contracts = contract.merge(new_contract, how="left", on="TA")

    # Filter rows where change_load is not equal to 0
    filtered_contracts = all_contracts[all_contracts['change_load'] != 0].copy()

    # Decrease contract to load_requested for rows where change_load is -1
    filtered_contracts.loc[filtered_contracts['change_load'] == -1, 'new_contract'] = filtered_contracts['load_requested']
    filtered_contracts.loc[(filtered_contracts['change_load'] == -1) & (filtered_contracts['load_requested'].isnull()), 'new_contract'] = filtered_contracts.apply(lambda row: decrease_contract_level(row['contract']), axis=1)

    # Fill NaN values with the original contract value
    filtered_contracts['new_contract'].fillna(filtered_contracts['contract'], inplace=True)

    # Create a new column "new_contract" in the original DataFrame with NaN values
    all_contracts['new_contract'] = np.nan

    # Update the "new_contract" column in the original DataFrame with the filtered values
    all_contracts.update(filtered_contracts[['new_contract']])
    all_contracts['new_contract'].fillna(all_contracts['contract'], inplace=True)

    # Drop emails which currently do not have a contract (ex. pedro.brinca)
    all_contracts = all_contracts[all_contracts.contract.notna()]



    # PART 4: FINAL DATA
    #########################################################################################################################################

    # PART 4.1: Merge all dataframes
    ###############################################################

    ta_preferences = adapted_df.merge(all_contracts, how="left", on="TA")
    market = ta_preferences.merge(course_demand, how="left", on="course", indicator=True)

    non_matching_values = market[market['_merge'] != 'both']
    market.drop(columns=["_merge"], inplace=True)

    non_matching_courses = non_matching_values[["course"]].drop_duplicates()
    non_matching_courses = non_matching_courses.copy()
    non_matching_courses[["course_code", "course_name", "period", "language"]] = non_matching_courses["course"].str.split(" \|\| ", expand=True)

    # Initialize an empty DataFrame to store the concatenated results
    concatenated_matches = pd.DataFrame()

    # Merge on 'course_code', 'period', and 'language'
    merged_courses = pd.merge(non_matching_courses, course_demand_extended, on=["course_code", "period", "language"], how="left", suffixes=("", "_new"))
    still_unmatched = merged_courses[merged_courses["course_new"].isna()][["course", "course_name", "course_code", "period", "language"]]
    concatenated_matches = pd.concat([concatenated_matches, merged_courses[~merged_courses["course_new"].isna()][["course", "course_new"]]])

    # Merge on 'course_name', 'period', and 'language'
    merged_courses = pd.merge(still_unmatched, course_demand_extended, on=["course_name", "period", "language"], how="left", suffixes=("", "_new"))
    still_unmatched = merged_courses[merged_courses["course_new"].isna()][["course", "course_name", "course_code", "period", "language"]]
    concatenated_matches = pd.concat([concatenated_matches, merged_courses[~merged_courses["course_new"].isna()][["course", "course_new"]]])

    # Merge on 'course_code' and 'period'
    merged_courses = pd.merge(still_unmatched, course_demand_extended, on=["course_code", "period"], how="left", suffixes=("", "_new"))
    still_unmatched = merged_courses[merged_courses["course_new"].isna()][["course", "course_name", "course_code", "period", "language"]]
    concatenated_matches = pd.concat([concatenated_matches, merged_courses[~merged_courses["course_new"].isna()][["course", "course_new"]]])

    # Merge concatenated_matches on the market DataFrame to add the "course_new" column
    market = pd.merge(market, concatenated_matches[["course", "course_new"]], on=["course"], how="left")
    market["course_new"].fillna(market["course"], inplace=True)
    market.rename(columns={"course": "course_old"}, inplace=True)
    market.drop(columns=["course_old"], inplace=True)
    market.rename(columns={"course_new": "course"}, inplace=True)

    # Merge market and course_demand on "course" column
    merged_market = pd.merge(market, course_demand[["course", "number_classes", "number_students"]], on="course", how="left", suffixes=("", "_demand"))

    # Fill NaN values in number_classes and number_students columns
    merged_market["number_classes"].fillna(merged_market["number_classes_demand"], inplace=True)
    merged_market["number_students"].fillna(merged_market["number_students_demand"], inplace=True)

    # Drop the unnecessary columns
    merged_market.drop(columns=["number_classes_demand", "number_students_demand"], inplace=True)

    # OUTPUT #8: COURSES FROM SURVEY (QUALTRICS) WITHOUT MATCH IN COURSE LIST (DSD)
    no_matches_final = merged_market[(merged_market.number_classes.isna()) | (merged_market.number_students.isna())][["course"]]
    output_8 = no_matches_final.drop_duplicates()

    # Drop these courses
    merged_market.dropna(subset=["number_classes", "number_students"], inplace=True)

    # PART 4.2: Compute capacities
    ###############################################################
    # Create the "semester" column based on the condition
    merged_market['semester'] = merged_market['course'].apply(lambda x: 1 if x.split(' || ')[2].startswith('S') else 0)
    merged_market['ms_capacity'] = merged_market['new_contract'] * 36

    # Define a function to apply the conditions
    def calculate_weight(row):
        if row['semester'] == 1:
            return (row['number_students'] * 2.33) / 16
        else:
            return (row['number_students'] * 1.25) / 16

    # Apply the function to create the 'ms_weight' column
    merged_market['ms_weight'] = merged_market.apply(calculate_weight, axis=1)

    # Set 'ms_capacity' to NaN when 'masters_course' is 0
    merged_market.loc[merged_market['masters_course'] == 0, 'ms_capacity'] = np.nan
    merged_market.loc[merged_market['masters_course'] == 0, 'ms_weight'] = np.nan

    # Final table
    final_market = pd.merge(merged_market, bs_weights_df, on=["course"], how="left", suffixes=("", "_new"), indicator=True)
    final_market.rename(columns={"weight": "bs_weight", "new_contract": "bs_capacity"}, inplace=True)
    final_market.drop(columns="_merge", inplace=True)
    final_market["bs_weight"] = final_market["bs_weight"] * final_market["number_classes"]
    final_market.loc[final_market['masters_course'] == 1, 'bs_capacity'] = np.nan
    final_market.loc[final_market['masters_course'] == 1, 'bs_weight'] = np.nan

    # OUTPUT #9: TAs AFFECTED BY COURSES WHICH ARE NOT MATCHED ON THE COURSE LIST (DSD)
    tas = final_market.TA.unique()
    output_9 = np.setdiff1d(completed_preferences, tas)


    # Part 5: OUTPUTS
    #########################################################################################################################################

    output_vars = [output_1, output_2, output_3, output_4, output_5, output_6, output_7, output_8]
    for output in output_vars:
        output.reset_index(inplace=True, drop=True)

    st.markdown('## Analysis & outputs', unsafe_allow_html=True)    

    st.markdown('### Data processing issues', unsafe_allow_html=True)    

    show_output_4 = st.checkbox("Unmatched TAs e-mails from survey (which are not in the database)")
    if show_output_4:
        st.write("")
        filtered_output_4 = output_4  
        filter_col = st.selectbox("Column", filtered_output_4.columns)
        unique_values = filtered_output_4[filter_col].unique().tolist()
        unique_values.insert(0, "")  
        filter_value = st.selectbox("Value", unique_values)
        if filter_value:
            if filter_value != "":
                filtered_output_4 = filtered_output_4[filtered_output_4[filter_col].str.replace(',', '') == filter_value]
        st.write(filtered_output_4)

    show_output_8 = st.checkbox("Courses from survey without match in course list")
    if show_output_8:
        filtered_output_8 = output_8  
        filter_col = st.selectbox("Column", filtered_output_8.columns)
        unique_values = filtered_output_8[filter_col].unique().tolist()
        unique_values.insert(0, "")  
        filter_value = st.selectbox("Value", unique_values)
        if filter_value:
            if filter_value != "":
                filtered_output_8 = filtered_output_8[filtered_output_8[filter_col].str.replace(',', '') == filter_value]
        st.write(filtered_output_8)

    show_output_9 = st.checkbox("TAs affected by unmatched courses (between course list and survey)")
    if show_output_9:
        output_9 = pd.DataFrame(output_9, columns=["TA"])
        st.write(output_9)


    st.markdown('### Contract changes', unsafe_allow_html=True)    

    show_output_2 = st.checkbox("TAs leaving this semester")
    if show_output_2:
        filtered_output_2 = output_2  # Assign output_2 to a filtered_output_2 variable
        filter_col = st.selectbox("Column", filtered_output_2.columns)
        unique_values = filtered_output_2[filter_col].unique().tolist()
        unique_values.insert(0, "")  
        filter_value = st.selectbox("Value", unique_values)
        if filter_value:
            if filter_value != "":
                filtered_output_2 = filtered_output_2[filtered_output_2[filter_col].str.replace(',', '') == filter_value]
        st.write(filtered_output_2)
    
    show_output_7 = st.checkbox("TAs who want to change contract workload")
    if show_output_7:
        output_7.rename(columns={'master_student': 'MS student', "PhD_restrictions":"PhD restriction", "change_load": "Contract change", "load_requested":"Load requested"}, inplace=True)
        mapping = {0: "No change", 1: "Increase", -1: "Decrease"}
        output_7["Contract change"] = output_7["Contract change"].map(mapping)
        filtered_output_7 = output_7  
        filter_col = st.selectbox("Column", filtered_output_7.columns)
        unique_values = filtered_output_7[filter_col].unique().tolist()
        unique_values.insert(0, "")  
        filter_value = st.selectbox("Value", unique_values)
        if filter_value:
            if filter_value != "":
                filtered_output_7 = filtered_output_7[filtered_output_7[filter_col].str.replace(',', '') == filter_value]
        st.write(filtered_output_7)

    st.markdown('### TAs to contact', unsafe_allow_html=True)    

    show_output_6 = st.checkbox("TAs who did not fill in the preferences and are not terminating")
    if show_output_6:
        output_6.rename(columns={'contract': 'Contract'}, inplace=True)
        filtered_output_6 = output_6  
        filter_col = st.selectbox("Column", filtered_output_6.columns)  
        unique_values = filtered_output_6[filter_col].unique().tolist()
        unique_values.insert(0, "")  
        filter_value = st.selectbox("Value", unique_values)
        if filter_value:
            if filter_value != "":
                filtered_output_6 = filtered_output_6[filtered_output_6[filter_col].str.replace(',', '') == filter_value]
        st.dataframe(filtered_output_6)
        # Provide download button for the Excel file
        filtered_output_6.to_excel("tas_to_call.xlsx", index=False)
        with open("tas_to_call.xlsx", "rb") as file:
            file_data = file.read()
            st.download_button(
                label="Download this table",
                data=file_data,
                file_name="ta_to_call.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    st.markdown('### Course preferences', unsafe_allow_html=True)

    show_output_1 = st.checkbox("Course list")
    if show_output_1:
        filtered_output_1 = output_1  
        filter_col = st.selectbox("Column", filtered_output_1.columns)
        unique_values = filtered_output_1[filter_col].unique().tolist()
        unique_values.insert(0, "")  
        filter_value = st.selectbox("Value", unique_values)
        if filter_value:
            if filter_value != "":
                filtered_output_1 = filtered_output_1[filtered_output_1[filter_col].str.replace(',', '') == filter_value]
        st.write(filtered_output_1)

    show_output_5 = st.checkbox("TAs ranked course preferences")
    if show_output_5:
        output_5.rename(columns={'course': 'Course', "preference":"Preference", "preference_type": "Program preference"}, inplace=True)
        mapping = {0: "BS", 1: "Indifferent", 2: "MS", np.NaN: "Indifferent"}
        output_5["Program preference"] = output_5["Program preference"].map(mapping)
        filtered_output_5 = output_5  
        filter_col = st.selectbox("Column", filtered_output_5.columns)
        unique_values = filtered_output_5[filter_col].unique().tolist()
        unique_values.insert(0, "")  
        filter_value = st.selectbox("Value", unique_values)
        if filter_value:
            if filter_value != "":
                filtered_output_5 = filtered_output_5[filtered_output_5[filter_col].str.replace(',', '') == filter_value]
        st.write(filtered_output_5)
        # Provide download button for the Excel file
        filtered_output_5.to_excel("ta_course_preferences.xlsx", index=False)
        with open("ta_course_preferences.xlsx", "rb") as file:
            file_data = file.read()
            st.download_button(
                label="Download this table",
                data=file_data,
                file_name="ta_course_preferences.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )   
        
    show_output_3 = st.checkbox("TAs' comments")
    if show_output_3:
        filtered_output_3 = output_3  
        filter_col = st.selectbox("Column", filtered_output_3.columns)
        unique_values = filtered_output_3[filter_col].unique().tolist()
        unique_values.insert(0, "")  
        filter_value = st.selectbox("Value", unique_values)
        if filter_value:
            if filter_value != "":
                filtered_output_3 = filtered_output_3[filtered_output_3[filter_col].str.replace(',', '') == filter_value]
        st.write(filtered_output_3)

    
    