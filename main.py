import base64
import csv
import logging
import time
from io import BytesIO, StringIO
import json
import pandas as pd
import requests
import streamlit as st

# Set up basic logging configuration
logging.basicConfig(
    level=logging.DEBUG,  # Set logging level to DEBUG, INFO, WARNING, ERROR
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(),  # Logs to the terminal (good for Streamlit Cloud)
        #logging.FileHandler("app.log")  # Logs to a file (optional)
    ]
)

# Marketplace API setup
MARKETPLACE_BASE_URL= st.secrets["MARKETPLACE_BASE_URL"]
BOL_CLIENT_ID = st.secrets["BOL_CLIENT_ID"]
BOL_CLIENT_SECRET= st.secrets["BOL_CLIENT_SECRET"]
BOL_TOKEN_URL= st.secrets["BOL_TOKEN_URL"]
ASANA_TOKEN = st.secrets["ASANA_TOKEN"]
marketplace_name = "bol"

# Initialize session state for keeping track of file paths
if "output_file" not in st.session_state:
    st.session_state.output_file = None

def analyze_listing():
    try:
        csv_file = 'https://files.channable.com/n8wWOX9ZCS6umlM-vKHUIw==.csv'
        df = pd.read_csv(csv_file, delimiter='\t')
        logging.info(f"Successfully read CSV file {len(df)} rows found.")
        return df
    except Exception as e:
        st.error(f"Error reading CSV file  {e}")
        raise
    except Exception as e:
        logging.error(f"An unexpected error occurred during the Processing of Listing File: {e}")
        st.error("An unexpected error occurred during the Processing of Listing File")

def update_excel_with_rating(listing_df, access_token):
    filtered_data = []
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/vnd.retailer.v9+json'
    }
    logging.info("Starting to update listing file with the ratings.")
    for index, row in listing_df.iterrows():
        ean = row['EAN']  # Make sure 'EAN' matches the exact column name in your local CSV
        ean = int(ean)
        ratings_response = get_product_ratings(ean, headers)
        ratings = ratings_response.get("ratings", []) if ratings_response else []

        # Filter ratings of 1, 2, or 3 with count > 0
        valid_ratings = [r['rating'] for r in ratings if r['rating'] in [1, 2, 3] and r['count'] > 0]

        if valid_ratings:
            min_rating = min(valid_ratings)
            filtered_data.append([ean, row['sku'], row['id'], min_rating])

        logging.info(f"Processed EAN: {ean} | SKU: {row['sku']}")
        # Delay for rate limiting
        time.sleep(0.8)
    return filtered_data

def write_filtered_ratings(data):
    logging.info(f"Writing filtered ratings to filtered_ratings.csv ...")
    try:
        df = pd.DataFrame(data, columns=["ean", "sku", "id", "rating"])
        # Create a BytesIO object to save the Excel file
        output = BytesIO()
        # Write the DataFrame to an Excel file in the BytesIO object
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Seek to the beginning of the BytesIO object
        output.seek(0)
        logging.info("Filtered ratings written to CSV successfully.")
        st.session_state.output_file = output
        logging.info("Filtered ratings written to CSV successfully.")
    except Exception as e:
        st.error(f"Error writing filtered ratings to CSV: {e}")


def update_excel_with_sku_description():
    try:
        logging.info("Starting to update filtered_ratings.csv with SKU description.")
        print("Starting to update filtered_ratings.csv with SKU description.")

        # Load the Excel file from session state
        input_file = st.session_state.output_file
        csv_file = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vS_mN7-KwnH2aN-afhBMbM_1IlBylxwgJByEkQU5M3HJQuSDx8-pk3HwaJ5TOLgNeD0SGcdgHikloFK/pub?gid=788370787&single=true&output=csv'

        # Read the CSV file into a DataFrame
        df_csv = pd.read_csv(csv_file, header=2)
        df_csv['Sku code'] = df_csv['Sku code'].astype(str)

        # Read the original filtered_ratings.csv into a DataFrame
        df_excel = pd.read_excel(input_file)
        df_excel['sku'] = df_excel['sku'].astype(str)

        # Merge based on 'sku' and 'Sku code'
        merged_df = pd.merge(df_excel, df_csv[['Sku code', 'Sku description']], left_on='sku',
                             right_on='Sku code', how='left')

        # Drop the 'Sku code' column as it's redundant
        merged_df.drop(columns=['Sku code'], inplace=True)

        # Save the merged DataFrame as an Excel file
        # Create a BytesIO object to store the CSV data
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, sheet_name="Sheet1",index=False)
        logging.info("Filtered ratings written to CSV successfully.")

        output.seek(0)
        # Store the updated file in session state
        st.session_state.output_file = output
        logging.info("Successfully updated filtered_ratings file with SKU description information. Saved as filtered_ratings - Desc Added.xlsx")

    except Exception as e:
        logging.error(f"An error occurred while updating the Excel file with SKU description: {e}")
        st.error("An error occurred while updating the Excel file with SKU description")


def update_excel_with_f1_to_use():
    try:
        logging.info("Starting to update F1s - Desc Added.xlsx with F1 to Use.")
        print("Starting to update F1s - Desc Added.xlsx with F1 to Use.")

        # Load the existing Excel file from session state
        input_file = st.session_state.output_file

        # Fetch the CSV file from the URL
        url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRxBqpSTMwezeOji3KXDlrp3855sQHFuYxmKsCIDwILg4iHMEx2BBmp87nwEgI__4g3rM6H65rIp0sF/pub?gid=0&single=true&output=csv"
        response = requests.get(url)
        csv_data = StringIO(response.text)
        df_csv = pd.read_csv(csv_data)
        # Store dataframes temporarily
        df_dict = {}

        df_excel = pd.read_excel(input_file)
        df_excel['sku'] = df_excel['sku'].astype(str)

        f1_to_use_values = []
        for sku in df_excel['sku']:
            found_row = df_csv.iloc[:, 1:16].apply(
                lambda row: row.astype(str).str.contains(str(sku), na=False).any(), axis=1)

            matching_rows = df_csv[found_row]
            if not matching_rows.empty:
                last_non_empty_value = matching_rows.iloc[0, 1:16].dropna().iloc[-1]
                f1_to_use_values.append(last_non_empty_value)
            else:
                f1_to_use_values.append(None)

        df_excel['F1 to Use'] = f1_to_use_values
        df_dict['Sheet1'] = df_excel
        # Write the updated data back to a BytesIO object
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            logging.info(f"Writing updated data to sheet.")
            df_excel.to_excel(writer, sheet_name='Sheet1', index=False)

        logging.info(f"Successfully updated Excel file with F1 to Use information.")
        output.seek(0)  # Reset the pointer of the BytesIO object
        st.session_state.output_file = output
    except Exception as e:
        st.error(f"An error occurred while updating the Excel file with F1 to Use: {e}")


def update_excel_with_barcodes(uploaded_barcodes):
    try:
        logging.info("Updating filtered_ratings_with_desc_and_F1_to_use.xlsx with Barcodes.")
        print("Updating filtered_ratings_with_desc_and_F1_to_use.xlsx with Barcodes.")

        input_file = st.session_state.output_file
        df_barcodes = pd.read_csv(uploaded_barcodes, header=3)

        xls = pd.ExcelFile(input_file)
        sheet_names = xls.sheet_names

        df_dict = {}
        for sheet in sheet_names:
            logging.info(f"Processing sheet: {sheet}")
            df_excel = pd.read_excel(input_file, sheet_name=sheet)

            if 'F1 to Use' in df_excel.columns:
                barcode_values = []
                gs1_brand_values = []

                for f1 in df_excel['F1 to Use']:
                    found_row = df_barcodes[df_barcodes['SKU'] == f1]

                    if not found_row.empty:
                        number_value = str(found_row['Number'].iloc[0]).replace('=', '').replace('"', '')
                        barcode_values.append(number_value)

                        gs1_brand_value = found_row['Main Brand'].iloc[0]
                        gs1_brand_values.append(gs1_brand_value)
                    else:
                        barcode_values.append(None)
                        gs1_brand_values.append(None)

                df_excel['Barcode'] = barcode_values
                df_excel['GS1 Brand'] = gs1_brand_values
                df_dict[sheet] = df_excel
            else:
                logging.warning(f"'F1 to Use' column not found in sheet {sheet}. Skipping this sheet.")

        del xls

        # Write the updated data back to a BytesIO object
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet, df in df_dict.items():
                logging.info(f"Writing updated data to sheet {sheet}.")
                df.to_excel(writer, sheet_name=sheet, index=False)

        logging.info(f"Successfully updated {output} with Barcodes.")
        # Store the output file path in session state so it can be downloaded later
        output.seek(0)  # Reset the pointer of the BytesIO object
        st.session_state.output_file = output

    except Exception as e:
        logging.error(f"An error occurred while updating the Excel file with Barcodes: {e}")
        st.error("An error occurred while updating the Excel file with Barcodes")

def get_access_token():
    logging.info("Fetching access token...")
    credentials = f"{BOL_CLIENT_ID}:{BOL_CLIENT_SECRET}"
    encoded_credentials = base64.b64encode(credentials.encode("utf-8")).decode("utf-8")
    headers = {
        "Authorization": f"Basic {encoded_credentials}",
        "Accept": "application/json"
    }
    try:
        # Make the request to get access token
        response = requests.post(BOL_TOKEN_URL, headers=headers)

        # Check if request was successful
        if response.status_code == 200:
            token_data = response.json()
            #print(f"response {token_data}.")
            access_token = token_data['access_token']
            return access_token
        else:
            message = f"Error fetching access token: {response.status_code} - {response.text}"
            return None
    except Exception as e:
        st.error(f"Exception occurred while fetching access token: {str(e)}")
        return None

def get_product_ratings(ean, headers):
    logging.info(f"Fetching product ratings for EAN: {ean}")
    try:
        url = f"https://api.bol.com/retailer/products/{ean}/ratings"
        response = requests.get(url, headers=headers)

        # If response is 401 Unauthorized, reauthorize and retry
        if response.status_code == 401:
            logging.warning(f"401 Unauthorized error for EAN {ean}. Reauthorizing...")
            headers['Authorization'] = "Bearer " + get_access_token()
            response = requests.get(url, headers=headers)

        # If response is 404 Not Found, log and return None
        if response.status_code == 404:
            logging.warning(f"404 Not Found error for EAN {ean}. Skipping this EAN.")
            return None

        response.raise_for_status()
        logging.info(f"Successfully fetched ratings for EAN: {ean}")
        return response.json()
    except requests.exceptions.HTTPError as http_err:
        logging.error(f"HTTP error occurred for EAN {ean}: {http_err}")
        return None
    except Exception as e:
        logging.error(f"An error occurred while getting product ratings for EAN {ean}: {e}")
        return None
def create_asana_tasks_from_excel(send_to_asana=True):
    print("create_asana_tasks_from_excel")
    if not send_to_asana:
        st.info("Task creation in Asana is disabled.")
        return

    # Asana API setup
    url = "https://app.asana.com/api/1.0/tasks?opt_fields="
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "authorization": f"Bearer {ASANA_TOKEN}"
    }

    # Load the updated F1s Excel file
    input_file = st.session_state.output_file
    for sheet_name in pd.ExcelFile(input_file).sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name)

        # Check if 'EAN' column exists in the DataFrame
        if 'ean' not in df.columns:
            print("The 'EAN' column is missing in the Excel sheet.")
            continue  # Skip processing this sheet if 'EAN' is missing

        # Fetch existing tasks for the current project
        project_id = '1205436216136693'
        existing_task_names = fetch_existing_asana_tasks(project_id, headers)

        for idx, row in df.iterrows():
            ean_value = row['ean']

            # Remove any leading apostrophes if the EAN is a string
            if isinstance(ean_value, str):
                ean_value = ean_value.lstrip("'")
            # Convert float EAN values to integer and then to string, but only if it's not NaN
            elif isinstance(ean_value, float) and not pd.isna(ean_value):
                ean_value = str(int(ean_value))

            if pd.notna(ean_value) and (isinstance(ean_value, str) or isinstance(ean_value, int)):
                ean_value = str(ean_value)

                # Value is valid, proceed with task creation
                task_name = f"F1 for {row['sku']} - {row['Sku description']}"

                if task_name in existing_task_names:
                    logging.warning(
                        f"Task already exists for SKU {row['sku']} in country Netherland, skipping Asana task creation.")
                    continue
                sku_to_f1 = row['sku']
                new_f1_sku = row['F1 to Use']
                new_f1_ean = ean_value.replace("'", "")  # Remove single quotes
                new_f1_brand = row['GS1 Brand']
                projects = ['1205436216136693']
                all_skus_data.append([sku_to_f1, new_f1_sku, new_f1_ean, new_f1_brand])

                # notes_content = (f"<body><b>SKU to be F1'd:</b> {sku_to_f1}\n"
                #                  f"<b>New F1 SKU:</b> {new_f1_sku}\n"
                #                  f"<b>New F1 EAN:</b> {new_f1_ean}\n"
                #                  f"<b>New F1 Brand:</b> {new_f1_brand}\n"
                #                  "\n"
                #                  "<b>PLEASE TICK EACH ITEM ON YOUR CHECKLIST AS YOU GO</b></body>")

                # # Look up the tag ID based on the sheet name
                # tags = ['1205436216136702']
                #
                # payload = {
                #     "data": {
                #         "projects": projects,
                #         "name": task_name,
                #         "html_notes": notes_content,
                #         "tags": tags  # Use the looked-up tag ID here
                #     }
                # }
                #
                # response = requests.post(url, json=payload, headers=headers)
                # print(f"Sending payload: {payload}")
                # task_data = response.json()

                # Check if task creation was successful and move it to the appropriate section
                # if 'data' in task_data and 'gid' in task_data['data']:
                #     task_gid = task_data['data']['gid']
                #     section_id = ['1205436216136695']
                #     move_url = f"https://app.asana.com/api/1.0/sections/{section_id}/addTask"
                #     move_payload = {
                #         "data": {
                #             "task": task_gid
                #         }
                #     }
                #     move_response = requests.post(move_url, json=move_payload, headers=headers)
                #     print(f"Moved task {task_gid} to section {section_id}. Response: {move_response.json()}")
                # else:
                #     logging.error(f"Failed to create task for SKU {row['sku']} in country Netherland. Response: {task_data}")

            else:
                print(
                    f"EAN '{ean_value}' (data type: {type(ean_value)}) is not a valid value for SKU {row['sku']} in country Netherland. Skipping Asana task creation.")
                if row['sku'] not in unique_seller_skus:
                    unique_seller_skus.add(row['sku'])  # Add to the unique set

                    # Add to the list of tasks needing new EANs
                    new_eans_needed.append({
                        'Seller SKU': row['sku'],
                        'Sku description': row['Sku description']
                    })
        # Create a DataFrame for the Excel file
        df_skus = pd.DataFrame(all_skus_data, columns=['SKU to be F1', 'New F1 SKU', 'New F1 EAN', 'New F1 Brand'])

        # Save the DataFrame to an Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_skus.to_excel(writer, index=False, sheet_name='Sheet1')
        output.seek(0)
        st.info(f"CSV file created")
        # Display the DataFrame on Streamlit
        st.write("Generated SKU Details DataFrame:", df_skus)
        projects = ['1205436216136693']
        tags = ['1205436216136702']
        payload = {
            "data": {
                "projects": projects,
                "name": "Bol SKU Details CSV Upload",
                "html_notes": "Please find the attached CSV containing the details of SKUs to be F1'd",
                "tags": tags  # Use the looked-up tag ID here
            }
        }
        # Create the task on Asana
        response = requests.post(url, json=payload, headers=headers)
        task_data = response.json()
        if 'data' in task_data and 'gid' in task_data['data']:
            task_gid = task_data['data']['gid']
            st.write(f"1 {task_gid}.")

            # Upload the CSV file as an attachment to the task
            upload_url = f"https://app.asana.com/api/1.0/tasks/{task_gid}/attachments"
            files = {'file': (
            'bol_sku_details.xlsx', output, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            attach_response = requests.post(upload_url, headers=headers, files=files)
            st.write(f"2 {attach_response.status_code}.")

            if attach_response.status_code == 200:
                st.success(f"Excel file successfully attached to task {task_gid}.")
            else:
                st.error(f"Failed to upload the Excel file. Response: {attach_response.json()}")

    if new_eans_needed:
        # Create the main task
        main_task_payload = {
            "data": {
                "name": "NEW F1's Needed",
                "assignee": "1202218809921567",
                "html_notes": "<body><b>Please can the following new F1's be created and added to the F1 Log <a href=\"https://docs.google.com/spreadsheets/d/1JesoDfHewylxsso0luFrY6KDclv3kvNjugnvMjRH2ak/edit#gid=0\" target=\"_blank\">here</a></b></body>",
                "followers": ["greg.stephenson@monstergroupuk.co.uk, 1202218809921567"],
                "workspace": "17406368418784"
            }
        }
        main_task_response = requests.post(url, json=main_task_payload, headers=headers)
        main_task_data = main_task_response.json()
        main_task_gid = main_task_data['data']['gid']

        # Create subtasks
        subtask_url = f"https://app.asana.com/api/1.0/tasks/{main_task_gid}/subtasks"
        for task in new_eans_needed:
            subtask_name = f"{task['Seller SKU']} - {task['Sku description']}"
            subtask_payload = {
                "data": {
                    "name": subtask_name
                }
            }
            subtask_response = requests.post(subtask_url, json=subtask_payload, headers=headers)
            print(f"Added subtask: {subtask_name}. Response: {subtask_response.json()}")


def fetch_existing_asana_tasks(project_id, headers):
    print("fetch_existing_asana_tasks")
    url = f"https://app.asana.com/api/1.0/tasks?project={project_id}&opt_fields=name"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return [task['name'] for task in json.loads(response.text)['data']]
    else:
        return []

# Initialize an empty set to store unique seller-skus
unique_seller_skus = set()
# Initialize an empty list to store tasks with missing EANs
new_eans_needed = []
# Prepare the list to store SKU details for CSV
all_skus_data = []
def main():
    st.set_page_config(page_title="BOL File Processor", page_icon="📄")

    st.markdown(
        """
        <h1 style='text-align: center;'>
            🔄 BOL F1s
        </h1>
        """,
        unsafe_allow_html=True
    )

    st.markdown("""<style>
        .css-1offfwp {padding-top: 1rem;}
        .css-1v3fvcr {background-color: #f8f9fa !important;}
        .block-container {padding: 10rem !important;}
        .stButton button, .stDownloadButton button {background-color: #4CAF50; color: white; border: none; border-radius: 5px; padding: 10px 20px; font-size: 16px; cursor: pointer;}
        .stButton button:hover, .stDownloadButton button:hover {background-color: #45a049;}
        .stFileUploader {border: 2px dashed #4CAF50 !important; border-radius: 10px;}
        </style>""", unsafe_allow_html=True)
    # File uploader widget for the user to upload their barcodes file
    uploaded_barcodes = st.file_uploader("Upload Barcode CSV file", type="csv")

    if uploaded_barcodes is not None and st.session_state.output_file is None:
        # When a file is uploaded, run the analysis
        with st.spinner("Processing your files. This may take a few moments..."):
            listing_df = analyze_listing()
            if listing_df is not None:
                access_token = get_access_token()
                if access_token:
                    filtered_rating_data = update_excel_with_rating(listing_df,access_token)
                    if filtered_rating_data:
                        write_filtered_ratings(filtered_rating_data)
                        update_excel_with_sku_description()
                        update_excel_with_f1_to_use()
                        update_excel_with_barcodes(uploaded_barcodes)
    # Check if the output file exists and show download button
    if st.session_state.output_file is not None:
        # Use Streamlit columns to place buttons side-by-side
        col1, col2, col3 = st.columns([0.1, 1, 1])
        # Column 1: Download Button
        with col2:
            st.download_button(label="Save File", data=st.session_state.output_file, file_name="F1_Barcodes.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Column 2: Trigger Asana Functionality
        with col3:
            if st.button("Create Asana Tasks"):
                st.info("Starting Asana task creation...")
                create_asana_tasks_from_excel(send_to_asana=True)  # Call your function here
                st.success("Asana tasks created successfully!")

if __name__ == "__main__":
    main()