import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl

# Streamlit app title
st.title("Zambia Warehouse Dispatch Supervision - Data Extraction and Combined DataFrame")

# File uploader widget
uploaded_file = st.file_uploader("Choose a file", type=['xlsx'])

# If a file is uploaded
if uploaded_file is not None:
    try:
        # Load the data from the "RawData" sheet, skipping the first row (header=1)
        df = pd.read_excel(uploaded_file, sheet_name="RawData", header=1)

        # Ensure "Added Time" is in datetime format
        if "Added Time" in df.columns:
            df['Added Time'] = pd.to_datetime(df['Added Time'], errors='coerce')
        else:
            st.error("The 'RawData' sheet does not contain an 'Added Time' column.")
            st.stop()

        # Identify columns
        bag_id_column = "BAG ID." if "BAG ID." in df.columns else None
        kico_seal_column = "KICO SEAL NO." if "KICO SEAL NO." in df.columns else None
        mms_seal_column = "MMS BAG SEAL NO" if "MMS BAG SEAL NO" in df.columns else None
        mms_zambia_truck_id = "MMS ZAMBIA TRUCK ID" if "MMS ZAMBIA TRUCK ID" in df.columns else None

        # Check if the required column is present
        if bag_id_column:
            # Extract components from the Bag ID column and create new columns
            def extract_bag_info(bag_id):
                parts = dict(item.split('=') for item in bag_id.split(',') if '=' in item)
                parts.update({item.split(': ')[0]: item.split(': ')[1] for item in bag_id.split(',') if ': ' in item})
                return parts

            # Apply extraction to create new columns from Bag ID details
            bag_info_df = df[bag_id_column].dropna().apply(extract_bag_info).apply(pd.Series)

            # Concatenate original and extracted dataframes
            combined_df = pd.concat([df, bag_info_df], axis=1)

            # Create "Bag Scanned & Manual" column with specific conditions
            combined_df["Bag Scanned & Manual"] = combined_df.apply(
                lambda row: row["Bag"] if len(str(row[bag_id_column])) > 20 else row[bag_id_column],
                axis=1
            )

            # Sort combined_df by Added Time in descending order
            combined_df = combined_df.sort_values(by="Added Time", ascending=False)

            # Display combined_df with total count
            st.write(f"Total Combined DataFrame Entries: {len(combined_df)}")
            st.write("Combined DataFrame with extracted components (Sorted by Added Time):")
            st.dataframe(combined_df)

            # Exception Table 1: Duplicates in "Bag Scanned & Manual" column, consolidated into single rows
            duplicates_df = combined_df[combined_df.duplicated(subset=["Bag Scanned & Manual"], keep=False)]

            grouped_duplicates = duplicates_df.groupby("Bag Scanned & Manual").apply(
                lambda group: pd.Series({
                    "Added Time": ', '.join(sorted(group["Added Time"].astype(str).unique(), reverse=True)),
                    "Bag Scanned & Manual": group["Bag Scanned & Manual"].iloc[0],
                    "KICO SEAL NO.": ', '.join(group[kico_seal_column].dropna().unique()) if group[
                        kico_seal_column].nunique() > 1 else group[kico_seal_column].iloc[0],
                    "MMS BAG SEAL NO": ', '.join(group[mms_seal_column].dropna().unique()) if mms_seal_column and group[
                        mms_seal_column].nunique() > 1 else group[mms_seal_column].iloc[0] if mms_seal_column else None,
                    "Seal": ', '.join(group["Seal"].dropna().unique()) if group["Seal"].nunique() > 1 else
                    group["Seal"].iloc[0],
                    "Lot": ', '.join(group["Lot"].dropna().unique()) if group["Lot"].nunique() > 1 else
                    group["Lot"].iloc[0],
                    "MMS ZAMBIA TRUCK ID": ', '.join(
                        group[mms_zambia_truck_id].dropna().unique()) if mms_zambia_truck_id and group[
                        mms_zambia_truck_id].nunique() > 1 else group[mms_zambia_truck_id].iloc[0]
                })
            ).reset_index(drop=True).sort_values(by="Added Time", ascending=False)

            st.write(f"Total Duplicates in 'Bag Scanned & Manual': {len(grouped_duplicates)}")
            st.write("Duplicates Exception Table (Consolidated, Based on 'Bag Scanned & Manual'):")
            st.dataframe(grouped_duplicates)

            # Exception Table 2: "BAG ID." entries with length between 16 and 25 characters
            length_exception_df = combined_df[combined_df[bag_id_column].str.len().between(16, 25)]
            length_exception_df = length_exception_df.sort_values(by="Added Time", ascending=False)

            # Filter to show the same columns as in the duplicates exception table plus "BAG ID."
            length_exception_df_filtered = length_exception_df[[
                "Added Time", bag_id_column, "Bag Scanned & Manual", kico_seal_column, mms_seal_column, "Seal", "Lot",
                mms_zambia_truck_id
            ]]

            st.write(f"Total 'BAG ID.' Entries with Length Between 16 and 25 Characters: {len(length_exception_df_filtered)}")
            st.write("Length Exception Table (Based on 'BAG ID.' Length 16-25):")
            st.dataframe(length_exception_df_filtered)

            # Exception Table 3: Entries where "MMS ZAMBIA TRUCK ID" does not end with "_ZAM"
            if mms_zambia_truck_id:
                truck_id_exception_df = combined_df[~combined_df[mms_zambia_truck_id].str.endswith('_ZAM', na=False)]
                truck_id_exception_df = truck_id_exception_df[[
                    "Added Time", "Bag Scanned & Manual", mms_zambia_truck_id, kico_seal_column, "Seal", "Lot"
                ]]
                st.write(f"Total 'MMS ZAMBIA TRUCK ID' Entries Not Ending with '_ZAM': {len(truck_id_exception_df)}")
                st.write("'MMS ZAMBIA TRUCK ID' Exception Table (Entries Not Ending with '_ZAM'):")
                st.dataframe(truck_id_exception_df)

            # Move Date-Time Picker and Filtered Data Section Below Exceptions
            st.write("Select a date-time range to filter the Combined DataFrame:")
            start_date = st.date_input("Start Date", value=combined_df["Added Time"].min().date())
            start_time = st.time_input("Start Time", value=pd.to_datetime("00:00").time())
            end_date = st.date_input("End Date", value=datetime.now().date())  # Default to current date
            end_time = st.time_input("End Time", value=datetime.now().time())  # Default to current time

            start_datetime = pd.to_datetime(f"{start_date} {start_time}")
            end_datetime = pd.to_datetime(f"{end_date} {end_time}")

            # Slice combined_df based on the selected date-time range
            combined_df_for_download = combined_df[
                (combined_df["Added Time"] >= start_datetime) & (combined_df["Added Time"] <= end_datetime)
            ]

            # Mapping for column names in the download CSV, excluding missing columns
            column_mappings = {
                "Bag Scanned & Manual": "name",
                "KICO SEAL NO.": "GDN_KICO_SEAL",
                "MMS BAG SEAL NO": "MMS_SEAL_NO",
                "MMS ZAMBIA TRUCK ID": "ZAMBIA_TRUCK_ID",
                "BAG LOADED DATE": "GDN_LOADED_DATE",
                "DISPATCH WAREHOUSE": "GDN_WAREHOUSE_NAME",
                "RECORD BAG CONDITION": "ZAM_GDN_BAG_CONDITION_STATUS",
                "Added Email ID": "WITNESS_GDN_USER",
                "Added Time": "GDN_FORM_COMPLETE"
            }

            # Check for column existence and create final mapping
            available_columns = {key: value for key, value in column_mappings.items() if
                                 key in combined_df_for_download.columns}
            mapped_df_for_download = combined_df_for_download.rename(columns=available_columns)

            # Add missing columns as empty if they don't exist in the data
            for col in column_mappings.values():
                if col not in mapped_df_for_download.columns:
                    mapped_df_for_download[col] = None

             # Add "+02:00" to all time columns in mapped_df_for_download
            for col in ["GDN_LOADED_DATE", "GDN_FORM_COMPLETE"]:  # Specify all columns with time information
                  if col in mapped_df_for_download.columns:
                        mapped_df_for_download[col] = mapped_df_for_download[col].astype(str) + "+02:00"

            # Reorder columns according to column_mappings
            mapped_df_for_download = mapped_df_for_download[column_mappings.values()]

            st.write(f"Total Filtered Entries: {len(mapped_df_for_download)}")
            st.write("Mapped DataFrame for Download:")
            st.dataframe(mapped_df_for_download)

            # Define the filename based on start and end date-time selections
            file_name = f"From_{start_date.strftime('%Y%m%d')}_{start_time.strftime('%H%M')}_to_{end_date.strftime('%Y%m%d')}_{end_time.strftime('%H%M')}_Dispatched.csv"

            # Convert filtered data to CSV for download
            csv_data = mapped_df_for_download.to_csv(index=False)
            st.download_button(
                label="Download Filtered Combined Data as CSV",
                data=csv_data,
                file_name=file_name,
                mime="text/csv"
            )

        else:
            st.error("The file does not contain the required column: 'BAG ID.'")
    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("Awaiting file upload...")
