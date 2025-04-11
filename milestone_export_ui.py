# milestone_export_ui.py
import pandas as pd
import streamlit as st
import io

st.title("Milestone Extractor & Exporter")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_to_df_map = {}
    milestone_outputs = {}

    for sheet_name in xls.sheet_names:
        df_full = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        rte_identifiers = df_full.iloc[0].fillna(method='ffill')
        df = df_full.drop(index=[0, 1]).reset_index(drop=True)

        all_milestones_data = []
        rte_columns = 7
        rows_per_leg = 14
        total_columns = len(df.columns)
        total_rtes = (total_columns + 1) // (rte_columns + 1)

        for rte_index in range(total_rtes):
            start_col = rte_index * (rte_columns + 1)
            end_col = start_col + rte_columns

            if start_col < total_columns:
                rte_id_prefix = rte_identifiers[start_col]
                total_legs = (len(df) + rows_per_leg - 1) // rows_per_leg

                for leg_index in range(total_legs):
                    start_row = leg_index * rows_per_leg
                    end_row = start_row + rows_per_leg
                    rte_id = f"{rte_id_prefix} Leg {leg_index + 1}"

                    current_leg_df = df.iloc[start_row:end_row, start_col:end_col]
                    current_leg_df.columns = ['Location', 'Location Type', 'Arrive (CST)', 'Depart (CST)', 'Miles',
                                              'Transit Time', 'Log-Point Time']
                    current_leg_df = current_leg_df.reset_index(drop=True)
                    current_leg_df["Row Number"] = range(1, len(current_leg_df) + 1)
                    total_rows = len(current_leg_df)
                    delivery_row = 6 if total_rows >= 6 else total_rows

                    for _, row in current_leg_df.iterrows():
                        row_number = row["Row Number"]
                        stop_type = "Pickup" if row_number < delivery_row else "Delivery"

                        if pd.notnull(row['Depart (CST)']):
                            miles_value = row["Miles"] if not pd.isnull(row["Miles"]) else df_full.iloc[1, start_col + 4]
                            all_milestones_data.append({'RTE ID': rte_id, 'Location': row['Location'],
                                                        'Milestones': f"{row['Location Type']} depart",
                                                        'Milestone Time': row['Depart (CST)'],
                                                        'Type': stop_type, 'Miles': miles_value})

                        if pd.notnull(row['Arrive (CST)']):
                            all_milestones_data.append({'RTE ID': rte_id, 'Location': row['Location'],
                                                        'Milestones': f"{row['Location Type']} arrival",
                                                        'Milestone Time': row['Arrive (CST)'],
                                                        'Type': stop_type, 'Miles': 0})

        milestones_df = pd.DataFrame(all_milestones_data)
        milestone_outputs[sheet_name] = milestones_df

    st.success("Milestones extracted for all sheets!")

    for sheet_name, df in milestone_outputs.items():
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        st.download_button(
            label=f"Download Milestones for Sheet '{sheet_name}'",
            data=output.getvalue(),
            file_name=f"Raw_Data_P04_{sheet_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
