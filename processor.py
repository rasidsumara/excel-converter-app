import pandas as pd


def process_and_transform_excel(uploaded_files):
    all_rows = []

    for uploaded_file in uploaded_files:
        try:
            # Excel file read karein
            df = pd.read_excel(uploaded_file)

            # Excel Columns Clean / Strip
            df.columns = df.columns.astype(str).str.strip()

            # Pandas me 0-indexed hota hai:
            # Column J = 10th Column (Index 9) -> e.g. 'TO Sales Ord.' / Batch / Quantity
            # Column Z = 26th Column (Index 25) -> e.g. 'Material' / Coating Number
            # Column C = 3rd Column (Index 2) -> 'No of PCS' ya 'Quantity' / 'MvT'

            # Column Names identify karein (By Index ya By Header)
            col_j = df.iloc[:, 9] if df.shape[1] > 9 else None  # Column J
            col_z = df.iloc[:, 25] if df.shape[1] > 25 else None  # Column Z
            col_c = df.iloc[:, 2] if df.shape[1] > 2 else None  # Column C (Values)

            if col_j is not None and col_z is not None and col_c is not None:
                temp_df = pd.DataFrame(
                    {"Col_J": col_j, "Col_Z": col_z, "Value_C": col_c}
                )
                all_rows.append(temp_df)

        except Exception as e:
            print(f"Error reading file {uploaded_file.name}: {e}")

    if not all_rows:
        return pd.DataFrame()

    # Sabhi uploaded files ke rows ko combine karein
    combined_df = pd.concat(all_rows, ignore_index=True)

    # Values ko Numeric me convert karein
    combined_df["Value_C"] = pd.to_numeric(
        combined_df["Value_C"], errors="coerce"
    ).fillna(0)

    # Positive aur Negative values ko segregating logic:
    # Positive values -> Column C me Sum
    # Negative values -> Column D me Sum (-1, -1 = -2)
    combined_df["Pos_Value"] = combined_df["Value_C"].apply(
        lambda x: x if x > 0 else 0
    )
    combined_df["Neg_Value"] = combined_df["Value_C"].apply(
        lambda x: x if x < 0 else 0
    )

    # Column Z par Grouping (Duplicate values merged into 1 single row)
    # Grouping ke waqt Col_J ki pehli value legi, Pos_Value & Neg_Value sum honge
    grouped_df = (
        combined_df.groupby("Col_Z", as_index=False)
        .agg(
            {
                "Col_J": "first",  # Column J ka data
                "Pos_Value": "sum",  # Positive values sum
                "Neg_Value": "sum",  # Negative values sum
            }
        )
        .reset_index(drop=True)
    )

    # Output Format (New Excel Structure)
    final_df = pd.DataFrame(
        {
            "Column A (Col J)": grouped_df["Col_J"],
            "Column B (Col Z)": grouped_df["Col_Z"],
            "Column C (Pos Sum)": grouped_df["Pos_Value"],
            "Column D (Neg Sum)": grouped_df["Neg_Value"],
        }
    )

    return final_df
