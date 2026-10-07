import pandas as pd


def process_excel_files(uploaded_files):
    all_data = []

    for uploaded_file in uploaded_files:
        try:
            # Excel read karein (Header row 0 par hai)
            df = pd.read_excel(uploaded_file)

            # Check karein ki required columns hain
            if df.shape[1] > 25:  # At least 26 columns (Index 0 to 25)
                # Column J -> Index 9, Column Z -> Index 25
                col_j_name = df.columns[9]  # Column J
                col_z_name = df.columns[25]  # Column Z

                # Relevant columns select karein
                temp_df = df[[col_j_name, col_z_name]].copy()

                # Column renaming: J -> A_val, Z -> B_val
                temp_df.columns = ["Col_A_Val", "Col_B_Val"]

                all_data.append(temp_df)
            else:
                # Agar column names direct Z/J se accessible hain
                col_j = df.iloc[:, 9] if df.shape[1] > 9 else None
                col_z = df.iloc[:, 25] if df.shape[1] > 25 else None

                if col_j is not None and col_z is not None:
                    temp_df = pd.DataFrame({"Col_A_Val": col_j, "Col_B_Val": col_z})
                    all_data.append(temp_df)

        except Exception as e:
            print(f"Error reading file {uploaded_file.name}: {e}")

    if not all_data:
        return pd.DataFrame()

    # Sabhi uploaded files ke data ko combine karein
    combined_df = pd.concat(all_data, ignore_index=True)

    # Values ko numeric convert karein (Rule conditions ke liye)
    combined_df["Numeric_A"] = pd.to_numeric(
        combined_df["Col_A_Val"], errors="coerce"
    ).fillna(0)

    # Positive and Negative segregation
    combined_df["Pos_Value"] = combined_df["Numeric_A"].apply(
        lambda x: x if x > 0 else 0
    )
    combined_df["Neg_Value"] = combined_df["Numeric_A"].apply(
        lambda x: abs(x) if x < 0 else 0
    )

    # Grouping by Column Z (Col_B_Val)
    # Same Z value ke liye C & D sum honge, Unique ke liye single row banegi
    grouped_df = (
        combined_df.groupby("Col_B_Val", as_index=False)
        .agg({"Pos_Value": "sum", "Neg_Value": "sum"})
        .reset_index(drop=True)
    )

    # Final Columns rearrange karein:
    # Column A: Original Value / Reference
    # Column B: Column Z Data
    # Column C: Sum of Positive Values
    # Column D: Sum of Negative Values
    final_output = pd.DataFrame(
        {
            "Column A": grouped_df["Col_B_Val"],  # Primary identifier
            "Column B (Col Z)": grouped_df["Col_B_Val"],
            "Column C (Pos Sum)": grouped_df["Pos_Value"],
            "Column D (Neg Sum)": grouped_df["Neg_Value"],
        }
    )

    return final_output
