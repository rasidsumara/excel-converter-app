import pandas as pd


def process_and_transform_excel(uploaded_files):
    all_rows = []

    for uploaded_file in uploaded_files:
        try:
            # Excel file read karein
            df = pd.read_excel(uploaded_file)

            # Column names clean karein
            df.columns = df.columns.astype(str).str.strip()

            # Columns by Index (0-indexed in Pandas):
            # Column C = Index 2  ('No of PCS' / Quantity values)
            # Column J = Index 9  ('TO Sales Ord.')
            # Column K = Index 10 ('TO SO item')
            # Column Z = Index 25 ('Material' / Specification / Unique ID)

            col_c = df.iloc[:, 2] if df.shape[1] > 2 else None
            col_j = df.iloc[:, 9] if df.shape[1] > 9 else None
            col_k = df.iloc[:, 10] if df.shape[1] > 10 else None
            col_z = df.iloc[:, 25] if df.shape[1] > 25 else None

            if (
                col_c is not None
                and col_j is not None
                and col_k is not None
                and col_z is not None
            ):
                temp_df = pd.DataFrame(
                    {
                        "Col_J": col_j,
                        "Col_K": col_k,
                        "Col_Z": col_z,
                        "Value_C": col_c,
                    }
                )
                all_rows.append(temp_df)

        except Exception as e:
            print(f"Error reading file {uploaded_file.name}: {e}")

    if not all_rows:
        return pd.DataFrame()

    # Sabhi files ke rows combine karein
    combined_df = pd.concat(all_rows, ignore_index=True)

    # Values ko numeric convert karein
    combined_df["Value_C"] = pd.to_numeric(
        combined_df["Value_C"], errors="coerce"
    ).fillna(0)

    # Positive aur Negative values filter logic
    combined_df["Pos_Value"] = combined_df["Value_C"].apply(
        lambda x: x if x > 0 else 0
    )
    combined_df["Neg_Value"] = combined_df["Value_C"].apply(
        lambda x: x if x < 0 else 0
    )

    # Grouping by Column Z (Col_Z duplicates merge hokar single row banenge)
    grouped_df = (
        combined_df.groupby("Col_Z", as_index=False)
        .agg(
            {
                "Col_J": "first",  # Column J data
                "Col_K": "first",  # Column K data (so item)
                "Pos_Value": "sum",  # Positive sum
                "Neg_Value": "sum",  # Negative sum
            }
        )
        .reset_index(drop=True)
    )

    # Final Columns Alignment (A, B, C, D, E)
    final_df = pd.DataFrame(
        {
            "Column A (Col J)": grouped_df["Col_J"],
            "so item": grouped_df["Col_K"],
            "Column B (Col Z)": grouped_df["Col_Z"],
            "Column C (Pos Sum)": grouped_df["Pos_Value"],
            "Column D (Neg Sum)": grouped_df["Neg_Value"],
        }
    )

    return final_df
