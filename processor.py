import pandas as pd


def process_and_transform_excel(uploaded_files):
    all_rows = []

    for uploaded_file in uploaded_files:
        try:
            # Excel read karein
            df = pd.read_excel(uploaded_file)

            # Strip column whitespace
            df.columns = [str(col).strip() for col in df.columns]

            # Index-based column selection (0-indexed):
            # Column C = Index 2  (PCS / Quantity values)
            # Column J = Index 9  (TO Sales Ord.)
            # Column K = Index 10 (TO SO item)
            # Column Z = Index 25 (Material / Unique ID)

            if df.shape[1] > 25:
                col_c = df.iloc[:, 2]
                col_j = df.iloc[:, 9]
                col_k = df.iloc[:, 10]
                col_z = df.iloc[:, 25]

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

    # Combine all data
    combined_df = pd.concat(all_rows, ignore_index=True)

    # Clean and Convert to strictly Numeric numbers for math operations
    combined_df["Value_C"] = (
        pd.to_numeric(combined_df["Value_C"], errors="coerce")
        .fillna(0)
        .astype(float)
    )
    combined_df["Col_K"] = (
        pd.to_numeric(combined_df["Col_K"], errors="coerce")
        .fillna(0)
        .astype(float)
    )

    # Positive and Negative separation for Column C values
    combined_df["Pos_Value"] = combined_df["Value_C"].apply(
        lambda x: x if x > 0 else 0
    )
    combined_df["Neg_Value"] = combined_df["Value_C"].apply(
        lambda x: x if x < 0 else 0
    )

    # Grouping by Column Z (Material Name):
    # - Col_J: Pehli row ka value (TO Sales Ord)
    # - Col_K: SUM of all SO Items (e.g., 50 + 50 = 100)
    # - Pos_Value & Neg_Value: SUM of Positive/Negative Values
    grouped_df = (
        combined_df.groupby("Col_Z", as_index=False)
        .agg(
            {
                "Col_J": "first",
                "Col_K": "sum",  # Force Sum on Column K
                "Pos_Value": "sum",
                "Neg_Value": "sum",
            }
        )
        .reset_index(drop=True)
    )

    # Resulting Clean Output Frame
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
