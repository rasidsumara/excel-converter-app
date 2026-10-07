import pandas as pd


def process_and_transform_excel(uploaded_files):
    all_rows = []

    for uploaded_file in uploaded_files:
        try:
            # Excel file read karein
            df = pd.read_excel(uploaded_file)

            # Header whitespace clean karein
            df.columns = [str(col).strip() for col in df.columns]

            # Index-based Column Mapping:
            # Column C = Index 2  (No of PCS / Quantity values)
            # Column J = Index 9  (TO Sales Ord.)
            # Column K = Index 10 (TO SO item)
            # Column Z = Index 25 (Material)

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

    # Column C value ko numeric convert karein taaki Positive/Negative SUM ho sake
    combined_df["Value_C"] = (
        pd.to_numeric(combined_df["Value_C"], errors="coerce")
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

    # Grouping by Unique combination of (Col_J, Col_K, Col_Z)
    # Isse J aur K ki unique values preserve rahengi aur duplicate rows aggregate ho jayengi
    grouped_df = (
        combined_df.groupby(["Col_J", "Col_K", "Col_Z"], as_index=False)
        .agg({"Pos_Value": "sum", "Neg_Value": "sum"})
        .reset_index(drop=True)
    )

    # Final Downloadable DataFrame Output
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
