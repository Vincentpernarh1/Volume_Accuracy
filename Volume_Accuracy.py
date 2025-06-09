import pandas as pd
from openpyxl import load_workbook
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
from tkinter import ttk
import re
import os
import time
from datetime import datetime
from PIL import Image, ImageTk  # Make sure you have Pillow installed
import sys

start_time = 0
global execution_minutes
# //output_folder 
output_folder_volume  = None

def resource_path(relative_path):
        """Get absolute path to resource, works for dev and PyInstaller .exe"""
        try:
            base_path = sys._MEIPASS  # Set by PyInstaller
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)
    

root = tk.Tk()

stellantis_logo = resource_path("assets/Vlc_img.png")

img_stellantis_logo = Image.open(stellantis_logo)
img_stellantis_logo = img_stellantis_logo.resize((530,40), Image.Resampling.LANCZOS)
photo_img_stellantis_logo = ImageTk.PhotoImage(img_stellantis_logo)
root.image = photo_img_stellantis_logo

    # Instruction label with image
image_frame = tk.Frame(root, bg="#f0f0f0")
image_frame.pack()


tk.Label(image_frame, image=photo_img_stellantis_logo, bg="#f0f0f0").pack(side="left", padx=20)


progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=10)
progress_label = tk.Label(root, text="")
progress_label.pack()


def traslate_598(trans):
    match = re.search(r"liv\.?\s*\d", trans.lower())
    if match:
        level = match.group()
    else :
        level = trans
    return level



def select_and_run_process():
    global start_time
    global output_folder_volume
    try:
        source_folder = filedialog.askdirectory(title="Select Folder with Input Files")
        if not source_folder:
            raise Exception("No source folder selected.")

        output_folder = filedialog.askdirectory(title="Select Output Folder")
        if not output_folder:
            raise Exception("No output folder selected.")
        
        output_folder_volume = os.path.join(output_folder, "Volume_Accuracy.xlsx")


        # Construct paths
        relatorio_file = os.path.join(source_folder, "Relatorio_61.xlsx")
        griglia_file = os.path.join(source_folder, "Griglia.xlsx")
        multi_de_para_file = os.path.join(source_folder, "tb_de_para.xlsx")
        start_time = time.time()
       
        # Check if files exist
        missing = [f for f in [relatorio_file, griglia_file, multi_de_para_file] if not os.path.exists(f)]
        if missing:
            raise FileNotFoundError("Missing required files in source folder.")

        # Run the main logic
        main_process(relatorio_file, griglia_file, multi_de_para_file, output_folder)
       
        progress_label.after(0, lambda: progress_label.config(text=""))
        progress_bar.stop()
        time.sleep(1)
        messagebox.showinfo("Success", f"Processamento completo. tempo de execucao: {execution_minutes} minutes")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

    finally:
        # Re-enable button and stop spinner
        run_button.config(state="normal")
        
        root.quit()
    

def run_process_thread():
    # Disable button and show spinner
    run_button.config(state="disabled")
    progress_label.config(text="Running process, please wait...")
    progress_bar.start()

    # Run actual process in a thread to keep GUI responsive
    threading.Thread(target=select_and_run_process).start()


livello_map = {
            "liv.0": "LL0",
            "liv.1": "LL1",
            "liv.2": "LL2",
            "liv.3": "LL3",
            "liv.4": "LL4",
            "liv.5": "LL5",
            "liv.6": "LL6",
            "liv.7": "LL7",
            "liv.8": "LL8",
            "liv.9": "LL9",
            "liv.10": "LL10",
            "liv.11": "LL11"
        }

map_598 = {
            "liv.0" : "Level 0",
            "liv.1" : "Level 1",
            "liv.2" : "Level 2",
            "liv.3" : "Level 3",
            "liv.4" : "Level 4",
            "liv.5" : "Level 5",
            "liv.6" : "Level 6",
            "liv.7" : "Level 7",
            "liv.8" : "Level 8",
            "liv.9" : "Level 9",
            "liv.10" :"Level 10",
            "LL0" : "Level 0",
            "LL1" : "Level 1",
            "LL2" : "Level 2",
            "LL3" : "Level 3",
            "LL4" : "Level 4",
            "LL5" : "Level 5",
            "LL6" : "Level 6",
            "LL7" : "Level 7",
            "LL8" : "Level 8",
            "LL9" : "Level 9",
            "LL10" :"Level 10",
            
        }

checkList = [
    
    "Cilindrata (l)",
    "Caratteristiche Motore",
    "Caratteristiche Cambio",
    "Famiglia Cambio",
    "Famiglia Motore",
    "Livello Allestimento",
    "Trime Level",
    "Livello Ecologia",
    "Tipo Trazione",
    "Mercati",
    "Markets",
    "Gear Family",
    "Famiglia Cambio"
]

markets = [
    "ARGENTINA",
    "BRASILE",
    "MESSICO",
    "ALTRI MERCATI",
    "BRAZIL",
    "MEXICO",
    "OTHER MARKETS"
]


def clean_markets():
    global markets
    markets = [i.strip().lower() for i in markets]
    return markets

    
def extract_cc_value(text):
  
    match = re.search(r'\(([^)]+)\)', text)
    if match:
        return match.group(1)
    return None


def exclude_if_not_exact_match(tokens, valid_set):
  
    excluded = []
    for token in tokens:
        
        if str(token).strip() not in valid_set:
            excluded.append(token)
    return excluded

def clean_volume(val):
    
    if isinstance(val, str):
        val = val.replace(",", "").strip()
    try:
        return float(val)
    except (ValueError, TypeError):
        return None

def load_dataframes(File_Rela_61, File_Griglia):
    wb_61 = load_workbook(filename=File_Rela_61, read_only=False, data_only=True)
    wb_griglia = load_workbook(filename=File_Griglia, read_only=False, data_only=True)

    if "61" not in wb_61.sheetnames:
        raise ValueError(f"Sheet '61' not found in {File_Rela_61}")

    data_61 = list(wb_61["61"].iter_rows(values_only=True))
    df_61 = pd.DataFrame(data_61[1:], columns=data_61[0]) if len(data_61) > 1 else pd.DataFrame()

    sheet_griglia = wb_griglia[wb_griglia.sheetnames[0]]
    data_griglia = list(sheet_griglia.iter_rows(values_only=True))
    df_griglia = pd.DataFrame(data_griglia[1:], columns=data_griglia[0]) if len(data_griglia) > 1 else pd.DataFrame()

    return df_61, df_griglia

def clean_griglia(df_griglia):
   
    df_griglia["Packet_cleaned"] = df_griglia["included"].astype(str).str.upper().str.replace(r'\s+', '', regex=True)
    df_griglia["Model_cleaned"] = df_griglia["Model"].astype(str).str.strip()
    df_griglia["Plant_cleaned"] = df_griglia["Plant"].astype(str).str.strip()
    df_griglia["Packet"] = df_griglia["Packet"].astype(str).str.strip()

    return df_griglia


def process_volume_table(df_61, df_griglia, field="excluded", min_col_name="Code_excluded"):
   
    results = []
    mode = "min" if field == "included" else "max"

    for _, row_61 in df_61.iterrows():
        model_61 = str(row_61.get("Modelo")).strip()
        plant_61 = str(row_61.get("Plant")).strip()
        packet_raw = str(row_61.get(field) or "")

        filtered_griglia = df_griglia[
            (df_griglia["Model_cleaned"] == model_61) &
            (df_griglia["Plant_cleaned"] == plant_61)
        ]

        packet_values = [val.replace(" ", "").strip().upper() for val in packet_raw.split(",") if val.strip()]

        unique_packet_values = set(packet_values)

        if filtered_griglia.empty:
            row_data = row_61.to_dict()
            row_data.update({
                "SINCOM": None,
                "volume_Head": None,
                "VolumeTT": None,
                min_col_name: 0
            })
            results.append(row_data)
            continue

        unique_sincom_rows = (
            filtered_griglia[['SINCOM', 'Volume Head', 'Volume TT']]
            .drop_duplicates()
            .dropna(subset=["SINCOM"])
        )

        for _, sincom_row in unique_sincom_rows.iterrows():
            sincom = sincom_row['SINCOM']
            volume_head = clean_volume(sincom_row.get('Volume Head'))
            volume_tt = clean_volume(sincom_row.get('Volume TT'))

            row_data = row_61.to_dict()
            row_data.update({
                "SINCOM": sincom,
                "volume_Head": volume_head,
                "VolumeTT": volume_tt,
                min_col_name: compute_volume_metric(unique_packet_values, sincom, filtered_griglia, mode=mode)
            })

            results.append(row_data)

   
    result_df = pd.DataFrame(results)
    preserved_cols = list(df_61.columns) + ["SINCOM", "volume_Head", "VolumeTT", min_col_name]
    preserved_cols = [col for col in preserved_cols if col in result_df.columns]
    return result_df[preserved_cols]


def compute_volume_metric(unique_packet_values, sincom, filtered_griglia, mode="max"):
   
    if not unique_packet_values:
        return 0

    volume_values = []

    for packet_code in unique_packet_values:

        if len(packet_code) == 1:
           packet_code = "00"+packet_code
        elif   len(packet_code) == 2:
            packet_code = "0"+packet_code
        

        # Step 1: Get exact Code match (always keep this)
        matched_by_code = filtered_griglia[
            (filtered_griglia["SINCOM"] == sincom) &
            (filtered_griglia["Code"].str.contains(packet_code, na=False, regex=False))
        ]

        # Step 2: Get additional Packet_cleaned match
        matched_by_packet = filtered_griglia[
            (filtered_griglia["SINCOM"] == sincom) &
            (filtered_griglia["Code"] != packet_code) &
            (filtered_griglia["Packet_cleaned"].str.lower().str.startswith("pack", na=False)) &
            (filtered_griglia["Packet_cleaned"].str.lower().str.contains(str(packet_code).lower(), na=False, regex=False))
        ]

       # Concatenate the two DataFrames
        matched_packets = pd.concat([matched_by_code, matched_by_packet])

        matched_packets['Volume'] = pd.to_numeric(matched_packets['Volume'], errors='coerce')
        matched_packets = matched_packets.dropna(subset=['Volume'])

        if not matched_packets.empty:
            matched_packets = matched_packets.loc[[matched_packets['Volume'].idxmax()]]
            
        else:
            # print("No valid Volume data to compute maximum.")
            continue

    
        if matched_packets.index.empty:
            if mode == "min":
                return 0  # In "min" mode, all packets must match
            else:
                continue  # In "max" mode, skip missing packets

        raw_volume = matched_packets.iloc[0].get("Volume")

        if len(matched_packets) > 1 and pd.notnull(matched_packets.iloc[1].get("Volume")):
            raw_volume = matched_packets.iloc[1].get("Volume")

        volume = clean_volume(raw_volume)
        if volume is None or volume == "":
            if mode == "min":
                return 0
            else:
                continue

        volume_values.append(volume)

    if not volume_values:
        return 0

    if mode == "min":
        # Only return min if we matched *all* unique_packet_values
        return min(volume_values) if len(volume_values) == len(unique_packet_values) else 0
    else:
        # In "max" mode, return max of what's available (even if partial)
        return max(volume_values)




def normalize_multivalues(raw):
    normalized = []
    pattern = re.compile(r"(\w+)\(([^)]+)\)([+-])")
    matches = pattern.findall(raw)
    for prefix, values, sign in matches:
        items = [v.strip() for v in values.split(",")]
        for item in items:
            normalized.append(f"{prefix}({item}){sign}")
    return normalized

def extract_and_save_structured_data(df_61, mapping_file_path, df_griglia):
    """
    Extracts and processes structured data from a DataFrame, handling multi-value tokens and volume calculations.
    
    This function performs a complex data transformation process that includes:
    - Normalizing multi-value tokens from a mapping file
    - Identifying included and excluded values for each row
    - Calculating volume mix based on included and excluded values
    - Grouping and aggregating data by specific columns
    
    Args:
        df_61 (pd.DataFrame): Input DataFrame to be processed
        mapping_file_path (str): Path to the Excel mapping file
        df_griglia (pd.DataFrame): Reference DataFrame for additional filtering
    
    Returns:
        pd.DataFrame: Processed DataFrame with volume mix and aggregated results
    """
    wb_mapping = load_workbook(mapping_file_path, data_only=True)
    df_mapping = pd.DataFrame(wb_mapping["Coded"].values)
    df_mapping.columns = df_mapping.iloc[0]
    df_mapping = df_mapping[1:].reset_index(drop=True)
    df_mapping["MultiValues"] = df_mapping["MultiValues"].astype(str).str.strip()
   
    df_61["Multi_included"] = ""
    df_61["Multi_excluded"] = ""

    multivalues_col = next((col for col in df_61.columns if str(col).strip().lower() == "multivalues"), None)
    if not multivalues_col:
        print("❌ Could not find a 'Multivalues' column.")
        return


    for idx, row in df_61.iterrows():
        raw = str(row.get(multivalues_col) or 0)
        tokens = normalize_multivalues(raw)

        plant = str(row["Plant"]).strip()
        model = str(row["Modelo"]).strip()
        sincom = row["SINCOM"]

        model_str = str(model).strip()
        plant_str = str(plant).strip()

        included, excluded = set(), set()
        packets_handled = set()

        # This is per row in your df_61
        for token in tokens:
            sign = token[-1]
            token_base = token[:-1].strip()

            match_row = df_mapping[df_mapping["MultiValues"] == token_base]

            translated = match_row["Resp.1"].values[0] if not match_row.empty else token_base

            # Special translation rules
            if (plant == "FIAPE" and model in ["226", "291","281"]) and "L" in token_base:
                translated = translated.lower().replace(" ", "").strip()
                translated = livello_map.get(translated, translated)
            if plant == "FIAPE" and model == "521" and token == "liv.5":
                translated = "LL5"
            
            if (plant == "FIAPE" and model in ["598","551"]) and "L" in token_base[0]:
                translated = translated.lower().replace(" ", "").strip()
                translated = map_598.get(translated, translated)

               
            # Get packet name
            packet_ita = str(match_row["Griglia Italiano"].values[0]).strip() if not match_row.empty else token_base
            packet_eng = str(match_row["Griglia Inglês"].values[0]).strip() if not match_row.empty else token_base

            # Avoid checking the same packet more than once
            if (packet_ita, packet_eng) in packets_handled:
                continue
            packets_handled.add((packet_ita, packet_eng))

            lang = False
            # Step 1: Try filtering using the Italian packet
            temp_df = df_griglia[
                (df_griglia["Model"].astype(str).str.strip() == model_str) &
                (df_griglia["Plant"].astype(str).str.strip() == plant_str) &
                (df_griglia["Packet"].astype(str).str.strip() == packet_ita)
            ]

            # Step 2: If no match found with Italian, fallback to English
            if temp_df.empty:
                temp_df = df_griglia[
                    (df_griglia["Model"].astype(str).str.strip() == model_str) &
                    (df_griglia["Plant"].astype(str).str.strip() == plant_str) &
                    (df_griglia["Packet"].astype(str).str.strip() == packet_eng)
                ]
                lang = True
                
            griglia_multis = set()
            for mv in temp_df["Multivalues"]:
                if pd.notna(mv):
                    griglia_multis.update(v.strip() for v in str(mv).split(","))

            # Determine all explicitly included/excluded from token list

            translated_tokens_from_input = set()
            for t in tokens:
                s = t[-1]
                base = t[:-1].strip()
                m_row = df_mapping[df_mapping["MultiValues"] == base]
                if lang :
                    trans = m_row["Resp.2"].values[0] if not m_row.empty else base
                else:
                
                    trans = m_row["Resp.1"].values[0] if not m_row.empty else base
               

                if (plant == "FIAPE"  and model in ["226", "291","281"]) and "L" in base:
                    trans = trans.lower().replace(" ", "").strip()
                    trans = livello_map.get(trans, trans)

                if ( plant == "FIASA" and model in ["281"]) and "L" in base:
                    trans = trans.lower().replace(" ", "").strip()
                    trans = livello_map.get(trans, trans)

                if plant == "FIAPE" and model == "521" and base == "liv.5":
                    trans = "LL5"

                if (plant == "FIAPE" and model in ["598","551"]) and "L" in base[0]:
                    trans = trans.lower().replace(" ", "").strip()
                    trans = traslate_598(trans)
                    trans = map_598.get(trans, trans)

                translated_tokens_from_input.add(trans)

                if s == "+":
                    included.add(trans)
                elif s == "-":
                    excluded.add(trans)

            translated_tokens_lower = {
                token.strip().lower()
                for mv in translated_tokens_from_input
                for token in str(mv).split(",")
                if token.strip()
            }


            if sign == "+":
                for mv in translated_tokens_from_input:
                    included.add(mv)
                for mv in griglia_multis:
                    if mv.lower() not in translated_tokens_lower:
                        excluded.add(mv)

            elif sign == "-":
                for mv in translated_tokens_from_input:
                    excluded.add(mv)
                for mv in griglia_multis:
                    if mv.lower() not in translated_tokens_lower:
                        included.add(mv)
                        # Don't include if it's part of the excluded token packet
                       
            # Normalize for case-insensitive comparison
            included_strs = list(map(str, included))

            excluded_strs = list(map(str, excluded))

            # Build lowercase sets for comparison
            included_lower_set = {val.lower() for val in included_strs}

            # Filter out from excluded any 
            # values that match included (case-insensitive)
            excluded_cleaned = [val for val in excluded_strs if val.lower() not in included_lower_set]

            # Save to DataFrame, preserving original casing
            df_61.at[idx, "Multi_included"] = ",".join(sorted(included_strs))
            df_61.at[idx, "Multi_excluded"] = ",".join(sorted(excluded_cleaned))



    df_61 = map_multi_included_to_griglia(df_61, df_griglia) 

    df_61 = map_multi_excluded_to_griglia(df_61, df_griglia)


    # Garantir que as colunas relevantes são convertidas corretamente
    df_61["Code_excluded"] = pd.to_numeric(df_61["Code_excluded"], errors="coerce").fillna(0).astype(float)
    df_61["multi_excluded_max_volume"] = pd.to_numeric(df_61["multi_excluded_max_volume"], errors="coerce").fillna(0).astype(float)

    df_61["Final_Volume_Excluded"] = np.maximum(
        df_61["Code_excluded"],
        df_61["multi_excluded_max_volume"]
    )

    df_61["Code_included"] = pd.to_numeric(df_61["Code_included"], errors="coerce").astype(float)
    df_61["multi_included_min_volume"] = pd.to_numeric(df_61["multi_included_min_volume"], errors="coerce").astype(float)

    # Garantir tipo correto da coluna volume_Head
    df_61["volume_Head"] = pd.to_numeric(df_61["volume_Head"], errors="coerce").astype(float)

    # Step 1: Initialize flag column for rows where fallback was applied
    df_61["Used_Fallback_HeadVolume"] = False

    # Step 1: Check if required included-related columns exist\\
    if all(col in df_61.columns for col in ["included", "Code_included", "multi_included_min_volume"]):
        
        
        # # Step 1.5: Apply minimum condition for valid included cases
        # mask_min = df_61["included"].notna() & (df_61["Code_included"] != 0)&(df_61["multi_included_min_volume"] !=0) &~(
        #             (df_61["multivalues"].astype(str).str.strip().str.lower().isin(["none", ""])) |
        #             (df_61["multivalues"].fillna("").astype(str).str.strip() == "")
        #         )
        
        
        # df_61.loc[mask_min, "Final_Volume_include"] = np.minimum(
        #     df_61.loc[mask_min, "Code_included"],
        #     df_61.loc[mask_min, "multi_included_min_volume"]
        # )


        # # Base rule to compute Final_Volume_include
        # mask_include_fail = df_61["included"].notna() & (df_61["Code_included"].fillna(0) == 0)
        # df_61["Final_Volume_include"] = np.where(
        #     mask_include_fail,
        #     0,
        #     np.maximum(
        #         df_61["Code_included"].fillna(0),
        #         df_61["multi_included_min_volume"].fillna(0)
        #     )
        # ).astype(float)
        





        # Step 1: Base rule - compute Final_Volume_include using maximum as default
        mask_include_fail = df_61["included"].notna() & (df_61["Code_included"].fillna(0) == 0)
        df_61["Final_Volume_include"] = np.where(
            mask_include_fail,
            0,
            np.maximum(
                df_61["Code_included"].fillna(0),
                df_61["multi_included_min_volume"].fillna(0)
            )
        ).astype(float)



        # Step 2: Override with minimum condition when both values are non-zero
        mask_min = (
            df_61["included"].notna() & 
            (df_61["Code_included"] != 0) &
            (df_61["multi_included_min_volume"] != 0) &
            ~(
                (df_61["multivalues"].astype(str).str.strip().str.lower().isin(["none", ""])) |
                (df_61["multivalues"].fillna("").astype(str).str.strip() == "")
            )
        )

        df_61.loc[mask_min, "Final_Volume_include"] = np.minimum(
            df_61.loc[mask_min, "Code_included"],
            df_61.loc[mask_min, "multi_included_min_volume"]
        )


        # Create mask to check conditions
        mask_include_condition = (
            # Check if any comma-separated value in Multi_included is in markets
            df_61["Multi_included"].fillna("").astype(str).apply(
                lambda x: any(val.strip().upper() in [market.upper() for market in markets] 
                            for val in x.split(",") if val.strip())
            ) &
            # AND condition: included is not null OR Code_included == 0
            (df_61["included"].notna() | (df_61["Code_included"].fillna(0) == 0))
        )


        # Only for rows that meet the condition
        df_61.loc[mask_include_condition, "Final_Volume_include"] = np.maximum(
            df_61.loc[mask_include_condition, "Code_included"].fillna(0),
            df_61.loc[mask_include_condition, "multi_included_min_volume"].fillna(0)
        ).astype(float)


        # Step 2: Add fallback logic to use volume_Head if both included and excluded failed
        if all(col in df_61.columns for col in ["excluded", "Code_excluded", "volume_Head"]):
            fallback_mask = (
                (df_61["Final_Volume_include"] == 0) &
                (
                    (df_61["multivalues"].astype(str).str.strip().str.lower().isin(["none", ""])) |
                    (df_61["multivalues"].fillna("").astype(str).str.strip() == "")
                ) &
                (
                    (df_61["included"].fillna("").astype(str).str.strip().str.lower() == "none") |
                    (df_61["included"].fillna("").astype(str).str.strip() == "") |
                    (df_61["included"].fillna("").astype(str).str.strip().str.lower() == "0")
                ) &
                df_61["excluded"].notna() &
                df_61["volume_Head"].notna()
            )
            
            # Apply fallback logic
            df_61.loc[fallback_mask, "Final_Volume_include"] = (
                df_61.loc[fallback_mask, "volume_Head"].astype(float) -
                df_61.loc[fallback_mask, "Code_excluded"].astype(float)
            )
            
            # Track which rows used the fallback
            df_61.loc[fallback_mask, "Used_Fallback_HeadVolume"] = True
        else:
            print("⚠️ Missing one of the fallback columns: 'excluded', 'Code_excluded', or 'volume_Head'")
    else:
        print("⚠️ Missing one of the required columns: 'included', 'Code_included', or 'multi_included_min_volume'")

        # Save intermediate file

    # Ensure numeric types
    df_61["Final_Volume_include"] = pd.to_numeric(df_61["Final_Volume_include"], errors="coerce")
    df_61["Final_Volume_Excluded"] = pd.to_numeric(df_61.get("Final_Volume_Excluded", 0), errors="coerce").fillna(0)

    # ✅ Calculate Volume_Mix per row depending on fallback
    df_61["Volume_Mix"] = np.where(
        df_61["Used_Fallback_HeadVolume"],
        df_61["Final_Volume_include"],
        np.maximum(df_61["Final_Volume_include"] - df_61["Final_Volume_Excluded"], 0)
    )

    df_61.to_excel("After_included_NEWCODE.xlsx", index=False)

    # Ensure Volume_Mix is numeric
    df_61["Volume_Mix"] = pd.to_numeric(df_61["Volume_Mix"], errors="coerce").astype(float)


    # Select only final columns (including group_key)
    final_columns = ["Modelo", "PN", "Plant", "multivalues", "included", "excluded","Unique_Key", "VolumeTT"]
    sum_columns = ["Volume_Mix"]
    # df_final is your cleaned, grouped DataFrame
      # Agrupar e somar
    df_61 = df_61.groupby(final_columns, dropna=False)[sum_columns].sum().reset_index()


    # Substituir Volume_Mix por VolumeTT quando todas as colunas estiverem vazias
    all_empty_mask = (
        (df_61["multivalues"].fillna("").astype(str).str.strip().str.lower() == "none") &
        (df_61["included"].fillna("").astype(str).str.strip() == "") &
        (df_61["excluded"].fillna("").astype(str).str.strip() == "")
    )
    df_61.loc[all_empty_mask, "Volume_Mix"] = df_61.loc[all_empty_mask, "VolumeTT"].astype(float)

    # Calcular Mix final
    df_61["Mix"] = df_61["Volume_Mix"].astype(float) / df_61["VolumeTT"].astype(float)
    end_time = time.time()

    df_61.to_excel(output_folder_volume , index=False)

    end_datetime = datetime.fromtimestamp(end_time).strftime('%d-%m-%y %H:%M:%S')
    global execution_minutes 
    execution_minutes= round((end_time - start_time) / 60, 2)

    print(f"=====================================================================")
    print(f"\n        Tempo de execução : {execution_minutes*60} segundos /  {execution_minutes} minutos\n")
    print(f"=====================================================================")

def map_multi_included_to_griglia(df_flattened, df_griglia):
    import numpy as np

    # Clean Griglia
    df_griglia["Model_cleaned"] = df_griglia["Model"].astype(str).str.strip()
    df_griglia["Plant_cleaned"] = df_griglia["Plant"].astype(str).str.strip()
    df_griglia["Code_cleaned"] = df_griglia["Code"].astype(str).str.strip().str.lower()
    df_griglia["Multivalues_cleaned_list"] = df_griglia["Multivalues"].astype(str).str.lower().str.replace(r'\s+', '', regex=True).str.split(",")
    df_griglia["Packet"] = df_griglia["Packet"].astype(str)

   
    # Clean input
    df_flattened["Modelo"] = df_flattened["Modelo"].astype(str).str.strip()
    df_flattened["Plant"] = df_flattened["Plant"].astype(str).str.strip()
    df_flattened["PN"] = df_flattened["PN"].astype(str).str.strip()
    df_flattened["SINCOM"] = df_flattened["SINCOM"].astype(str).str.strip()
    

    null_like = ["", "none", "nan", "0"]
    min_volumes = []

    for idx, row in df_flattened.iterrows():
        model = row["Modelo"]
        plant = row["Plant"]
        sincom = row["SINCOM"]
        multi_included = str(row.get("Multi_included", "")).lower()
        included_tokens = [val.strip().replace(" ", "") for val in multi_included.split(",") if val.strip()]

        matched_volumes = []

        # Filter Griglia for current model/plant/null-like code
        filtered_griglia = df_griglia[
            (df_griglia["Model_cleaned"] == model) &
            (df_griglia["Plant_cleaned"] == plant) 
            # &(df_griglia["Code_cleaned"].isin(null_like))
        ]

       
        for token in included_tokens:
            original_token = token

            # Apply special mapping for FIAPE
            if plant == "FIAPE" and model in ["226", "291", "281"]:
                token = livello_map.get(token, token)
            if plant == "FIAPE" and model == "521" and token == "liv.5":
                token = livello_map.get(token, token)


            if token in clean_markets():
                # Match from tokenized list 
    
                token_griglia_rows = filtered_griglia[
    
                    filtered_griglia["included"].str.lower().str.replace(" ", "").str.contains(token.lower(), na=False, regex=False)
                ]

            # Apply special mapping for FIAPE
            elif len(token) <= 3:
                # Match from tokenized list
                token_griglia_rows = filtered_griglia[
                    filtered_griglia["Multivalues_cleaned_list"].apply(lambda x: token in x if isinstance(x, list) else False)
                ]
            else:
                # Fallback contains
                token_griglia_rows = filtered_griglia[
                    filtered_griglia["Multivalues"].str.lower().str.replace(" ", "").str.contains(token, na=False, regex=False)
                ]


            for _, gr_row in token_griglia_rows.iterrows():
                sincom_val = str(gr_row["SINCOM"]).strip()
                volume_head = gr_row.get("Volume Head", "")

                if sincom_val == sincom:
                    try:
                        volume_val = float(volume_head)
                        matched_volumes.append(volume_val)
                    except (ValueError, TypeError):
                        continue

        # Append min or NaN
        min_volumes.append(min(matched_volumes) if matched_volumes else np.nan)

    df_flattened["multi_included_min_volume"] = min_volumes
    print("✅ Updated with 'multi_included_min_volume' using improved matching.")
    return df_flattened


def map_multi_excluded_to_griglia(df_flattened, griglia_path):
    import numpy as np

    # Clean griglia columns
    griglia_path["Model_cleaned"] = griglia_path["Model"].astype(str).str.strip()
    griglia_path["Plant_cleaned"] = griglia_path["Plant"].astype(str).str.strip()
    griglia_path["Code_cleaned"] = griglia_path["Code"].astype(str).str.strip().str.lower()
    griglia_path["Multivalues_cleaned"] = griglia_path["Multivalues"].astype(str).str.lower()

    # Explode multivalues for accurate token-level matching
    griglia_exploded = griglia_path.copy()
    griglia_exploded["Multivalue_token"] = griglia_exploded["Multivalues_cleaned"].str.split(",").apply(lambda x: [t.strip() for t in x])
    griglia_exploded = griglia_exploded.explode("Multivalue_token").reset_index(drop=True)

    # Clean base dataframe
    df_flattened["Modelo"] = df_flattened["Modelo"].astype(str).str.strip()
    df_flattened["Plant"] = df_flattened["Plant"].astype(str).str.strip()
    df_flattened["PN"] = df_flattened["PN"].astype(str).str.strip()
    df_flattened["SINCOM"] = df_flattened["SINCOM"].astype(str).str.strip()
    df_flattened["multivalues"] = df_flattened["multivalues"].astype(str).str.strip()

    # Pre-fill result column
    df_flattened["multi_excluded_volumes"] = [[] for _ in range(len(df_flattened))]

    null_like = ["", "none", "nan", "0"]

    for idx, row in df_flattened.iterrows():
        model = row["Modelo"]
        plant = row["Plant"]
        sincom = row["SINCOM"]
        multi_excluded = str(row.get("Multi_excluded", "")).lower()
        excluded_tokens = [val.strip() for val in multi_excluded.split(",") if val.strip()]
        matched_volumes = []

        # Filter Griglia rows by model, plant, and empty code
        filtered_griglia = griglia_exploded[
            (griglia_exploded["Model_cleaned"] == model) &
            (griglia_exploded["Plant_cleaned"] == plant) 
            # & (griglia_exploded["Code_cleaned"].isn(null_like))
        ].copy()


        for token in excluded_tokens:
            # Custom mapping rules
            if plant == "FIAPE" and model in ["226", "291", "281"]:
                token = livello_map.get(token, token)
            if plant == "FIAPE" and model == "521" and token == "liv.5":
                token = livello_map.get(token, token)

            token = token.lower()

            token_griglia_rows = filtered_griglia[
                filtered_griglia["Multivalue_token"] == token
            ]

            for _, gr_row in token_griglia_rows.iterrows():
                sincom_val = str(gr_row["SINCOM"]).strip()
                volume_head = gr_row.get("Volume Head", "")

                if sincom_val == sincom:
                    try:
                        volume_val = float(volume_head)
                        matched_volumes.append(volume_val)
                    except (ValueError, TypeError):
                        continue

        df_flattened.at[idx, "multi_excluded_volumes"] = matched_volumes if isinstance(matched_volumes, list) else []

    # Compute max of excluded volumes
    def safe_max(val):
        return max(val) if isinstance(val, list) and len(val) > 0 else 0

    df_flattened["multi_excluded_max_volume"] = df_flattened["multi_excluded_volumes"].apply(safe_max)

    print("✅ Updated with 'multi_excluded_max_volume' from matched tokens.")
    return df_flattened


def main_process(file_61, file_griglia, mapping_file, final_output="final_combined_output.xlsx"):
    
    try:
       
        df_61, df_griglia = load_dataframes(file_61, file_griglia)
       
        df_griglia = clean_griglia(df_griglia)
        included_df = process_volume_table(df_61, df_griglia, field="included", min_col_name="Code_included")
        excluded_df = process_volume_table(df_61, df_griglia, field="excluded", min_col_name="Code_excluded")

        
        if excluded_df.empty and included_df.empty:
            print("❌ Both DataFrames are empty.")
            return

        merge_keys = list(df_61.columns) + ["SINCOM", "volume_Head", "VolumeTT"]
        merge_keys = list(dict.fromkeys(merge_keys))

        
        merged_df = pd.merge(excluded_df, included_df, on=merge_keys, how="outer")
        # merged_df.to_excel("Merge_cheack.xlsx",index=False)


        extract_and_save_structured_data(merged_df, mapping_file, df_griglia)
        
        # merged_df.to_excel(final_output, index=False)
        print(f"✅ Final result saved to '{final_output}' with all 61 rows and new columns preserved.")

    except Exception as e:
        print(f"❌ Error: {e}")


if __name__ == "__main__":
    
    root.title("Volume Accuracy Processor")
    root.geometry("550x300")
    root.configure(bg="#f0f0f0")

    title_label = tk.Label(root, text="MIX GENERATION", font=("Helvetica", 16, "bold"), bg="#f0f0f0")
    title_label.pack(pady=20)

    run_button = tk.Button(root, text="Select Folders and Run", command=run_process_thread, height=2, width=30,
                        bg="#4CAF50", fg="white", font=("Helvetica", 12, "bold"))
    run_button.pack(pady=20)

    progress_label = tk.Label(root, text="", font=("Helvetica", 10), bg="#f0f0f0")
    progress_label.pack(pady=(0, 10))

    root.mainloop()