import pandas as pd
import numpy as np
import argparse
from functools import partial


def calculate_scaled(row, stats, section):
    # (R-G1)*(M-G)/(M1-G1) + G
    slot = row['Slot']
    return (row[section.upper()] - stats[f"{slot}_{section}_g"]) * (stats[f"{section}_m"] - stats[f"{section}_g"]) / (
                stats[f"{slot}_{section}_m"] - stats[f"{slot}_{section}_g"]) + stats[f"{section}_g"]


def calculate(input_file):
    df = pd.read_excel(input_file, sheet_name='Sheet4')
    stats = dict()
    df_stats = df.describe()
    stats["varc_g"] = df_stats['VARC']['mean'] + df_stats['VARC']['std']
    stats["dilr_g"] = df_stats['DILR']['mean'] + df_stats['DILR']['std']
    stats["qa_g"] = df_stats['QA']['mean'] + df_stats['QA']['std']
    stats["varc_m"] = df.sort_values(by='VARC', ascending=False)['VARC'].head(210).mean()
    stats["dilr_m"] = df.sort_values(by='DILR', ascending=False)['DILR'].head(210).mean()
    stats["qa_m"] = df.sort_values(by='QA', ascending=False)['QA'].head(210).mean()
    for i in range(1, 4):
        df_slot = df.loc[df['Slot'] == i, :]
        df_slot_stats = df_slot.describe()
        stats[f"{i}_varc_g"] = df_slot_stats['VARC']['mean'] + df_slot_stats['VARC']['std']
        stats[f"{i}_dilr_g"] = df_slot_stats['DILR']['mean'] + df_slot_stats['DILR']['std']
        stats[f"{i}_qa_g"] = df_slot_stats['QA']['mean'] + df_slot_stats['QA']['std']
        stats[f"{i}_varc_m"] = df_slot.sort_values(by='VARC', ascending=False)['VARC'].head(70).mean()
        stats[f"{i}_dilr_m"] = df_slot.sort_values(by='DILR', ascending=False)['DILR'].head(70).mean()
        stats[f"{i}_qa_m"] = df_slot.sort_values(by='QA', ascending=False)['QA'].head(70).mean()

    calculate_varc_scaled = partial(calculate_scaled, stats=stats, section='varc')
    calculate_dilr_scaled = partial(calculate_scaled, stats=stats, section='dilr')
    calculate_qa_scaled = partial(calculate_scaled, stats=stats, section='qa')

    df['varc_scaled'] = df.apply(calculate_varc_scaled, axis=1)
    df['dilr_scaled'] = df.apply(calculate_dilr_scaled, axis=1)
    df['qa_scaled'] = df.apply(calculate_qa_scaled, axis=1)
    df['total_scaled'] = df['varc_scaled'] + df['dilr_scaled'] + df['qa_scaled']
    column_order = ["RANK", "Rnk", "VARC", "DILR", "QA", "Total", "Slot", "varc_scaled", "dilr_scaled",
                    "qa_scaled", "total_scaled"]
    df_final = df[column_order]
    return df_final


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', dest='inp_file', help='Raw input excel file', required=True)
    parser.add_argument('-o', dest='output_file', help='Output excel file', required=True)
    args = parser.parse_args()
    df = calculate(args.inp_file)
    df.to_excel(args.output_file)
    print('Done')
