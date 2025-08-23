import pandas as pd

df = pd.read_excel("CoA_Level.xlsx", dtype=str)
df[df["FS_Element"]=="R"].iloc[0]["L1_code"]