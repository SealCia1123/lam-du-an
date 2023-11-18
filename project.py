import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from functools import reduce

# Phần thêm data từ file excel và merge data
df1 = pd.read_excel("D:/lam-du-an/FPT_23.10_17.11.xlsx")
df2 = pd.read_excel("D:/lam-du-an/FPT_25.09_20.10.xlsx")
df3 = pd.read_excel("D:/lam-du-an/FPT_24.08_24.09.xlsx")
df4 = pd.read_excel("D:/lam-du-an/FPT_27.07_23.08.xlsx")
df5 = pd.read_excel("D:/lam-du-an/FPT_29.06_26.07.xlsx")
df6 = pd.read_excel("D:/lam-du-an/FPT_01.06_28.06.xlsx")
df7 = pd.read_excel("D:/lam-du-an/FPT_17.05_31.05.xlsx")

dfl = [df1, df2, df3, df4, df5, df6, df7]

df_merged = reduce(lambda left, right: pd.merge(left, right, how="outer"), dfl)

print(df_merged)

# Viết code dưới này


# Phần xuất ra file excel hoàn thiện
df_merged.to_excel("test.xlsx", index=False)
