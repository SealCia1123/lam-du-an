import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from functools import reduce

# Phần thêm data từ file excel và merge data
df1 = pd.read_excel('D:/lam-du-an/Phan1.xlsx')
df2 = pd.read_excel('D:/lam-du-an/Phan2.xlsx')
df3 = pd.read_excel('D:/lam-du-an/Phan3.xlsx')
df4 = pd.read_excel('D:/lam-du-an/Phan4.xlsx')

# Merge data của 3 tháng lại thành 1 dataframe tổng
dfl = [df1, df2, df3, df4]
df_merged = reduce(lambda left, right: pd.merge(left, right, how='outer'), dfl)

# Sửa lỗi 00:00 trong file excel
df_merged['Ngày'] = pd.to_datetime(df_merged['Ngày'])
df_merged['Ngày'] = df_merged['Ngày'].astype(str)

# Tính giá cổ phiếu theo ngày
df_merged['Giá mua cổ phiếu theo ngày (VNĐ)'] = (df_merged['Mua: Giá trị (tỷ VNĐ)'] * 1000000000) / df_merged[
    'Mua: Khối lượng']
df_merged['Giá bán cổ phiếu theo ngày (VNĐ)'] = (df_merged['Bán: Giá trị (tỷ VNĐ)'] * 1000000000) / df_merged[
    'Bán: Khối lượng']
df_merged = df_merged.fillna(0)

# Tính các giá trị trung bình, phương sai, độ lệch chuẩn, tứ phân vị của các cột:
# kl giá trị ròng, kl mua, kl bán (câu 4)
trung_binh = df_merged[['Giao dịch ròng: Khối lượng', 'Mua: Khối lượng', 'Bán: Khối lượng']].mean()
phuong_sai = df_merged[['Giao dịch ròng: Khối lượng', 'Mua: Khối lượng', 'Bán: Khối lượng']].var()
do_lech_chuan = df_merged[['Giao dịch ròng: Khối lượng', 'Mua: Khối lượng', 'Bán: Khối lượng']].std()
tu_phan_vi_25 = df_merged[['Giao dịch ròng: Khối lượng', 'Mua: Khối lượng', 'Bán: Khối lượng']].quantile(0.25)
tu_phan_vi_50 = df_merged[['Giao dịch ròng: Khối lượng', 'Mua: Khối lượng', 'Bán: Khối lượng']].quantile(0.5)
tu_phan_vi_75 = df_merged[['Giao dịch ròng: Khối lượng', 'Mua: Khối lượng', 'Bán: Khối lượng']].quantile(0.75)
tu_phan_vi_100 = df_merged[['Giao dịch ròng: Khối lượng', 'Mua: Khối lượng', 'Bán: Khối lượng']].quantile(1)

tham_so_khoi_luong_rong = [trung_binh['Giao dịch ròng: Khối lượng'], phuong_sai['Giao dịch ròng: Khối lượng'],
                           do_lech_chuan['Giao dịch ròng: Khối lượng'], tu_phan_vi_25['Giao dịch ròng: Khối lượng'],
                           tu_phan_vi_50['Giao dịch ròng: Khối lượng'], tu_phan_vi_75['Giao dịch ròng: Khối lượng'],
                           tu_phan_vi_100['Giao dịch ròng: Khối lượng']]

tham_so_khoi_luong_mua = [trung_binh['Mua: Khối lượng'], phuong_sai['Mua: Khối lượng'],
                          do_lech_chuan['Mua: Khối lượng'], tu_phan_vi_25['Mua: Khối lượng'],
                          tu_phan_vi_50['Mua: Khối lượng'], tu_phan_vi_75['Mua: Khối lượng'],
                          tu_phan_vi_100['Mua: Khối lượng']]

tham_so_khoi_luong_ban = [trung_binh['Bán: Khối lượng'], phuong_sai['Bán: Khối lượng'],
                          do_lech_chuan['Bán: Khối lượng'], tu_phan_vi_25['Bán: Khối lượng'],
                          tu_phan_vi_50['Bán: Khối lượng'], tu_phan_vi_75['Bán: Khối lượng'],
                          tu_phan_vi_100['Bán: Khối lượng']]

tham_so = {
    '': ['Trung bình', 'Phương sai', 'Độ lệch chuẩn', 'Tứ phân vị mức 25%', 'Tứ phân vị mức 50%', 'Tứ phân vị mức 75%',
         'Tứ phân vị mức 100%']}
df_thamso = pd.DataFrame(tham_so)
df_thamso['Khối lượng giao dịch ròng'] = tham_so_khoi_luong_rong
df_thamso['Khối lượng giao dịch mua'] = tham_so_khoi_luong_mua
df_thamso['Khối lượng giao dịch bán'] = tham_so_khoi_luong_ban

# Vẽ biểu đồ
print(df_merged)
arr = np.array([0] * len(df_merged['Ngày']))
plt.plot(df_merged['Ngày'], df_merged['Giao dịch ròng: Khối lượng'], linestyle='-.', label='Khối lượng giao dịch ròng')
plt.plot(df_merged['Ngày'], df_merged['Mua: Khối lượng'], linestyle='--', label='Khối lượng mua')
plt.plot(df_merged['Ngày'], df_merged['Bán: Khối lượng'], linestyle=':', label='Khối lượng bán')
plt.plot(df_merged['Ngày'], arr, label='Điểm mua bằng bán')
plt.ylabel('Khối lượng')
plt.xlabel('Thời gian')
plt.title('BIÊU ĐỒ XU HƯỚNG MUA, BÁN CỦA MÃ FPT')
plt.legend()
plt.show()

# Xuất ra file excel và căn chỉnh khoảng cách của column
# Phải cài thư viện xlsxwriter
df_merged_writer = pd.ExcelWriter('D:/lam-du-an/Bảng thống kê mã FPT trong 6 tháng.xlsx', engine='xlsxwriter')
df_merged.to_excel(df_merged_writer, sheet_name='Khoi ngoai', index=False, na_rep='NaN')
for column in df_merged:
    column_length = max(df_merged[column].astype(str).map(len).max(), len(column))
    col_idx = df_merged.columns.get_loc(column)
    df_merged_writer.sheets['Khoi ngoai'].set_column(col_idx, col_idx, column_length)
df_merged_writer._save()

df_thamso_writer = pd.ExcelWriter('D:/lam-du-an/Các tham số của mã FPT.xlsx', engine='xlsxwriter')
df_thamso.to_excel(df_thamso_writer, sheet_name='Sheet1', index=False, na_rep='NaN')
for column in df_thamso:
    column_length = max(df_thamso[column].astype(str).map(len).max(), len(column))
    col_idx = df_thamso.columns.get_loc(column)
    df_thamso_writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_length)
df_thamso_writer._save()
