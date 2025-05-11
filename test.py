import pandas as pd
import re

# Đọc file Excel
file_path = 'EVN_INV_009A___Bảng_liệt_k_070525-T2.xlsx'
df = pd.read_excel(file_path, sheet_name='EVN_INV_009A___Bảng_liệt_k_0705')

# Định nghĩa hàm kiểm tra mã phiếu
pattern = re.compile(r'^(02|03)\.O09\.42\.\d{4}$')

def is_ma_phieu(s):
    if isinstance(s, str):
        return bool(pattern.match(s.strip()))
    return False

# Duyệt cột 2 để gán mã phiếu cho từng dòng
ma_phieu_current = None
ma_phieu_list = []

for val in df.iloc[:, 1]:  # cột 2 trong DataFrame
    if is_ma_phieu(val):
        ma_phieu_current = val.strip()
    ma_phieu_list.append(ma_phieu_current)

# Thêm cột "Mã phiếu" vào DataFrame
df['Mã phiếu'] = ma_phieu_list

# Định dạng lại cột ngày thành dd/mm/yyyy
def format_date(date):
    if pd.isna(date):
        return ''
    try:
        return pd.to_datetime(date).strftime('%d/%m/%Y')
    except:
        return ''

# Tạo DataFrame kết quả với các cột cần thiết
df_ketqua = df[[
    'Mã phiếu',      # Mã phiếu
    'Ngày',           # Ngày
    'Diễn giải',      # Diễn giải
    'Mã vật tư',      # Mã vật tư
    'Tên vật tư',    # Tên vật tư
    'Đvt',            # Đơn vị tính
    'Số lượng'       # Số lượng
]]

# Áp dụng định dạng ngày mới
df_ketqua['Ngày'] = df_ketqua['Ngày'].apply(format_date)

# Xuất ra file Excel mới
from datetime import datetime

# Tạo tên file với timestamp để tránh trùng lắp
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
output_file = f'Ket_qua_xu_ly_{timestamp}.xlsx'
df_ketqua.to_excel(output_file, index=False)
print(f'Đã xuất kết quả ra file: {output_file}')

# Hiển thị 5 dòng đầu tiên để kiểm tra
print("\n5 dòng đầu tiên của kết quả:")
print(df_ketqua.head())
