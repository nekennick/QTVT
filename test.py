import pandas as pd
import re
import os
import glob
from datetime import datetime

# Disable SettingWithCopyWarning
pd.options.mode.chained_assignment = None

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def get_excel_files():
    # Lấy danh sách tất cả các file Excel trong thư mục hiện tại
    excel_files = glob.glob('*.xlsx')
    return excel_files

def show_menu(files):
    clear_screen()
    print('=' * 50)
    print('Quyết Toán VT SCTX by KHOATA')
    print('=' * 50)
    print('Danh sách file Excel:')
    for idx, file in enumerate(files, 1):
        print(f'{idx}. {file}')
    print('\n0. Thoát chương trình')
    print('=' * 50)

def format_date(date):
    if pd.isna(date):
        return ''
    try:
        return pd.to_datetime(date).strftime('%d/%m/%Y')
    except:
        return ''

def format_quantity(qty):
    if pd.isna(qty):
        return ''
    if isinstance(qty, str):
        if qty == 'Yêu cầu':
            return qty
        try:
            # Chuyển đổi chuỗi số sang số thực và làm tròn thành số nguyên
            qty = round(float(qty.replace(',', '.')))
        except:
            return qty
    try:
        # Chuyển đổi sang số nguyên
        return str(round(float(qty)))
    except:
        return qty

def is_ma_phieu(s):
    pattern = re.compile(r'^(02|03)\.O09\.42\.\d{4}$')
    if isinstance(s, str):
        return bool(pattern.match(s.strip()))
    return False

def process_excel_file(file_path):
    try:
        # Đọc file Excel
        print(f'\nĐang đọc file {file_path}...')
        df = pd.read_excel(file_path, sheet_name='EVN_INV_009A___Bảng_liệt_k_0705')

        # Duyệt cột 2 để gán mã phiếu cho từng dòng
        ma_phieu_current = None
        ma_phieu_list = []

        for val in df.iloc[:, 1]:  # cột 2 trong DataFrame
            if is_ma_phieu(val):
                ma_phieu_current = val.strip()
            ma_phieu_list.append(ma_phieu_current)

        # Thêm cột "Mã phiếu" vào DataFrame
        df['Mã phiếu'] = ma_phieu_list

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

        df = df.copy()
        df['Ngày'] = df['Ngày'].apply(format_date)
        df['Số lượng'] = df['Số lượng'].apply(format_quantity)

        # Xuất ra file Excel mới
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f'Ket_qua_xu_ly_{timestamp}.xlsx'
        df.to_excel(output_file, index=False)
        
        print('\n' + '=' * 50)
        print(f'ĐÃ XUẤT KẾT QUẢ THÀNH CÔNG!')
        print(f'File kết quả: {output_file}')
        print('=' * 50)
        
        return True
    except Exception as e:
        print('\n' + '=' * 50)
        print('LƯU Ý: LÀM GỌN FILE GỐC, TRƯỚC KHI XUẤT!')
        print(f'Chi tiết lỗi: {str(e)}')
        print('=' * 50)
        return False

def main():
    while True:
        # Hiển thị menu
        excel_files = get_excel_files()
        if not excel_files:
            print('Không tìm thấy file Excel nào trong thư mục hiện tại!')
            input('Nhấn Enter để thoát...')
            break
            
        show_menu(excel_files)
        
        # Nhận lựa chọn từ người dùng
        choice = input('\nChọn file cần xử lý (nhập số): ')
        
        if choice == '0':
            print('\nCảm ơn đã sử dụng chương trình!')
            break
            
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(excel_files):
                selected_file = excel_files[idx]
                success = process_excel_file(selected_file)
                
                if success:
                    input('\nNhấn Enter để tiếp tục...')
                else:
                    input('\nNhấn Enter để thử lại...')
            else:
                print('\nLựa chọn không hợp lệ!')
                input('Nhấn Enter để thử lại...')
        except ValueError:
            print('\nVui lòng nhập số!')
            input('Nhấn Enter để thử lại...')

if __name__ == '__main__':
    main()
