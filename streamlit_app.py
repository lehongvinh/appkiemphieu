import streamlit as st
import pandas as pd

# 1. CẤU HÌNH TRANG WEB
st.set_page_config(page_title="Hệ thống Kiểm phiếu tự động", layout="wide")
st.title("📊 Hệ thống Kiểm phiếu Tự động")

# Khởi tạo bộ nhớ tạm cho ứng dụng (Session State)
if 'danh_sach' not in st.session_state:
    st.session_state.danh_sach = []  # Chứa danh sách ứng viên

# 2. KHU VỰC CẤU HÌNH CHUNG
st.header("1. Thông tin")
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    vong = st.selectbox("Chọn vòng tín nhiệm:", ["Vòng đề cử", "Vòng 1", "Vòng 2", "Vòng 3"])
with col2:
    phiu_phatra = st.number_input("Số phiếu phat ra:", min_value=0, value=0, step=1)
with col3:
    phiu_thuvao = st.number_input("Số phiếu thu vao:", min_value=0, value=0, step=1)
with col4:
    phiu_bathople = st.number_input("Số phiếu bat hop le:", min_value=0, value=0, step=1)
with col5:
    phiu_hop_le = st.number_input("Số phiếu hợp lệ:", min_value=0, value=100, step=1)
with col6:
    phiu_trang = st.number_input("Số phiếu trắng (cũng hợp lệ):", min_value=0, value=0, step=1)

# Tổng mẫu số để tính %
tong_mau_so = phiu_hop_le + phiu_trang

# 3. TẢI DANH SÁCH ỨNG VIÊN (Excel hoặc Nhập tay mẫu)
st.write("---")
st.header("2. Quản lý danh sách ứng viên")
uploaded_file = st.file_uploader("Tải lên file Excel danh sách ứng viên (Cột đầu tiên là Họ và Tên):", type=["xlsx", "xls"])

if uploaded_file and not st.session_state.danh_sach:
    try:
        df = pd.read_excel(uploaded_file)
        # Lấy tên từ cột đầu tiên của file Excel
        ten_ung_vien = df.iloc[:, 0].dropna().tolist()
        # Tạo cấu trúc dữ liệu cho từng người
        st.session_state.danh_sach = [
            {"Ten": ten, "SoPhieu": 0, "Vong": vong} for ten in ten_ung_vien
        ]
        st.success(f"Đã tải thành công {len(ten_ung_vien)} ứng viên từ file Excel!")
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")

# Nút tạo dữ liệu giả lập (Nếu chưa có file Excel để test)
if st.button("Tạo dữ liệu thử nghiệm (10 ứng viên mẫu)"):
    st.session_state.danh_sach = [
        {"Ten": f"Nguyễn Văn {chr(65+i)}", "SoPhieu": 0, "Vong": vong} for i in range(10)
    ]

# 4. KHU VỰC TALLY / NHẬP LIỆU NHANH (NHẮC NHỚ TÊN)
if st.session_state.danh_sach:
    st.write("---")
    st.header("3. Nhập liệu nhanh ")
    
    # Lấy danh sách tên để làm gợi ý nhắc nhớ
    danh_sach_ten = [uv["Ten"] for uv in st.session_state.danh_sach]
    
    # Ô tìm kiếm thông minh: Gõ chữ nào hiện tên người đó
    ten_chon = st.selectbox("Gõ tên ứng viên để tìm kiếm nhanh:", [""] + danh_sach_ten)
    
    if ten_chon:
        col_btn1, col_btn2 = st.columns([1, 5])
        with col_btn1:
            if st.button("➕ Cộng 1 phiếu", type="primary"):
                for uv in st.session_state.danh_sach:
                    if uv["Ten"] == ten_chon:
                        uv["SoPhieu"] += 1
                        uv["Vong"] = vong  # Cập nhật vòng hiện tại luôn
                        st.toast(f"Đã cộng 1 phiếu cho: {ten_chon}!", icon="✅")

# 5. BẢNG KẾT QUẢ THỜI GIAN THỰC & TỰ ĐỘNG XẾP HẠNG
if st.session_state.danh_sach:
    st.write("---")
    st.header("4. Bảng xếp hạng và kết quả kết quả")
    
    # Chuyển đổi dữ liệu sang dạng bảng DataFrame để tính toán
    df_kq = pd.DataFrame(st.session_state.danh_sach)
    
    # Áp dụng công thức tính %
    if tong_mau_so > 0:
        df_kq["Tỷ Lệ %"] = (df_kq["SoPhieu"] * 100 / tong_mau_so).round(2)
    else:
        df_kq["Tỷ Lệ %"] = 0.0
        
    # Tạo cột hiển thị dạng VD: 84/100
    df_kq["Số Phiếu Tín Nhiệm"] = df_kq["SoPhieu"].astype(str) + " / " + str(tong_mau_so)
    
    # Logic Đạt / Không đạt (Ví dụ tiêu chuẩn trên 50%)
    df_kq["Đánh Giá"] = df_kq["Tỷ Lệ %"].apply(lambda x: "🟢 ĐẠT" if x >= 50.0 else "🔴 CHƯA ĐẠT")
    
    # Tự động sắp xếp (Xếp hạng) theo số phiếu giảm dần
    df_kq = df_kq.sort_values(by="SoPhieu", ascending=False).reset_index(drop=True)
    df_kq.index += 1  # Để STT chạy từ 1
    
    # Chọn và sắp xếp lại các cột hiển thị giống mẫu ảnh của bạn
    df_hien_thi = df_kq[["Ten", "Số Phiếu Tín Nhiệm", "Vong", "Tỷ Lệ %", "Đánh Giá"]]
    df_hien_thi.columns = ["Họ và Tên", "Số Phiếu Tín Nhiệm", "Số Vòng", "Tỷ Lệ %", "Ghi Chú (Đánh giá)"]
    
    # Hiển thị bảng ra màn hình web
    st.dataframe(df_hien_thi, use_container_width=True)
