import streamlit as st
import pandas as pd
from unidecode import unidecode

# 1. CẤU HÌNH TRANG WEB
st.set_page_config(page_title="Hệ thống Bầu cử & Kiểm phiếu tự động", layout="wide")
st.title("📊 Hệ thống Bầu cử & Kiểm phiếu Tự động")

# Khởi tạo bộ nhớ tạm cho ứng dụng (Session State)
if 'danh_sach' not in st.session_state:
    st.session_state.danh_sach = []  # Chứa danh sách ứng viên
if 'tu_khoa_tim' not in st.session_state:
    st.session_state.tu_khoa_tim = ""  # Từ khóa ô tìm kiếm
if 'nguoi_click_gan_nhat' not in st.session_state:
    st.session_state.nguoi_click_gan_nhat = ""
if 'so_lan_click_lien_tiep' not in st.session_state:
    st.session_state.so_lan_click_lien_tiep = 0

# 2. GỘP CHUNG: MỤC 1 - THIẾT LẬP THÔNG SỐ (6 CỘT) & TẢI DANH SÁCH ỨNG VIÊN
st.header("1. Thiết lập thông số và Tải danh sách ứng viên")

# Tạo 6 cột cho thông số chung
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    vong = st.selectbox("Chọn vòng tín nhiệm:", ["Vòng đề cử", "Vòng 1", "Vòng 2", "Vòng 3"])
with col2:
    phiu_phat_ra = st.number_input("Số phiếu phát ra:", min_value=0, value=0, step=1)
with col3:
    phiu_thu_vao = st.number_input("Số phiếu thu vào:", min_value=0, value=0, step=1)
with col4:
    phiu_bat_hop_le = st.number_input("Phiếu bất hợp lệ:", min_value=0, value=0, step=1)
with col5:
    phiu_hop_le = st.number_input("Số phiếu hợp lệ:", min_value=0, value=100, step=1)
with col6:
    phiu_trang = st.number_input("Số phiếu trắng (hợp lệ):", min_value=0, value=0, step=1)

# Tổng mẫu số thực tế dùng để tính toán %
tong_mau_so = phiu_hop_le + phiu_trang

# Khu vực upload file Excel nằm ngay bên dưới thanh cấu hình số liệu cho rộng rãi
st.write("<br>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("Tải lên danh sách ứng viên từ Excel (Cột đầu là Họ tên):", type=["xlsx", "xls"])
if uploaded_file and not st.session_state.danh_sach:
    try:
        df = pd.read_excel(uploaded_file)
        ten_ung_vien = df.iloc[:, 0].dropna().tolist()
        st.session_state.danh_sach = [
            {"Ten": ten, "SoPhieu": 0, "Vong": vong} for ten in ten_ung_vien
        ]
        st.success(f"Đã tải {len(ten_ung_vien)} ứng viên từ Excel!")
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")

# Nút test nhanh hệ thống dữ liệu mẫu
col_test, _ = st.columns([2, 8])
with col_test:
    if st.button("Tạo dữ liệu thử nghiệm"):
        st.session_state.danh_sach = [
            {"Ten": f"Nguyễn Văn {chr(65+i)}", "SoPhieu": 0, "Vong": vong} for i in range(10)
        ]
        st.session_state.danh_sach.extend([
            {"Ten": "Nguyễn Văn Tuấn", "SoPhieu": 0, "Vong": vong},
            {"Ten": "Trần Minh Hoàng", "SoPhieu": 0, "Vong": vong},
            {"Ten": "Lê Thị Diễm", "SoPhieu": 0, "Vong": vong}
        ])

# 3. MỤC 2 - BỘ ĐIỀU KHIỂN KIỂM PHIẾU & BẢNG XẾP HẠNG THỜI GIAN THỰC
if st.session_state.danh_sach:
    st.write("---")
    st.header("2. Bộ điều khiển Kiểm phiếu & Bảng xếp hạng")
    
    col_nhap_lieu, col_bxh = st.columns(2)
    
    # Xử lý dữ liệu bảng kết quả trước để hiển thị
    df_kq = pd.DataFrame(st.session_state.danh_sach)
    if tong_mau_so > 0:
        df_kq["Tỷ Lệ %"] = (df_kq["SoPhieu"] * 100 / tong_mau_so).round(2)
    else:
        df_kq["Tỷ Lệ %"] = 0.0
        
    df_kq["Số Phiếu Tín Nhiệm"] = df_kq["SoPhieu"].astype(str) + " / " + str(tong_mau_so)
    df_kq["Đánh Giá"] = df_kq["Tỷ Lệ %"].apply(lambda x: "🟢 ĐẠT" if x >= 50.0 else "🔴 CHƯA ĐẠT")
    
    df_kq = df_kq.sort_values(by="SoPhieu", ascending=False).reset_index(drop=True)
    df_kq.index += 1
    
    df_hien_thi = df_kq[["Ten", "Số Phiếu Tín Nhiệm", "Vong", "Tỷ Lệ %", "Đánh Giá"]]
    df_hien_thi.columns = ["Họ và Tên", "Số Phiếu Tín Nhiệm", "Số Vòng", "Tỷ Lệ %", "Ghi Chú (Đánh giá)"]

    # --- BÊN TRÁI: Ô NHẬP PHIẾU NHANH ---
    with col_nhap_lieu:
        st.subheader("⌨️ Ô nhập phiếu nhanh")
        tu_khoa = st.text_input("Gõ tên ứng viên (Không dấu):", value=st.session_state.tu_khoa_tim, placeholder="Ví dụ: tuan, hoang...")
        
        if tu_khoa:
            tu_khoa_khong_dau = unidecode(tu_khoa).lower().strip()
            goi_y_khop = []
            for uv in st.session_state.danh_sach:
                ten_khong_dau = unidecode(uv["Ten"]).lower()
                if tu_khoa_khong_dau in ten_khong_dau:
                    goi_y_khop.append(uv["Ten"])
            
            if goi_y_khop:
                st.write("*Click chọn đúng tên bên dưới để cộng phiếu:*")
                for ten_goi_y in goi_y_khop[:5]:
                    if st.button(f"👤 {ten_goi_y}", key=f"btn_{ten_goi_y}", use_container_width=True):
                        # Cộng phiếu
                        for uv in st.session_state.danh_sach:
                            if uv["Ten"] == ten_goi_y:
                                uv["SoPhieu"] += 1
                                uv["Vong"] = vong
                        
                        # Đếm số lần click liên tiếp
                        if st.session_state.nguoi_click_gan_nhat == ten_goi_y:
                            st.session_state.so_lan_click_lien_tiep += 1
                        else:
                            st.session_state.nguoi_click_gan_nhat = ten_goi_y
                            st.session_state.so_lan_click_lien_tiep = 1
                        
                        if st.session_state.so_lan_click_lien_tiep == 1:
                            st.toast(f"✅ Đã cộng 1 phiếu cho: {ten_goi_y}!", icon="🎉")
                        else:
                            st.toast(f"🔥 Đã cộng {st.session_state.so_lan_click_lien_tiep} phiếu liên tiếp cho: {ten_goi_y}!", icon="📊")
                        
                        st.session_state.tu_khoa_tim = ""
                        st.rerun()
            else:
                st.caption("❌ Không tìm thấy ứng viên.")
                
    # --- BÊN PHẢI: BẢNG XẾP HẠNG THỜI GIAN THỰC ---
    with col_bxh:
        st.subheader("🏆 Bảng xếp hạng thời gian thực")
        st.dataframe(df_hien_thi, use_container_width=True, height=450)

# 4. MỤC 3 - PHẦN CHỈNH SỬA & ĐIỀU CHỈNH DỮ LIỆU ĐỂ RIÊNG BIỆT
if st.session_state.danh_sach:
    st.write("---")
    st.header("3. Phần chỉnh sửa & Điều chỉnh dữ liệu")
    
    danh_sach_ten_sua = [uv["Ten"] for uv in st.session_state.danh_sach]
    ten_can_sua = st.selectbox("Chọn ứng viên cần sửa số liệu hoặc xóa hẳn khỏi danh sách:", [""] + danh_sach_ten_sua, key="select_sua")
    
    if ten_can_sua:
        info_hien_tai = next(item for item in st.session_state.danh_sach if item["Ten"] == ten_can_sua)
        col_info, col_action = st.columns(2)
        
        with col_info:
            st.info(f"Ứng viên đang chọn: **{ten_can_sua}**  \nSố phiếu hiện tại trên hệ thống: **{info_hien_tai['SoPhieu']}** phiếu.")
            so_phieu_moi = st.number_input("Nhập chính xác số phiếu thay thế:", min_value=0, value=int(info_hien_tai["SoPhieu"]), step=1)
            
        with col_action:
            st.write("<br>", unsafe_allow_html=True)
            col_btn_sua, col_btn_xoa = st.columns(2)
            with col_btn_sua:
                if st.button("💾 Lưu thay đổi số phiếu", use_container_width=True, type="primary", key="save_phieu"):
                    for uv in st.session_state.danh_sach:
                        if uv["Ten"] == ten_can_sua:
                            uv["SoPhieu"] = so_phieu_moi
                            st.session_state.nguoi_click_gan_nhat = ""
                            st.session_state.so_lan_click_lien_tiep = 0
                            st.success(f"Đã cập nhật lại số phiếu của {ten_can_sua} thành {so_phieu_moi}!")
                            st.rerun()
            with col_btn_xoa:
                if st.button("🗑️ Xóa ứng viên này", use_container_width=True, key="del_uv"):
                    st.session_state.danh_sach = [item for item in st.session_state.danh_sach if item["Ten"] != ten_can_sua]
                    st.session_state.nguoi_click_gan_nhat = ""
                    st.session_state.so_lan_click_lien_tiep = 0
                    st.warning(f"Đã xóa hoàn toàn ứng viên {ten_can_sua}!")
                    st.rerun()
