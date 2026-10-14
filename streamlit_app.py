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
# BỔ SUNG: Bộ nhớ đệm lưu vết người được click gần nhất và số lần click liên tiếp
if 'nguoi_click_gan_nhat' not in st.session_state:
    st.session_state.nguoi_click_gan_nhat = ""
if 'so_lan_click_lien_tiep' not in st.session_state:
    st.session_state.so_lan_click_lien_tiep = 0

# 2. KHU VỰC CẤU HÌNH CHUNG
st.header("1. Thông tin cấu hình chung")
col1, col2, col3 = st.columns(3)

with col1:
    vong = st.selectbox("Chọn vòng tín nhiệm:", ["Vòng đề cử", "Vòng 1", "Vòng 2", "Vòng 3"])
with col2:
    phiu_hop_le = st.number_input("Số phiếu hợp lệ:", min_value=0, value=100, step=1)
with col3:
    phiu_trang = st.number_input("Số phiếu trắng (cũng hợp lệ):", min_value=0, value=0, step=1)

tong_mau_so = phiu_hop_le + phiu_trang

# 3. TẢI DANH SÁCH ỨNG VIÊN
st.write("---")
st.header("2. Quản lý danh sách ứng viên")
uploaded_file = st.file_uploader("Tải lên file Excel danh sách ứng viên (Cột đầu tiên là Họ và Tên):", type=["xlsx", "xls"])

if uploaded_file and not st.session_state.danh_sach:
    try:
        df = pd.read_excel(uploaded_file)
        ten_ung_vien = df.iloc[:, 0].dropna().tolist()
        st.session_state.danh_sach = [
            {"Ten": ten, "SoPhieu": 0, "Vong": vong} for ten in ten_ung_vien
        ]
        st.success(f"Đã tải thành công {len(ten_ung_vien)} ứng viên từ file Excel!")
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")

if st.button("Tạo dữ liệu thử nghiệm (10 ứng viên mẫu)"):
    st.session_state.danh_sach = [
        {"Ten": f"Nguyễn Văn {chr(65+i)}", "SoPhieu": 0, "Vong": vong} for i in range(10)
    ]
    st.session_state.danh_sach.extend([
        {"Ten": "Nguyễn Văn Tuấn", "SoPhieu": 0, "Vong": vong},
        {"Ten": "Trần Minh Hoàng", "SoPhieu": 0, "Vong": vong},
        {"Ten": "Lê Thị Diễm", "SoPhieu": 0, "Vong": vong}
    ])

# 4. KHU VỰC NHẬP LIỆU SIÊU TỐC KHÔNG DẤU & ĐẾM CHUỖI THỜI GIAN THỰC
if st.session_state.danh_sach:
    st.write("---")
    st.header("3. Kiểm phiếu siêu tốc (Gõ KHÔNG DẤU để tìm nhanh)")
    
    tu_khoa = st.text_input("Gõ tên ứng viên cần thêm phiếu (Không cần gõ dấu tiếng Việt):", value=st.session_state.tu_khoa_tim)
    
    if tu_khoa:
        tu_khoa_khong_dau = unidecode(tu_khoa).lower().strip()
        
        goi_y_khop = []
        for uv in st.session_state.danh_sach:
            ten_khong_dau = unidecode(uv["Ten"]).lower()
            if tu_khoa_khong_dau in ten_khong_dau:
                goi_y_khop.append(uv["Ten"])
        
        if goi_y_khop:
            st.write("💡 *Kết quả tìm thấy (Click chọn đúng tên để tự động cộng 1 phiếu):*")
            
            for ten_goi_y in goi_y_khop[:5]:
                if st.button(f"👤 {ten_goi_y}", key=f"btn_{ten_goi_y}", use_container_width=True):
                    # Tăng phiếu trong danh sách chính
                    for uv in st.session_state.danh_sach:
                        if uv["Ten"] == ten_goi_y:
                            uv["SoPhieu"] += 1
                            uv["Vong"] = vong
                    
                    # Logic đếm số lần click liên tiếp theo thời gian thực
                    if st.session_state.nguoi_click_gan_nhat == ten_goi_y:
                        st.session_state.so_lan_click_lien_tiep += 1
                    else:
                        st.session_state.nguoi_click_gan_nhat = ten_goi_y
                        st.session_state.so_lan_click_lien_tiep = 1
                    
                    # Bắn thông báo cập nhật số lần thực tế
                    if st.session_state.so_lan_click_lien_tiep == 1:
                        st.toast(f"✅ Đã cộng 1 phiếu cho: {ten_goi_y}!", icon="🎉")
                    else:
                        st.toast(f"🔥 Đã cộng {st.session_state.so_lan_click_lien_tiep} phiếu liên tiếp cho: {ten_goi_y}!", icon="📊")
                    
                    st.session_state.tu_khoa_tim = ""
                    st.rerun()
        else:
            st.caption("❌ Không tìm thấy ứng viên nào phù hợp với từ khóa.")

# 5. BẢNG KẾT QUẢ & KHU VỰC SỬA SAI
if st.session_state.danh_sach:
    st.write("---")
    st.header("4. Bảng xếp hạng và Công cụ Điều chỉnh số liệu")
    
    col_bang, col_sua = st.columns(2)
    
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
    
    with col_bang:
        st.subheader("Bảng Xếp Hạng Thời Gian Thực")
        st.dataframe(df_hien_thi, use_container_width=True, height=500)
        
    with col_sua:
        st.subheader("🛠️ Công cụ Sửa / Xóa")
        danh_sach_ten_sua = [uv["Ten"] for uv in st.session_state.danh_sach]
        ten_can_sua = st.selectbox("Chọn ứng viên cần sửa/xóa:", [""] + danh_sach_ten_sua, key="select_sua")
        
        if ten_can_sua:
            info_hien_tai = next(item for item in st.session_state.danh_sach if item["Ten"] == ten_can_sua)
            st.info(f"Đang chọn: **{ten_can_sua}**\n\nSố phiếu hiện tại: **{info_hien_tai['SoPhieu']}** phiếu.")
            so_phieu_moi = st.number_input("Nhập lại số phiếu đúng:", min_value=0, value=int(info_hien_tai["SoPhieu"]), step=1)
            
            col_btn_sua, col_btn_xoa = st.columns(2)
            with col_btn_sua:
                if st.button("💾 Cập nhật số phiếu", use_container_width=True, type="primary", key="save_phieu"):
                    for uv in st.session_state.danh_sach:
                        if uv["Ten"] == ten_can_sua:
                            uv["SoPhieu"] = so_phieu_moi
                            # Reset bộ đếm chuỗi khi can thiệp thủ công
                            st.session_state.nguoi_click_gan_nhat = ""
                            st.session_state.so_lan_click_lien_tiep = 0
                            st.success(f"Đã sửa số phiếu của {ten_can_sua} thành {so_phieu_moi}!")
                            st.rerun()
            with col_btn_xoa:
                if st.button("🗑️ Xóa ứng viên này", use_container_width=True, key="del_uv"):
                    st.session_state.danh_sach = [item for item in st.session_state.danh_sach if item["Ten"] != ten_can_sua]
                    st.session_state.nguoi_click_gan_nhat = ""
                    st.session_state.so_lan_click_lien_tiep = 0
                    st.warning(f"Đã xóa ứng viên {ten_can_sua} khỏi danh sách!")
                    st.rerun()
