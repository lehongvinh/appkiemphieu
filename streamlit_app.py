import streamlit as st
import pandas as pd

# 1. CẤU HÌNH TRANG WEB
st.set_page_config(page_title="Hệ thống Kiểm phiếu tự động", layout="wide")
st.title("📊 Hệ thống Bầu cử & Kiểm phiếu Tự động")

# Khởi tạo bộ nhớ tạm cho ứng dụng (Session State)
if 'danh_sach' not in st.session_state:
    st.session_state.danh_sach = []  # Chứa danh sách ứng viên
if 'luot_dem' not in st.session_state:
    st.session_state.luot_dem = 0  # Dùng để tự động reset ô nhập liệu nhanh

# 2. KHU VỰC CẤU HÌNH CHUNG
st.header("1. Thông tin chung")
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

# 4. KHU VỰC NHẬP LIỆU SIÊU TỐC (CHỌN LÀ TỰ CỘNG PHIẾU)
if st.session_state.danh_sach:
    st.write("---")
    st.header("3. Kiểm phiếu siêu tốc (Chọn tên = Tự động cộng phiếu)")
    
    danh_sach_ten = [uv["Ten"] for uv in st.session_state.danh_sach]
    
    ten_chon = st.selectbox(
        "Gõ chữ để tìm và click chọn tên ứng viên (Hệ thống sẽ tự cộng 1 phiếu ngay lập tức):", 
        [""] + danh_sach_ten,
        key=f"tally_box_{st.session_state.luot_dem}"
    )
    
    if ten_chon != "":
        for uv in st.session_state.danh_sach:
            if uv["Ten"] == ten_chon:
                uv["SoPhieu"] += 1
                uv["Vong"] = vong
                st.toast(f"🎉 Đã cộng 1 phiếu cho: {ten_chon}!", icon="✅")
        
        st.session_state.luot_dem += 1
        st.rerun()

# 5. BẢNG KẾT QUẢ & KHU VỰC SỬA SAI TRÊN CÙNG MỘT DÒNG
if st.session_state.danh_sach:
    st.write("---")
    st.header("4. Bảng xếp hạng và Công cụ Điều chỉnh số liệu")
    
    # Chia màn hình làm 2 cột: Cột trái hiện Bảng kết quả, Cột phải hiện bảng Sửa/Xóa
    col_bang, col_sua = st.columns([2, 1])
    
    # --- XỬ LÝ DỮ LIỆU ĐỂ HIỂN THỊ ---
    df_kq = pd.DataFrame(st.session_state.danh_sach)
    if tong_mau_so > 0:
        df_kq["Tỷ Lệ %"] = (df_kq["SoPhieu"] * 100 / tong_mau_so).round(2)
    else:
        df_kq["Tỷ Lệ %"] = 0.0
        
    df_kq["Số Phiếu Tín Nhiệm"] = df_kq["SoPhieu"].astype(str) + " / " + str(tong_mau_so)
    df_kq["Đánh Giá"] = df_kq["Tỷ Lệ %"].apply(lambda x: "🟢 ĐẠT" if x >= 50.0 else "🔴 CHƯA ĐẠT")
    
    # Tự động xếp hạng theo số phiếu giảm dần
    df_kq = df_kq.sort_values(by="SoPhieu", ascending=False).reset_index(drop=True)
    df_kq.index += 1
    
    df_hien_thi = df_kq[["Ten", "Số Phiếu Tín Nhiệm", "Vong", "Tỷ Lệ %", "Đánh Giá"]]
    df_hien_thi.columns = ["Họ và Tên", "Số Phiếu Tín Nhiệm", "Số Vòng", "Tỷ Lệ %", "Ghi Chú (Đánh giá)"]
    
    # --- CỘT TRÁI: HIỂN THỊ BẢNG KẾT QUẢ CHÍNH ---
    with col_bang:
        st.subheader("Bảng Xếp Hạng Thời Gian Thực")
        st.dataframe(df_hien_thi, use_container_width=True, height=500)
        
    # --- CỘT PHẢI: KHU VỰC SỬA / XÓA ---
    with col_sua:
        st.subheader("🛠️ Công cụ Sửa / Xóa")
        
        # Chọn tên người cần can thiệp sửa sai
        danh_sach_ten_sua = [uv["Ten"] for uv in st.session_state.danh_sach]
        ten_can_sua = st.selectbox("Chọn ứng viên cần sửa/xóa:", [""] + danh_sach_ten_sua, key="select_sua")
        
        if ten_can_sua:
            # Tìm thông tin hiện tại của người được chọn
            info_hien_tai = next(item for item in st.session_state.danh_sach if item["Ten"] == ten_can_sua)
            
            st.info(f"Đang chọn: **{ten_can_sua}**\n\nSố phiếu hiện tại: **{info_hien_tai['SoPhieu']}** phiếu.")
            
            # 1. Chức năng sửa số phiếu trực tiếp
            so_phieu_moi = st.number_input("Nhập lại số phiếu đúng:", min_value=0, value=int(info_hien_tai["SoPhieu"]), step=1)
            
            col_btn_sua, col_btn_xoa = st.columns(2)
            
            with col_btn_sua:
                if st.button("💾 Cập nhật số phiếu", use_container_width=True, type="primary"):
                    for uv in st.session_state.danh_sach:
                        if uv["Ten"] == ten_can_sua:
                            uv["SoPhieu"] = so_phieu_moi
                            st.success(f"Đã sửa số phiếu của {ten_can_sua} thành {so_phieu_moi}!")
                            st.rerun()
            
            # 2. Chức năng xóa hẳn ứng viên khỏi danh sách
            with col_btn_xoa:
                if st.button("🗑️ Xóa ứng viên này", use_container_width=True):
                    st.session_state.danh_sach = [item for item in st.session_state.danh_sach if item["Ten"] != ten_can_sua]
                    st.warning(f"Đã xóa ứng viên {ten_can_sua} khỏi danh sách!")
                    st.rerun()
