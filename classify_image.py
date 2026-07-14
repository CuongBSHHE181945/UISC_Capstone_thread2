import os
import shutil
import pandas as pd
import re
import sys
# CẤU HÌNH HỆ THỐNG
DATASET_DIR = "Dataset từ nhóm 4"
BINARY_OUTPUT_DIR = "output_binary"
# Cấu hình hành vi sao chép/di chuyển:
# - True: Di chuyển tệp tin (tiết kiệm không gian)
# - False: Sao chép tệp tin (giữ nguyên tệp gốc trong thư mục "Đã cắt")
MOVE_FILES = False
# Cấu hình hiển thị log chi tiết từng ảnh (True/False)
VERBOSE = False
def normalize_label(label):
    """
    Chuẩn hóa nhãn về dạng Title Case để gom nhóm các từ viết hoa thường khác nhau.
    Ví dụ: 'reject' -> 'Reject', 'fetus, lesion' -> 'Fetus, Lesion'
    """
    if pd.isna(label):
        return "Unlabeled"
    
    label_str = str(label).strip()
    if not label_str:
        return "Unlabeled"
        
    # Tách nhãn bằng dấu phẩy nếu là nhãn đa lớp, viết hoa chữ cái đầu từng từ, nối lại
    parts = [part.strip().capitalize() for part in label_str.split(',') if part.strip()]
    if not parts:
        return "Unlabeled"
    return ", ".join(parts)
def match_dataset_folder(excel_filename, dataset_subdirs):
    """
    Khớp tên file Excel với thư mục dataset tương ứng:
    - Loại bỏ phần mở rộng .xlsx
    - Loại bỏ từ khóa 'đã cắt' (không phân biệt hoa thường)
    - Loại bỏ khoảng trắng và dấu gạch dưới thừa xung quanh
    """
    # Lấy tên file không có đuôi
    base_name, _ = os.path.splitext(excel_filename)
    
    # Loại bỏ từ khóa "đã cắt" (và các khoảng trắng xung quanh)
    clean_name = re.sub(r'đã\s+cắt', '', base_name, flags=re.IGNORECASE)
    
    # Loại bỏ khoảng trắng, dấu gạch dưới, gạch ngang thừa ở đầu/cuối
    clean_name = clean_name.strip(' _-')
    
    # Tìm kiếm thư mục khớp (không phân biệt chữ hoa thường)
    for folder in dataset_subdirs:
        if folder.lower() == clean_name.lower():
            return folder
            
    return None
def process_classification():
    # Cấu hình encoding utf-8 cho terminal Windows để in tiếng Việt không lỗi
    if sys.platform.startswith('win'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass
    print("=" * 70)
    print(" BẮT ĐẦU TIẾN TRÌNH PHÂN LOẠI ẢNH TỰ ĐỘNG ")
    print("=" * 70)
    # 1. Kiểm tra thư mục dữ liệu gốc
    if not os.path.exists(DATASET_DIR):
        print(f"Lỗi: Không tìm thấy thư mục dataset gốc tại '{DATASET_DIR}'")
        return
        
    excel_dir = os.path.join(DATASET_DIR, "Dataset từ nhóm 4_excel")
    if not os.path.exists(excel_dir):
        print(f"Lỗi: Không tìm thấy thư mục chứa file Excel tại '{excel_dir}'")
        return
    # Lấy danh sách các thư mục con trong DATASET_DIR
    dataset_subdirs = [
        d for d in os.listdir(DATASET_DIR) 
        if os.path.isdir(os.path.join(DATASET_DIR, d)) and d != "Dataset từ nhóm 4_excel"
    ]
    # Quét các file Excel
    all_excel_files = [f for f in os.listdir(excel_dir) if f.endswith('.xlsx')]
    
    # Lọc chỉ lấy các file chứa chữ "đã cắt" (không phân biệt hoa thường)
    target_excel_files = [f for f in all_excel_files if 'đã cắt' in f.lower()]
    print(f"Tìm thấy tổng cộng: {len(all_excel_files)} file Excel.")
    print(f"Số file Excel có chữ 'đã cắt' (cần xử lý): {len(target_excel_files)}")
    for f in target_excel_files:
        print(f"  - {f}")
    print("-" * 70)
    # Khởi tạo thống kê
    total_images_processed = 0
    detailed_stats = {}  # {nhãn: số lượng}
    binary_stats = {"reject": 0, "non-reject": 0}
    dataset_stats = {} # {tên_dataset: số_lượng_ảnh}
    # Tạo thư mục phân loại nhị phân
    os.makedirs(os.path.join(BINARY_OUTPUT_DIR, "reject"), exist_ok=True)
    os.makedirs(os.path.join(BINARY_OUTPUT_DIR, "non-reject"), exist_ok=True)
    # Duyệt qua từng file Excel để xử lý
    for excel_file in target_excel_files:
        excel_path = os.path.join(excel_dir, excel_file)
        
        # Khớp thư mục dataset tương ứng
        matched_folder = match_dataset_folder(excel_file, dataset_subdirs)
        if not matched_folder:
            print(f"Cảnh báo: Không tìm thấy thư mục dataset khớp với file Excel: '{excel_file}'")
            continue
            
        dataset_folder_path = os.path.join(DATASET_DIR, matched_folder)
        print(f"\n[XỬ LÝ] File: '{excel_file}' \n  -> Khớp với thư mục: '{matched_folder}'")
        # Tìm thư mục con chứa ảnh đã cắt (ví dụ: "Đã cắt", "Đã Cắt")
        cropped_subfolder = None
        for item in os.listdir(dataset_folder_path):
            item_path = os.path.join(dataset_folder_path, item)
            if os.path.isdir(item_path) and 'đã cắt' in item.lower():
                cropped_subfolder = item
                break
                
        if not cropped_subfolder:
            print(f"  Cảnh báo: Không tìm thấy thư mục con 'Đã cắt' bên trong '{matched_folder}'. Sử dụng thư mục gốc của dataset con.")
            source_images_dir = dataset_folder_path
        else:
            source_images_dir = os.path.join(dataset_folder_path, cropped_subfolder)
            print(f"  -> Thư mục chứa ảnh nguồn: '{cropped_subfolder}'")
        # Đọc dữ liệu từ file Excel (dùng header=None để tự xử lý dòng tiêu đề)
        try:
            df = pd.read_excel(excel_path, header=None)
        except Exception as e:
            print(f"  Lỗi: Không thể đọc file Excel '{excel_file}': {e}")
            continue
        images_in_dataset = 0
        
        # Duyệt qua từng dòng trong Excel
        for idx, row in df.iterrows():
            if len(row) < 2:
                continue
                
            val0 = row[0]
            val1 = row[1]
            
            if pd.isna(val0):
                continue
                
            img_name = str(val0).strip()
            
            # Kiểm tra xem tên có phải định dạng ảnh hay không
            _, ext = os.path.splitext(img_name.lower())
            if ext not in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
                continue
                
            # Bỏ qua dòng tiêu đề nếu có
            if 'image_name' in img_name.lower() or 'image name' in img_name.lower():
                continue
            # Chuẩn hóa nhãn
            label = normalize_label(val1)
            
            # Đường dẫn ảnh nguồn gốc
            src_image_path = os.path.join(source_images_dir, img_name)
            
            # Kiểm tra ảnh tồn tại (hỗ trợ tìm kiếm không phân biệt hoa thường tên file trên ổ đĩa)
            actual_src_path = None
            if os.path.exists(src_image_path):
                actual_src_path = src_image_path
            else:
                if os.path.exists(source_images_dir):
                    for filename in os.listdir(source_images_dir):
                        if filename.lower() == img_name.lower():
                            actual_src_path = os.path.join(source_images_dir, filename)
                            img_name = filename # Lấy lại tên file chuẩn trên ổ đĩa
                            break
                            
            if not actual_src_path:
                if VERBOSE:
                    print(f"    [!] Không tìm thấy ảnh trên đĩa: '{img_name}' tại '{source_images_dir}'")
                continue
            # --- A. PHÂN LOẠI NHỊ PHÂN (output_binary) ---
            # Xác định lớp nhị phân
            is_reject = (label.lower() == "reject")
            binary_class = "reject" if is_reject else "non-reject"
            
            # Tên file mới kèm tiền tố tên dataset tránh trùng lặp
            binary_filename = f"{matched_folder}_{img_name}"
            dest_binary_path = os.path.join(BINARY_OUTPUT_DIR, binary_class, binary_filename)
            
            # Thực hiện copy sang thư mục nhị phân trước
            try:
                shutil.copy2(actual_src_path, dest_binary_path)
                binary_stats[binary_class] += 1
            except Exception as e:
                print(f"    Lỗi khi copy ảnh sang {binary_class}: {e}")
                continue
            # --- B. PHÂN LOẠI CHI TIẾT (Theo từng thư mục dataset con) ---
            dest_detailed_dir = os.path.join(dataset_folder_path, label)
            os.makedirs(dest_detailed_dir, exist_ok=True)
            dest_detailed_path = os.path.join(dest_detailed_dir, img_name)
            
            # Thực hiện copy hoặc di chuyển cho phân loại chi tiết
            try:
                if MOVE_FILES:
                    shutil.move(actual_src_path, dest_detailed_path)
                else:
                    shutil.copy2(actual_src_path, dest_detailed_path)
                    
                # Cập nhật thống kê
                detailed_stats[label] = detailed_stats.get(label, 0) + 1
                images_in_dataset += 1
                total_images_processed += 1
            except Exception as e:
                print(f"    Lỗi khi phân loại chi tiết ảnh '{img_name}': {e}")
                continue
            if VERBOSE:
                action_str = "Moved" if MOVE_FILES else "Copied"
                print(f"    + {action_str} '{img_name}' -> nhãn '{label}' & nhị phân '{binary_class}'")
        print(f"  => Đã phân loại thành công {images_in_dataset} ảnh trong thư mục này.")
        dataset_stats[matched_folder] = images_in_dataset
    # IN BÁO CÁO THỐNG KÊ
    print("\n" + "=" * 70)
    print(" BÁO CÁO THỐNG KÊ KẾT QUẢ PHÂN LOẠI ")
    print("=" * 70)
    print(f"Tổng số lượng ảnh đã phân loại thành công: {total_images_processed}")
    
    print("\n1. Thống kê theo thư mục Dataset con:")
    for ds_name, count in dataset_stats.items():
        print(f"   - {ds_name:<40}: {count} ảnh")
        
    print("\n2. Thống kê chi tiết theo nhãn lớp (Detailed classification):")
    for lbl, count in sorted(detailed_stats.items()):
        print(f"   - {lbl:<15}: {count} ảnh")
        
    print("\n3. Thống kê phân loại nhị phân (Binary classification):")
    print(f"   - reject         : {binary_stats['reject']} ảnh (được lưu tại '{BINARY_OUTPUT_DIR}/reject')")
    print(f"   - non-reject     : {binary_stats['non-reject']} ảnh (được lưu tại '{BINARY_OUTPUT_DIR}/non-reject')")
    print("=" * 70)
    print("TIẾN TRÌNH HOÀN THÀNH CÔNG!")
    print("=" * 70)
if __name__ == '__main__':
    process_classification()
