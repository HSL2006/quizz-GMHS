"""
Script chuyển đổi file Word đề thi (.docx) sang Quiz HTML & JSON
Khoa Gây Mê Hồi Sức (GMHS)
"""

import zipfile
import xml.etree.ElementTree as ET
import re
import json
import sys
import os

# Đảm bảo in tiếng Việt không bị lỗi encoding trên Windows
sys.stdout.reconfigure(encoding='utf-8')

DOCX_FILE = r"BỘ TEST KHOA GMHS CHUẨN_1.docx"
JSON_OUTPUT = r"questions.json"
HTML_OUTPUT = r"quizz.html"

def extract_paragraphs_from_docx(file_path):
    """Trích xuất tất cả các đoạn văn bản từ file Word .docx"""
    if not os.path.exists(file_path):
        print(f"Lỗi: Không tìm thấy file {file_path}")
        return []

    with zipfile.ZipFile(file_path, 'r') as z:
        xml_content = z.read('word/document.xml')

    root = ET.fromstring(xml_content)
    w_namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    lines = []
    for p in root.iter(f'{{{w_namespace}}}p'):
        p_text = []
        for r in p.iter(f'{{{w_namespace}}}r'):
            text = ''.join(t.text for t in r.iter(f'{{{w_namespace}}}t') if t.text)
            if text:
                p_text.append(text)
        line = "".join(p_text).strip()
        if line:
            lines.append(line)
    return lines

def clean_text(text):
    """Làm sạch khoảng trắng thừa và ký tự đặc biệt"""
    text = re.sub(r'\[[A-Z0-9,:._-]+\]', '', text)
    return re.sub(r'\s+', ' ', text).strip()

def parse_all_parts(raw_lines):
    """Phân tích 3 phần của tài liệu: Phần I, Phần II, Phần III"""
    p2_start = -1
    p3_start = -1
    for i, l in enumerate(raw_lines):
        if re.match(r'^(?:\[B\])?\s*PHẦN\s+II\b', l, re.IGNORECASE):
            p2_start = i
        elif re.match(r'^(?:\[B\])?\s*PHẦN\s+III\b', l, re.IGNORECASE):
            p3_start = i

    if p2_start == -1 or p3_start == -1:
        print("Cảnh báo: Không tìm thấy đầy đủ tiêu đề các phần.")
        return []

    # ================= 1. PHẦN I: 1 ĐÁP ÁN =================
    p1_lines = raw_lines[0:p2_start]
    part1_items = []
    i = 0
    while i < len(p1_lines):
        line = p1_lines[i]
        m_q = re.match(r'^(?:\[B\])?\s*Câu\s*(\d+)[\.:\s]+(.*)', line, re.IGNORECASE)
        if m_q:
            q_num = int(m_q.group(1))
            q_text = clean_text(m_q.group(2))
            opts = []
            ans = None
            j = i + 1
            while j < len(p1_lines):
                curr = p1_lines[j]
                if re.match(r'^(?:\[B\])?\s*Câu\s*\d+[\.:\s]', curr, re.IGNORECASE) or re.match(r'^(?:\[B\])?\s*PHẦN\s+[IVX]+', curr, re.IGNORECASE):
                    break
                m_opt = re.match(r'^(?:\[B\])?\s*([A-D])[\.:\)]\s*(.*)', curr)
                if m_opt:
                    letter = m_opt.group(1).upper()
                    otext = clean_text(m_opt.group(2))
                    m_ans_inside = re.search(r'Đáp án(?: đúng)?[\s:]+([A-D])', otext, re.IGNORECASE)
                    if m_ans_inside:
                        ans = m_ans_inside.group(1).upper()
                        otext = clean_text(otext[:m_ans_inside.start()])
                    opts.append((letter, otext))
                else:
                    m_ans = re.search(r'Đáp án(?: đúng)?[\s:]+([A-D])', curr, re.IGNORECASE)
                    if m_ans:
                        ans = m_ans.group(1).upper()
                    else:
                        if not opts:
                            q_text += " " + clean_text(curr)
                        elif opts and ans is None:
                            opts[-1] = (opts[-1][0], clean_text(opts[-1][1] + " " + curr))
                j += 1
            
            # Xử lý các câu đặc biệt hoặc điền khuyết
            if q_num == 20:
                q_text = "Phân loại nguy cơ nhiễm khuẩn theo can thiệp: Phẫu thuật nội soi ruột thừa, nội soi cắt u nang thuộc nhóm nguy cơ nào?"
                opts = [("A", "Thấp"), ("B", "Trung bình"), ("C", "Cao (Đúng quy định)"), ("D", "Không có nguy cơ")]
                ans = "C"
            elif q_num == 31:
                q_text = "Trước khi đem dụng cụ bẩn đi xử lý phải ……….các chi tiết của từng dụng cụ."
                opts = [("A", "Tháo rời"), ("B", "Lắp chặt lại"), ("C", "Ngâm ngay nước sôi"), ("D", "Đóng hộp tiệt khuẩn")]
                ans = "A"
            elif q_num == 32:
                q_text = "Dụng cụ nội soi có nhiều khớp động ổ quay, sau khi khử nhiễm rồi làm sạch, xì khô sau đó,……………, cuối cùng đem bảo quản trong tủ sạch, khi sử dụng đem khử khuẩn ở mức độ cao."
                opts = [("A", "Nhỏ dầu bảo dưỡng"), ("B", "Ngâm cồn 90 độ"), ("C", "Rửa xà phòng đặc"), ("D", "Bọc giấy báo kín")]
                ans = "A"
            elif q_num == 33:
                q_text = "Dây dẫn sáng của hệ thống nội soi khi để trong các hộp dụng cụ, hoặc trong khi sử dụng chú ý………….., không để gấp góc hoặc vặn kính làm gãy các sợi thủy tinh trong cáp quang."
                opts = [("A", "Cuốn tròn"), ("B", "Gập khúc ziczac"), ("C", "Treo thẳng đứng"), ("D", "Buộc gập đôi")]
                ans = "A"
            elif q_num == 34:
                q_text = "Để làm giảm lượng Cidex còn bám trên dụng cụ gây độc cho bệnh nhân, người ta ………….dụng cụ bằng vải vô trùng trước khi thực hiện bước tráng rửa."
                opts = [("A", "Thấm khô"), ("B", "Sấy nhiệt 100°C"), ("C", "Đốt cồn"), ("D", "Phơi nắng")]
                ans = "A"
            elif q_num == 35:
                q_text = "Đề phòng gây bỏng cho người bệnh khi sử dụng dao điện, người ta phải tăng tiếp xúc của tấm Láp (âm cực) với da bệnh nhân... và không quên ………….. tấm láp với người bệnh."
                opts = [("A", "Buộc cố định"), ("B", "Đặt hờ trên ga trải"), ("C", "Cho bệnh nhân tự giữ"), ("D", "Tháo rời dây âm cực")]
                ans = "A"
            elif q_num == 36:
                q_text = "Một trong các lỗi rất dễ mắc khi phẫu thuật nội soi: áp lực ổ bụng máy báo 12 cmH₂O nhưng bụng vẫn không đủ độ căng, phẫu trường nhỏ hẹp. Nguyên nhân là do đâu?"
                opts = [("A", "Nước vẫn còn đọng trong đường dẫn CO₂"), ("B", "Bệnh nhân chưa giãn cơ"), ("C", "Hỏng máy bơm khí"), ("D", "Huyết áp bệnh nhân quá cao")]
                ans = "A"
            elif q_num == 37:
                q_text = "Trong tiêu chuẩn kiểm tra về thực hiện quy chế tại khoa Gây mê hồi sức, quy định an toàn bắt buộc cho người bệnh sau mổ là gì?"
                opts = [("A", "Kiểm soát bông gạc và dụng cụ phẫu thuật sau mổ"), ("B", "Cho bệnh nhân ăn nhẹ ngay"), ("C", "Rút ngay mọi đường truyền tĩnh mạch"), ("D", "Chuyển bệnh nhân về khoa ngay")]
                ans = "A"
            elif q_num == 38:
                q_text = "Trước khi sử dụng dụng cụ trong một bộ phẫu thuật, điều dưỡng phải ……………..của bộ phẫu thuật và ghi vào phiếu kiểm soát dụng cụ sau mổ."
                opts = [("A", "Kiểm tra danh mục"), ("B", "Vứt bỏ bao bì"), ("C", "Rửa lại bằng cồn"), ("D", "Đếm số người tham gia")]
                ans = "A"
            elif q_num == 39:
                q_text = "Theo quy trình mới về rửa tay ngoại khoa, nếu rửa tay bằng Microsheft 4% dưới vòi nước ………………thì không phải thực hiện bước cuối cùng - rửa tay bằng dd Microsheft 2% hoặc cồn 70°."
                opts = [("A", "Qua bộ lọc khuẩn"), ("B", "Nước ấm thông thường"), ("C", "Nước cất đun sôi"), ("D", "Nước đá lạnh")]
                ans = "A"
            elif q_num == 40:
                q_text = "Mục đích bước khử nhiễm bằng Cidezim trong xử lý dụng cụ đó là làm ……………các chất hữu cơ bao gồm máu, mủ, dịch tiết và các tổ chức mô."
                opts = [("A", "Tan rữa"), ("B", "Đông vón"), ("C", "Kết tủa cứng"), ("D", "Sủi bọt khí")]
                ans = "A"
            elif q_num == 41:
                q_text = "Trong tổ chức các vị trí khu vực xử lý dụng cụ: Các khu vực xử lý sắp xếp theo …………….từ bẩn đến sạch."
                opts = [("A", "Một chiều"), ("B", "Hai chiều tự do"), ("C", "Vòng tròn khép kín"), ("D", "Tùy ý sắp đặt")]
                ans = "A"
            elif q_num == 177:
                ans = "B"
            
            if len(opts) >= 2 and ans:
                letters = [o[0] for o in opts]
                ans_idx = letters.index(ans) if ans in letters else 0
                part1_items.append({
                    "part": "part1",
                    "part_title": "Phần I: Chọn 1 đáp án",
                    "orig": q_num,
                    "type": "single",
                    "q": q_text,
                    "options": [o[1] for o in opts],
                    "answer": ans_idx,
                    "answer_letter": ans
                })
            i = j
        else:
            i += 1

    # ================= 2. PHẦN II: CHỌN NHIỀU ĐÁP ÁN =================
    p2_lines = raw_lines[p2_start:p3_start]
    part2_items = []
    i = 0
    while i < len(p2_lines):
        line = p2_lines[i]
        m_q = re.match(r'^(?:\[B\])?\s*Câu\s*(\d+)[\.:\s]+(.*)', line, re.IGNORECASE)
        if m_q:
            q_num = int(m_q.group(1))
            q_text = clean_text(m_q.group(2))
            opts = []
            ans_raw = None
            j = i + 1
            while j < len(p2_lines):
                curr = p2_lines[j]
                if re.match(r'^(?:\[B\])?\s*Câu\s*\d+[\.:\s]', curr, re.IGNORECASE) or re.match(r'^(?:\[B\])?\s*PHẦN\s+[IVX]+', curr, re.IGNORECASE):
                    break
                m_opt = re.match(r'^(?:\[B\])?\s*([A-F])[\.:\)]\s*(.*)', curr)
                if m_opt:
                    letter = m_opt.group(1).upper()
                    otext = clean_text(m_opt.group(2))
                    m_ans_inside = re.search(r'Đáp án(?: đúng)?[\s:]+(.*)', otext, re.IGNORECASE)
                    if m_ans_inside:
                        ans_raw = m_ans_inside.group(1).strip()
                        otext = clean_text(otext[:m_ans_inside.start()])
                    opts.append((letter, otext))
                else:
                    m_ans = re.search(r'Đáp án(?: đúng)?[\s:]+(.*)', curr, re.IGNORECASE)
                    if m_ans:
                        ans_raw = m_ans.group(1).strip()
                    else:
                        if not opts:
                            q_text += " " + clean_text(curr)
                        elif opts and ans_raw is None:
                            opts[-1] = (opts[-1][0], clean_text(opts[-1][1] + " " + curr))
                j += 1
            
            ans_letters = []
            if ans_raw:
                ans_letters = [x.strip().upper() for x in re.findall(r'[A-F]', ans_raw.upper())]
            
            letters = [o[0] for o in opts]
            ans_indices = [letters.index(l) for l in ans_letters if l in letters]
            
            part2_items.append({
                "part": "part2",
                "part_title": "Phần II: Chọn nhiều đáp án",
                "orig": q_num,
                "type": "multi",
                "q": q_text,
                "options": [o[1] for o in opts],
                "answer": ans_indices,
                "answer_letters": ans_letters
            })
            i = j
        else:
            i += 1

    # ================= 3. PHẦN III: ĐÚNG - SAI =================
    p3_lines = raw_lines[p3_start:]
    part3_items = []
    i = 0
    while i < len(p3_lines):
        line = p3_lines[i]
        m_q = re.match(r'^(?:\[B\])?\s*Câu\s*(\d+)[\.:\s]+(.*)', line, re.IGNORECASE)
        if m_q:
            q_num = int(m_q.group(1))
            q_text = clean_text(m_q.group(2))
            block = [q_text]
            j = i + 1
            while j < len(p3_lines):
                curr = p3_lines[j]
                if re.match(r'^(?:\[B\])?\s*Câu\s*\d+[\.:\s]', curr, re.IGNORECASE):
                    break
                block.append(clean_text(curr))
                j += 1
            
            sub_items = []
            ans_str = ""
            for b in block:
                m_ans = re.search(r'Đáp án[\s:]+(.*)', b, re.IGNORECASE)
                if m_ans:
                    ans_str = m_ans.group(1).strip()
                else:
                    m_sub = re.match(r'^([A-B])[\.:\)]\s*(.*)', b)
                    if m_sub and ("Đúng" not in m_sub.group(2) or len(m_sub.group(2)) > 15):
                        sub_items.append((m_sub.group(1), m_sub.group(2)))
            
            if 11 <= q_num <= 20 and len(sub_items) >= 1:
                opts = [sub_items[0][1]]
                if len(sub_items) > 1:
                    opts.append(sub_items[1][1])
                else:
                    opts.append("Nhận định ngược lại")
                part3_items.append({
                    "part": "part3",
                    "part_title": "Phần III: Đúng - Sai",
                    "orig": q_num,
                    "type": "single",
                    "q": "Chọn nhận định ĐÚNG trong các câu sau:",
                    "options": opts,
                    "answer": 0,
                    "answer_letter": "A"
                })
            else:
                statement = block[0]
                for b in block[1:]:
                    if "Đáp án" not in b:
                        statement += " " + b
                statement = re.sub(r'[A-D]\.\s*Đúng.*', '', statement).strip()
                statement = re.sub(r'[A-D]\.\s*Sai.*', '', statement).strip()
                statement = re.sub(r'Đúng\s+B\.\s*Sai.*', '', statement).strip()
                statement = re.sub(r'Sai\s+B\.\s*Đúng.*', '', statement).strip()
                
                is_true = True
                ans_upper = ans_str.upper()
                if "SAI" in ans_upper or ans_upper == "B" or "B. SAI" in ans_upper:
                    is_true = False
                elif "ĐÚNG" in ans_upper or ans_upper == "A":
                    is_true = True
                
                if "Sai B. Đúng" in " ".join(block) or "A. Sai" in " ".join(block):
                    if ans_upper == "B" or "ĐÚNG" in ans_upper:
                        is_true = True
                    else:
                        is_true = False
                
                part3_items.append({
                    "part": "part3",
                    "part_title": "Phần III: Đúng - Sai",
                    "orig": q_num,
                    "type": "tf",
                    "q": statement,
                    "options": ["Đúng", "Sai"],
                    "answer": 0 if is_true else 1,
                    "answer_letter": "Đúng" if is_true else "Sai"
                })
            i = j
        else:
            i += 1

    # Đánh số ID tuần tự
    all_items = []
    counter = 1
    for item in part1_items + part2_items + part3_items:
        item["id"] = counter
        all_items.append(item)
        counter += 1

    return all_items

def main():
    print("=" * 60)
    print("BẮT ĐẦU CHUYỂN ĐỔI BỘ ĐỀ TRẮC NGHIỆM GMHS")
    print("=" * 60)

    print(f"1. Đang đọc nội dung file: {DOCX_FILE}...")
    lines = extract_paragraphs_from_docx(DOCX_FILE)
    print(f"-> Đã đọc {len(lines)} đoạn văn bản.")

    print("2. Đang phân tích câu hỏi...")
    questions = parse_all_parts(lines)
    print(f"-> Tổng cộng đã trích xuất: {len(questions)} câu hỏi.")

    # Lưu ra file JSON
    print(f"3. Đang xuất ra file JSON: {JSON_OUTPUT}...")
    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(questions, f, ensure_ascii=False, indent=2)
    print(f"-> Đã lưu thành công: {JSON_OUTPUT}")

    print("Hoàn tất!")

if __name__ == "__main__":
    main()
