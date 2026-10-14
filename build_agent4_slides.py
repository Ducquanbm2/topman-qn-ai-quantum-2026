import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

SRC = Path('Bản sao của Tiêu đề (5).pptx.pptx')
OUT = Path('Agent4_CFO_ao_tong_quan.pptx')

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ET.register_namespace('a', A); ET.register_namespace('p', P); ET.register_namespace('r', R)

def tnodes(root):
    return root.findall('.//{%s}t' % A)

def replace_all(root, mapping):
    for t in tnodes(root):
        if t.text in mapping:
            t.text = mapping[t.text]

def add_textbox(root, x, y, w, h, lines, size=1800, color='FFFFFF'):
    spTree = root.find('.//{%s}spTree' % P)
    sp = ET.Element('{%s}sp' % P)
    nv = ET.SubElement(sp, '{%s}nvSpPr' % P)
    ET.SubElement(nv, '{%s}cNvPr' % P, {'id':'9001','name':'Agent4 content'})
    ET.SubElement(nv, '{%s}cNvSpPr' % P)
    ET.SubElement(nv, '{%s}nvPr' % P)
    spPr = ET.SubElement(sp, '{%s}spPr' % P)
    xfrm = ET.SubElement(spPr, '{%s}xfrm' % A)
    ET.SubElement(xfrm, '{%s}off' % A, {'x':str(x),'y':str(y)})
    ET.SubElement(xfrm, '{%s}ext' % A, {'cx':str(w),'cy':str(h)})
    tx = ET.SubElement(sp, '{%s}txBody' % P)
    ET.SubElement(tx, '{%s}bodyPr' % A, {'wrap':'square','anchor':'t'})
    ET.SubElement(tx, '{%s}lstStyle' % A)
    for line in lines:
        p = ET.SubElement(tx, '{%s}p' % A)
        pPr = ET.SubElement(p, '{%s}pPr' % A, {'lvl':'0'})
        ET.SubElement(pPr, '{%s}buChar' % A, {'char':'•'})
        r = ET.SubElement(p, '{%s}r' % A)
        rPr = ET.SubElement(r, '{%s}rPr' % A, {'lang':'vi-VN','sz':str(size),'b':'0'})
        solid = ET.SubElement(rPr, '{%s}solidFill' % A)
        ET.SubElement(solid, '{%s}srgbClr' % A, {'val':color})
        ET.SubElement(r, '{%s}t' % A).text = line
        ET.SubElement(p, '{%s}endParaRPr' % A, {'lang':'vi-VN','sz':str(size)})
    spTree.append(sp)

def add_box(root, x, y, w, h, title, body, fill='1F4E79'):
    spTree = root.find('.//{%s}spTree' % P)
    sp = ET.Element('{%s}sp' % P)
    nv = ET.SubElement(sp, '{%s}nvSpPr' % P)
    ET.SubElement(nv, '{%s}cNvPr' % P, {'id':str(9100+x+y), 'name':title})
    ET.SubElement(nv, '{%s}cNvSpPr' % P)
    ET.SubElement(nv, '{%s}nvPr' % P)
    spPr = ET.SubElement(sp, '{%s}spPr' % P)
    xfrm = ET.SubElement(spPr, '{%s}xfrm' % A)
    ET.SubElement(xfrm, '{%s}off' % A, {'x':str(x),'y':str(y)})
    ET.SubElement(xfrm, '{%s}ext' % A, {'cx':str(w),'cy':str(h)})
    prst = ET.SubElement(spPr, '{%s}prstGeom' % A, {'prst':'roundRect'})
    ET.SubElement(prst, '{%s}avLst' % A)
    fillEl = ET.SubElement(spPr, '{%s}solidFill' % A)
    ET.SubElement(fillEl, '{%s}srgbClr' % A, {'val':fill})
    ln = ET.SubElement(spPr, '{%s}ln' % A)
    ET.SubElement(ln, '{%s}solidFill' % A).append(ET.Element('{%s}srgbClr' % A, {'val':'FFFFFF'}))
    tx = ET.SubElement(sp, '{%s}txBody' % P)
    ET.SubElement(tx, '{%s}bodyPr' % A, {'wrap':'square','anchor':'ctr'})
    ET.SubElement(tx, '{%s}lstStyle' % A)
    for text, bold, sz in [(title, True, 1700), (body, False, 1250)]:
        p = ET.SubElement(tx, '{%s}p' % A)
        pPr = ET.SubElement(p, '{%s}pPr' % A, {'algn':'ctr'})
        r = ET.SubElement(p, '{%s}r' % A)
        ET.SubElement(r, '{%s}rPr' % A, {'lang':'vi-VN','sz':str(sz),'b':'1' if bold else '0'} )
        ET.SubElement(r, '{%s}t' % A).text = text
        ET.SubElement(p, '{%s}endParaRPr' % A, {'lang':'vi-VN','sz':str(sz)})
    spTree.append(sp)

def add_arrow(root, x1, y1, x2, y2):
    spTree = root.find('.//{%s}spTree' % P)
    cx, cy = x2-x1, y2-y1
    sp = ET.Element('{%s}cxnSp' % P)
    nv = ET.SubElement(sp, '{%s}nvCxnSpPr' % P)
    ET.SubElement(nv, '{%s}cNvPr' % P, {'id':str(9200+x1+y1), 'name':'flow arrow'})
    ET.SubElement(nv, '{%s}cNvCxnSpPr' % P)
    ET.SubElement(nv, '{%s}nvPr' % P)
    spPr = ET.SubElement(sp, '{%s}spPr' % P)
    xfrm = ET.SubElement(spPr, '{%s}xfrm' % A)
    ET.SubElement(xfrm, '{%s}off' % A, {'x':str(x1),'y':str(y1)})
    ET.SubElement(xfrm, '{%s}ext' % A, {'cx':str(cx),'cy':str(cy)})
    geom = ET.SubElement(spPr, '{%s}prstGeom' % A, {'prst':'line'}); ET.SubElement(geom, '{%s}avLst' % A)
    ln = ET.SubElement(spPr, '{%s}ln' % A, {'w':'19050'})
    ET.SubElement(ln, '{%s}solidFill' % A).append(ET.Element('{%s}srgbClr' % A, {'val':'5B9BD5'}))
    ET.SubElement(ln, '{%s}tailEnd' % A, {'type':'none'})
    ET.SubElement(ln, '{%s}headEnd' % A, {'type':'triangle'})
    spTree.append(sp)

repls = {
1: {'AGENT 3 — TRỢ LÝ PHÁP LÝ VÀ THUẾ':'AGENT 4 — TRỢ LÝ ĐIỀU HÀNH (CFO ẢO)',
    'Trang 1':'Trang 1','Tự động nhận diện, trích xuất và  kiểm tra hóa đơn theo luật thuế VN':'Hỏi đáp tài chính tức thời cho Ban Giám đốc', '✅❌🟡':'💬  🔎  ✅'},
2: {'Vấn đề đang giải quyết':'Vấn đề đang giải quyết',
    'Doanh nghiệp nhận chứng từ từ nhiều nguồn: file XML, ảnh scan, Excel, sao kê ngân hàng | Kế toán phải dò thủ công từng tờ → tốn thời gian, dễ sót lỗi | Sai thuế suất, thiếu chứng từ thanh toán → bị truy thu và phạt khi thanh tra | Luật thuế thay đổi liên tục → khó cập nhật kịp':'Lãnh đạo thường phải chờ kế toán tổng hợp số liệu mới biết tình hình.\nDữ liệu nằm ở sổ cái, Excel và nhiều báo cáo khác nhau.\nCâu hỏi đơn giản như “Tháng này lãi bao nhiêu?” dễ mất nhiều thời gian để trả lời.\nKhó truy ngược con số được tính từ chứng từ nào.',
    'ví dụ:':'Ví dụ:', '⚠️ 1 kế toán chỉ dò được 100–150 hóa đơn/ngày Cuối tháng dồn 5.000+ tờ → quá tải, sai sót là tất yếu':'⚠️ Chậm thông tin → quyết định chậm; số liệu thủ công → dễ nhầm hoặc thiếu nhất quán',
    '💰 Chỉ 1 hóa đơn sai quy định = mất trắng 20% tiền thuế khấu trừ + phạt gấp đôi':'💡 Agent 4 biến câu hỏi tự nhiên thành câu trả lời có số liệu và nguồn kiểm chứng'},
3: {' Tổng quan giải pháp':'Tổng quan kiến trúc Agent 4'},
4: {'Dữ liệu (Data)':'Dữ liệu (Data)', 'Đầu vào: Hóa đơn tiếng Việt dạng PDF/ảnh scan (giai đoạn đầu) | Nguồn data huấn luyện & kiểm thử: | Hóa đơn GTGT mẫu (theo form chuẩn Thông tư 78/2021) | Bộ dataset tự tạo: scan/chụp ảnh hóa đơn thật, gán nhãn các trường (MST, tên, ngày, tiền, thuế suất) | Luật thuế: Luật 48/2024, NĐ 181/2025, NQ 204/2025 (chính sách giảm VAT 2%) | Đầu ra: File JSON chuẩn hoá chứa toàn bộ thông tin hóa đơn + kết quả kiểm tra đúng/sai':'Đầu vào: sổ cái kế toán theo TT133 (gl_entry) và các bảng danh mục.\nNguồn dữ liệu: doanh thu, chi phí, tiền mặt/ngân hàng, công nợ khách hàng & nhà cung cấp, tồn kho.\nDữ liệu được chuẩn hoá theo thời gian, tài khoản, đối tác và sản phẩm/dịch vụ.\nĐầu ra: các chỉ số tài chính đã kiểm định + câu trả lời tiếng Việt dễ hiểu.', 'Ví dụ trên dành chi input đầu vào là pdf/png':'Ví dụ: “Công nợ khách hàng nào đang cao nhất?”'},
5: {'Phương pháp & Model sử dụng':'Phương pháp & Model sử dụng'},
6: {'Các trường trích xuất từ hóa đơn':'Các thành phần chính của câu trả lời'},
7: {'Kết quả kiểm tra':'Kết quả trả lời', 'Hệ thống trả về 3 loại kết quả cho mỗi chứng từ':'Mỗi câu trả lời gồm 3 phần: con số chính xác, cách hiểu ngắn gọn, nguồn dữ liệu/thời điểm chốt số', 'Kèm theo giải thích bằng tiếng Việt, trích dẫn đúng điều luật.':'Nếu câu hỏi mơ hồ, hệ thống hỏi lại hoặc nêu rõ các cách hiểu. Người dùng vẫn có quyền kiểm tra và quyết định.'}
}

extra = {
3: ['Người dùng hỏi bằng ngôn ngữ tự nhiên: “Tháng này lãi bao nhiêu?”', 'LLM hiểu ý định → chọn Metric ID và tham số thời gian', 'Semantic Layer / Function Calling gọi công thức SQL đã kiểm thử', 'Database trả số liệu → LLM diễn giải thành câu trả lời dễ hiểu', 'Mỗi câu trả lời có nguồn, thời gian chốt và dấu vết kiểm toán'],
5: ['LLM: dùng để hiểu câu hỏi và diễn đạt câu trả lời; không tự tính số', 'Semantic Layer: thư viện 15–20 metrics chuẩn như doanh thu, lãi, tiền, công nợ', 'Function Calling: gọi đúng hàm đã định nghĩa sẵn', 'Model đề xuất: LLM Qwen-class tự host cho dữ liệu nhạy cảm; SQL/DuckDB cho tính toán', 'Nguyên tắc: “LLM diễn giải — hệ thống tính toán”'],
6: ['Câu hỏi / Intent: người dùng đang muốn biết điều gì?', 'Metric ID: net_profit, cash_and_equivalents, ar_balance…', 'Khoảng thời gian: tháng, quý, năm hoặc đến một ngày cụ thể', 'Chiều phân tích: theo khách hàng, nhà cung cấp, tài khoản, sản phẩm', 'Kết quả: số liệu + giải thích + nguồn truy vết'],
}

with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.startswith('ppt/slides/slide') and item.filename.endswith('.xml') and item.filename.split('/')[-1][5:-4].isdigit():
            num = int(item.filename.split('/')[-1][5:-4])
            root = ET.fromstring(data)
            replace_all(root, repls.get(num, {}))
            if num in extra:
                if num != 3:
                    add_textbox(root, 850000, 1900000, 16000000, 4300000, extra[num], 1700)
            if num == 3:
                # Sơ đồ kiến trúc chính: đọc từ trái sang phải, ít thuật ngữ, dễ thuyết trình.
                y, w, h = 2500000, 2600000, 1350000
                xs = [450000, 3600000, 6750000, 9900000, 13050000]
                boxes = [
                    ('1. Người dùng', 'Hỏi bằng tiếng Việt', '4472C4'),
                    ('2. LLM', 'Hiểu ý định\nthời gian', '70AD47'),
                    ('3. Semantic Layer', 'Chọn metric\ncông thức chuẩn', 'ED7D31'),
                    ('4. Database', 'Tính số chính xác\ntừ sổ cái', 'A5A5A5'),
                    ('5. CFO ảo trả lời', 'Diễn giải + nguồn\nthời điểm chốt số', '5B9BD5'),
                ]
                for x, (title, body, fill) in zip(xs, boxes): add_box(root, x, y, w, h, title, body, fill)
                for x in xs[:-1]: add_arrow(root, x+w+50000, y+h//2, x+3000000, y+h//2)
                add_textbox(root, 1150000, 4600000, 14500000, 900000, ['Nguyên tắc an toàn: LLM không tự bịa hoặc tự tính số — chỉ diễn giải kết quả từ công thức đã kiểm thử.'], 1450)
            data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        zout.writestr(item, data)
print(f'Created {OUT}')
