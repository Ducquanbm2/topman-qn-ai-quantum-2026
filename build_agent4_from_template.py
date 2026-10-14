import zipfile, copy
from pathlib import Path
from xml.etree import ElementTree as ET

SRC=Path('Template.pptx'); OUT=Path('Agent4_CFO_ao_template_moi.pptx')
P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main'; R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ET.register_namespace('a',A); ET.register_namespace('p',P); ET.register_namespace('r',R)
def E(ns,tag,attrs=None): return ET.Element('{%s}%s'%(ns,tag),attrs or {})
def text_box(root,x,y,w,h,text,size=2200,color='FFFFFF',bold=False,align='l'):
    tree=root.find('.//{%s}spTree'%P); sp=E(P,'sp'); nv=E(P,'nvSpPr'); nv.append(E(P,'cNvPr',{'id':str(10000+x+y),'name':'content'})); nv.append(E(P,'cNvSpPr')); nv.append(E(P,'nvPr')); sp.append(nv)
    pr=E(P,'spPr'); xf=E(A,'xfrm'); xf.append(E(A,'off',{'x':str(x),'y':str(y)})); xf.append(E(A,'ext',{'cx':str(w),'cy':str(h)})); pr.append(xf); geom=E(A,'prstGeom',{'prst':'rect'}); geom.append(E(A,'avLst')); pr.append(geom); sp.append(pr)
    tx=E(P,'txBody'); tx.append(E(A,'bodyPr',{'wrap':'square','anchor':'ctr'})); tx.append(E(A,'lstStyle')); p=E(A,'p'); p.append(E(A,'pPr',{'algn':align})); rr=E(A,'r'); rr.append(E(A,'rPr',{'lang':'vi-VN','sz':str(size),'b':'1' if bold else '0'})); tt=E(A,'t'); tt.text=text; rr.append(tt); p.append(rr); p.append(E(A,'endParaRPr',{'lang':'vi-VN','sz':str(size)})); tx.append(p); sp.append(tx); tree.append(sp)
def box(root,x,y,w,h,title,body,fill):
    tree=root.find('.//{%s}spTree'%P); sp=E(P,'sp'); nv=E(P,'nvSpPr'); nv.append(E(P,'cNvPr',{'id':str(11000+x+y),'name':title})); nv.append(E(P,'cNvSpPr')); nv.append(E(P,'nvPr')); sp.append(nv); pr=E(P,'spPr'); xf=E(A,'xfrm'); xf.append(E(A,'off',{'x':str(x),'y':str(y)})); xf.append(E(A,'ext',{'cx':str(w),'cy':str(h)})); pr.append(xf); g=E(A,'prstGeom',{'prst':'roundRect'}); g.append(E(A,'avLst')); pr.append(g); sf=E(A,'solidFill'); sf.append(E(A,'srgbClr',{'val':fill})); pr.append(sf); sp.append(pr)
    tx=E(P,'txBody'); tx.append(E(A,'bodyPr',{'wrap':'square','anchor':'ctr'})); tx.append(E(A,'lstStyle'))
    for s,sz,b in [(title,1500,True),(body,1050,False)]:
        p=E(A,'p'); p.append(E(A,'pPr',{'algn':'ctr'})); r=E(A,'r'); r.append(E(A,'rPr',{'lang':'vi-VN','sz':str(sz),'b':'1' if b else '0'})); tt=E(A,'t'); tt.text=s; r.append(tt); p.append(r); p.append(E(A,'endParaRPr',{'lang':'vi-VN','sz':str(sz)})); tx.append(p)
    sp.append(tx); tree.append(sp)
def arrow(root,x,y,w):
    tree=root.find('.//{%s}spTree'%P); sp=E(P,'cxnSp'); nv=E(P,'nvCxnSpPr'); nv.append(E(P,'cNvPr',{'id':str(12000+x+y),'name':'arrow'})); nv.append(E(P,'cNvCxnSpPr')); nv.append(E(P,'nvPr')); sp.append(nv); pr=E(P,'spPr'); xf=E(A,'xfrm'); xf.append(E(A,'off',{'x':str(x),'y':str(y)})); xf.append(E(A,'ext',{'cx':str(w),'cy':'0'})); pr.append(xf); g=E(A,'prstGeom',{'prst':'line'}); g.append(E(A,'avLst')); pr.append(g); ln=E(A,'ln',{'w':'19050'}); sf=E(A,'solidFill'); sf.append(E(A,'srgbClr',{'val':'6FA8DC'})); ln.append(sf); ln.append(E(A,'headEnd',{'type':'triangle'})); pr.append(ln); sp.append(pr); tree.append(sp)
def slide(title, page, kind):
    root=ET.fromstring(base_slide)
    text_box(root,500000,250000,17000000,700000,title,2800,'FFFFFF',True)
    text_box(root,500000,950000,17000000,350000,f'AGENT 4 • TRỢ LÝ ĐIỀU HÀNH (CFO ẢO)  |  {page}',900,'B7C9E2')
    if kind=='problem':
        items=[('01','Chờ báo cáo','Muốn biết lãi, tiền, công nợ phải chờ kế toán tổng hợp'),('02','Dữ liệu rời rạc','Mỗi phần mềm/Excel một cách lưu, dễ lệch số'),('03','Câu hỏi mơ hồ','“Lãi” có thể là lãi gộp, lãi ròng hoặc tiền còn lại'),('04','Thiếu niềm tin','Không biết con số lấy từ đâu, tính đến ngày nào')]
        for i,(n,t,b) in enumerate(items): box(root,650000+(i%2)*8300000,1800000+(i//2)*2200000,7200000,1500000,n+'  '+t,b,['2F75B5','70AD47','ED7D31','A64D79'][i])
    elif kind=='data':
        xs=[600000,4700000,8800000,12900000]; vals=[('Nguồn dữ liệu','Sổ cái\nDoanh thu • chi phí\nTiền • công nợ','4472C4'),('Chuẩn hóa','Tài khoản TT133\nĐối tác • thời gian','70AD47'),('Kho dữ liệu','Một nguồn dữ liệu\ncho mọi trợ lý','ED7D31'),('Kết quả','Metrics chuẩn\nđể hỏi & xem biểu đồ','5B9BD5')]
        for x,v in zip(xs,vals): box(root,x,2600000,3200000,1700000,*v)
        for x in xs[:-1]: arrow(root,x+3300000,3450000,700000)
        text_box(root,1300000,5200000,15000000,900000,'Hình dung đơn giản: gom dữ liệu kế toán về một “sổ cái chuẩn”, rồi mọi câu trả lời đều lấy từ đó.',1500,'FFFFFF',False,'c')
    elif kind=='architecture':
        xs=[350000,3650000,6950000,10250000,13550000]; vals=[('Người dùng','Hỏi bằng chữ\nhoặc giọng nói','4472C4'),('Hiểu câu hỏi','LLM nhận diện\nmetric + thời gian','70AD47'),('Chọn đường đi','Router chọn hàm\nhoặc truy vấn dự phòng','ED7D31'),('Tính toán','Metric library +\nDatabase tài chính','A64D79'),('Trả lời','Số liệu + biểu đồ\n+ “Xem cách tính”','5B9BD5')]
        for x,v in zip(xs,vals): box(root,x,2500000,2800000,1800000,*v)
        for x in xs[:-1]: arrow(root,x+2900000,3400000,400000)
        text_box(root,1500000,5200000,14500000,800000,'Đường chính: gọi các chỉ số đã kiểm chứng.  Đường phụ: Text-to-SQL chỉ đọc, giới hạn schema và phải kiểm tra lại.',1400,'FFFFFF',False,'c')
    elif kind=='method':
        vals=[('LLM','Hiểu ý định\n& diễn giải','4472C4'),('Router','Chọn đúng\nmetric / bảng','70AD47'),('Function Calling','Gọi hàm đã\nđịnh nghĩa sẵn','ED7D31'),('SQL / DuckDB','Tính toán\nđịnh lượng','A64D79'),('Verifier','Kiểm tra số\n& nguồn trích dẫn','5B9BD5')]; xs=[500000,3900000,7300000,10700000,14100000]
        for x,v in zip(xs,vals): box(root,x,2300000,2600000,1700000,*v)
        for x in xs[:-1]: arrow(root,x+2700000,3150000,650000)
        text_box(root,1300000,5000000,15000000,1000000,'Cách nói dễ hiểu: AI hiểu câu hỏi, nhưng không tự bịa phép tính. Phần tính số chạy bằng công thức và dữ liệu đã kiểm chứng.',1500,'FFFFFF',False,'c')
    elif kind=='model':
        box(root,850000,2200000,4800000,1800000,'Model ngôn ngữ','LLM API (ví dụ Gemini)\nhoặc model self-host Qwen-class','4472C4'); box(root,7000000,2200000,4800000,1800000,'Model tính toán','Metric SQL + DuckDB\nnhanh, rẻ, đúng','70AD47'); box(root,13150000,2200000,4000000,1800000,'Kiểm chứng','Rule + verifier\nchặn câu trả lời sai','ED7D31'); arrow(root,5650000,3100000,1200000); arrow(root,11800000,3100000,1200000); text_box(root,1700000,5000000,14500000,900000,'Model ngôn ngữ lo phần giao tiếp. Model / engine dữ liệu lo phần con số. Hai vai trò tách biệt để an toàn.',1500,'FFFFFF',False,'c')
    elif kind=='experience':
        vals=[('Câu hỏi','“Tháng này lãi\nbao nhiêu?”','4472C4'),('Câu trả lời','Lợi nhuận sau thuế\n= … VNĐ','70AD47'),('Biểu đồ','Xu hướng theo tháng\nhoặc theo chi nhánh','ED7D31'),('Truy vết','Nguồn dữ liệu\n+ thời điểm chốt số','5B9BD5')]; xs=[650000,4800000,8950000,13100000]
        for x,v in zip(xs,vals): box(root,x,2600000,3200000,1700000,*v)
        for x in xs[:-1]: arrow(root,x+3300000,3450000,700000)
        text_box(root,1350000,5200000,14500000,850000,'Nếu câu hỏi mơ hồ hoặc dữ liệu chưa đủ: hệ thống hỏi lại / cảnh báo, không đoán bừa.',1500,'FFFFFF',False,'c')
    elif kind=='summary':
        box(root,1000000,2200000,5000000,1900000,'Giá trị chính','Trả lời nhanh các câu hỏi\nđiều hành hằng ngày','4472C4'); box(root,7000000,2200000,5000000,1900000,'Điểm khác biệt','Nhanh + đúng + truy vết\nkhông phụ thuộc Text-to-SQL tự do','70AD47'); box(root,13000000,2200000,4000000,1900000,'Lộ trình','Metrics → hỏi đáp →\nbản tin chủ động','ED7D31')
    key = {
      'problem':'Ý chính: Đây không chỉ là chatbot. Bài toán là biến dữ liệu kế toán thành thông tin điều hành nhanh và đáng tin cậy.',
      'data':'Ý chính: Một nguồn dữ liệu chuẩn giúp các trợ lý không trả về những con số khác nhau cho cùng một câu hỏi.',
      'architecture':'Ý chính: Ưu tiên đường đi an toàn và có kiểm soát; chỉ dùng đường dự phòng khi câu hỏi nằm ngoài thư viện chỉ số.',
      'method':'Ý chính: LLM phụ trách giao tiếp; metric library và database phụ trách sự chính xác.',
      'model':'Ý chính: Không cần một model làm tất cả. Mỗi thành phần làm đúng phần việc để tối ưu chi phí và giảm rủi ro.',
      'experience':'Ý chính: Người dùng nhận được câu trả lời có thể kiểm tra, không chỉ một con số “trông có vẻ đúng”.',
      'summary':'Thông điệp kết luận: CFO ảo là lớp điều hành ở trên cùng — giúp lãnh đạo hỏi, hiểu và hành động dựa trên cùng một nguồn dữ liệu tài chính đáng tin cậy.'
    }
    if kind in key:
        text_box(root,900000,6500000,16000000,800000,key[kind],1450,'FFD966',True,'c')
    return root

titles=[('AGENT 4 — TRỢ LÝ ĐIỀU HÀNH (CFO ẢO)','1','cover'),('Vấn đề thực tế','2','problem'),('Dữ liệu đến từ đâu?','3','data'),('Tổng quan kiến trúc','4','architecture'),('Phương pháp sử dụng','5','method'),('Model sử dụng','6','model'),('Trải nghiệm người dùng','7','experience'),('Tóm tắt','8','summary')]
with zipfile.ZipFile(SRC) as zin:
    base_slide=zin.read('ppt/slides/slide2.xml')
    files={i.filename:zin.read(i.filename) for i in zin.infolist()}
    pres=ET.fromstring(files['ppt/presentation.xml']); sld=pres.find('.//{%s}sldIdLst'%P)
    # Use a blank template slide for every content slide; title slide is also rebuilt for consistency.
    for i,(title,page,kind) in enumerate(titles,1):
        root=slide(title,page,kind)
        files[f'ppt/slides/slide{i}.xml']=ET.tostring(root,encoding='utf-8',xml_declaration=True)
        if i >= 3:
            files[f'ppt/slides/_rels/slide{i}.xml.rels'] = files['ppt/slides/_rels/slide2.xml.rels']
    # Remove any old slide ids and add eight new ones.
    for child in list(sld): sld.remove(child)
    for i in range(1,9):
        rid=f'rId{20+i}'; sld.append(E(P,'sldId',{'id':str(255+i),'r:id':rid}))
        files[f'ppt/_rels/presentation.xml.rels']=files['ppt/_rels/presentation.xml.rels']
    relroot=ET.fromstring(files['ppt/_rels/presentation.xml.rels']);
    for i in range(1,9): relroot.append(E('http://schemas.openxmlformats.org/package/2006/relationships','Relationship',{'Id':f'rId{20+i}','Target':f'slides/slide{i}.xml','Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'}))
    files['ppt/_rels/presentation.xml.rels']=ET.tostring(relroot,encoding='utf-8',xml_declaration=True)
    files['ppt/presentation.xml']=ET.tostring(pres,encoding='utf-8',xml_declaration=True)
    # content types: add slide overrides
    ct=ET.fromstring(files['[Content_Types].xml']); nsct='http://schemas.openxmlformats.org/package/2006/content-types'
    for i in range(3,9): ct.append(E(nsct,'Override',{'PartName':f'/ppt/slides/slide{i}.xml','ContentType':'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'}))
    files['[Content_Types].xml']=ET.tostring(ct,encoding='utf-8',xml_declaration=True)
    with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as zout:
        for name,data in files.items(): zout.writestr(name,data)
print('Created',OUT)
