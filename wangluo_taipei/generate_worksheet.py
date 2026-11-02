import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'<w:{edge} {nsdecls("w")} w:val="{edge_data.get("val", "single")}" w:sz="{edge_data.get("sz", 4)}" w:space="0" w:color="{edge_data.get("color", "CCCCCC")}"/>'
            tcBorders.append(parse_xml(tag))
        else:
            tag = f'<w:{edge} {nsdecls("w")} w:val="none"/>'
            tcBorders.append(parse_xml(tag))
    tcPr.append(tcBorders)

def set_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def create_worksheet():
    doc = docx.Document()
    
    # 邊界設定 (約 1.7 cm)
    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)

    # 樣式基礎字型
    normal_style = doc.styles['Normal']
    normal_style.font.name = '微軟正黑體'
    normal_style.font.size = Pt(9.5)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')

    # ================= 1. 表頭大標題 =================
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    
    run_badge = title_p.add_run("【學思軒・八上閱思龍門】數位媒體識讀專題\n")
    run_badge.font.size = Pt(9.5)
    run_badge.font.bold = True
    run_badge.font.color.rgb = RGBColor(29, 78, 216) # 藍色系

    run_title = title_p.add_run("《網羅臺北——數位文本品質與正確性評鑑》簡要學習單")
    run_title.font.size = Pt(15.5)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(15, 23, 42)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_after = Pt(6)
    run_sub = sub_p.add_run("出處光譜掃描 × CRAAP+ 五星評鑑雷達 × 大稻埕案例擂台 × 實作探究任務")
    run_sub.font.size = Pt(8.5)
    run_sub.font.color.rgb = RGBColor(100, 116, 139)

    # 基本資料欄 Table
    info_table = doc.add_table(rows=1, cols=5)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.autofit = False

    col_widths = [Inches(1.5), Inches(1.2), Inches(1.8), Inches(1.4), Inches(1.1)]
    row = info_table.rows[0]
    set_row_cant_split(row)
    
    headers = ["班級：__________", "座號：_____", "姓名：____________", "教師評定：", "得分："]
    for i, cell in enumerate(row.cells):
        cell.width = col_widths[i]
        set_cell_background(cell, "EFF6FF") # 藍色淺底
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        set_cell_border(cell, top={"sz":4,"color":"93C5FD"}, bottom={"sz":4,"color":"93C5FD"},
                              left={"sz":4,"color":"93C5FD"}, right={"sz":4,"color":"93C5FD"})
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(headers[i])
        run.font.size = Pt(9)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 58, 138)

    # ================= 任務一：出處光譜特徵速查 =================
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    sec1_p = doc.add_paragraph()
    sec1_p.paragraph_format.space_before = Pt(4)
    sec1_p.paragraph_format.space_after = Pt(3)
    r1 = sec1_p.add_run("任務一：認識數位文本出處光譜（特性與識讀重點）")
    r1.font.size = Pt(11)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(30, 58, 138)

    # 五大出處表格
    t1 = doc.add_table(rows=6, cols=4)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.autofit = False
    t1_widths = [Inches(1.5), Inches(1.8), Inches(2.3), Inches(1.4)]
    
    t1_headers = ["出處類型", "代表實例", "核心優勢與可靠價值", "識讀陷阱與盲點"]
    for j, cell in enumerate(t1.rows[0].cells):
        cell.width = t1_widths[j]
        set_cell_background(cell, "DBEAFE")
        set_cell_margins(cell, top=60, bottom=60, left=70, right=70)
        set_cell_border(cell, top={"sz":4,"color":"93C5FD"}, bottom={"sz":6,"color":"3B82F6"},
                              left={"sz":4,"color":"CBD5E1"}, right={"sz":4,"color":"CBD5E1"})
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(t1_headers[j])
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(30, 58, 138)

    sources_data = [
        ("1. 政府與公部門\n(.gov.tw)", "臺北市政府觀傳局\n文化局文資資料庫", "數據、法規、公告具法律權威，一手資訊準確。", "多為政績宣傳與觀光推廣，較少探討負面爭議或民間困境。"),
        ("2. 協作知識庫\n(如：維基百科)", "中文維基百科\n國家文化記憶庫", "架構完整、收錄歷史沿革，附註腳供溯源查核。", "人人皆可匿名編輯，需檢查文末註腳，不可全盤盲信。"),
        ("3. 專業調查新聞\n(如：報導者/公視)", "《報導者》深度專題\n公視新聞實測追蹤", "具名記者實地採訪、引用客觀數據、正反平衡報導。", "具有議題倡議色彩，需檢視個案抽樣是否足以代表全貌。"),
        ("4. 社群與部落格\n(Dcard/PTT/IG)", "美食旅遊部落格\n在地鄉民論壇閒聊", "第一線生活現場感生動、用語親切貼近常民經驗。", "高度主觀、缺乏客觀考證，常隱含商業置入（業配文）。"),
        ("5. 內容農場\n(Clickbait 農場站)", "每日爆笑網、奇趣熱榜\n無名農場轉載站", "完全無教學與參考價值！純為騙取點閱率與廣告費。", "標題聳動（震驚/老爺爺痛哭）、匿名無日期、廣告滿版。")
    ]

    for row_idx, data in enumerate(sources_data, start=1):
        row = t1.rows[row_idx]
        set_row_cant_split(row)
        bg_color = "FFFFFF" if row_idx % 2 == 1 else "F8FAFC"
        for col_idx, text in enumerate(data):
            cell = row.cells[col_idx]
            cell.width = t1_widths[col_idx]
            set_cell_background(cell, bg_color)
            set_cell_margins(cell, top=50, bottom=50, left=60, right=60)
            set_cell_border(cell, top={"sz":2,"color":"E2E8F0"}, bottom={"sz":2,"color":"E2E8F0"},
                                  left={"sz":2,"color":"E2E8F0"}, right={"sz":2,"color":"E2E8F0"})
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.size = Pt(8)
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(15, 23, 42)
            else:
                r.font.color.rgb = RGBColor(51, 65, 85)

    # ================= 任務二：五星評鑑雷達法寶規準 =================
    sec2_p = doc.add_paragraph()
    sec2_p.paragraph_format.space_before = Pt(6)
    sec2_p.paragraph_format.space_after = Pt(3)
    r2 = sec2_p.add_run("任務二：掌握五星評鑑雷達法寶（CRAAP+ 原則）")
    r2.font.size = Pt(11)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(30, 58, 138)

    t2 = doc.add_table(rows=6, cols=3)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2.autofit = False
    t2_widths = [Inches(1.8), Inches(2.2), Inches(3.0)]

    t2_headers = ["評鑑指標向度", "自我提問核心", "評分檢核標準 (1~5★)"]
    for j, cell in enumerate(t2.rows[0].cells):
        cell.width = t2_widths[j]
        set_cell_background(cell, "EFF6FF")
        set_cell_margins(cell, top=60, bottom=60, left=70, right=70)
        set_cell_border(cell, top={"sz":4,"color":"93C5FD"}, bottom={"sz":6,"color":"3B82F6"},
                              left={"sz":4,"color":"CBD5E1"}, right={"sz":4,"color":"CBD5E1"})
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(t2_headers[j])
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(30, 58, 138)

    criteria_data = [
        ("1. 權威歸屬 (Authority)", "這篇是誰寫的？他憑什麼？", "是否有具名作者、專業機構或編輯團隊？作者是否具備相關專業？"),
        ("2. 佐證查核 (Accuracy)", "證據在哪裡？能夠查驗嗎？", "是否附有一手統計數據、受訪具名、公文字號或可點擊的外部文獻來源？"),
        ("3. 立場動機 (Purpose)", "為何而寫？想賣東西還是宣傳？", "是客觀傳播、教育研究，還是隱藏商業置入（業配開箱）或單一立場政令？"),
        ("4. 時效歷程 (Currency)", "什麼時候寫的？資訊更新了嗎？", "是否有明確的時間戳記？所提之法規、捷運路線或開放時間是否過時？"),
        ("5. 格式避雷 (Presentation)", "標題是否騙人？排版乾淨嗎？", "標題與內文是否吻合？排版是否乾淨（無農場聳動語句、簡轉繁破綻或滿版廣告）？")
    ]

    for row_idx, data in enumerate(criteria_data, start=1):
        row = t2.rows[row_idx]
        set_row_cant_split(row)
        bg_color = "FFFFFF" if row_idx % 2 == 1 else "F8FAFC"
        for col_idx, text in enumerate(data):
            cell = row.cells[col_idx]
            cell.width = t2_widths[col_idx]
            set_cell_background(cell, bg_color)
            set_cell_margins(cell, top=50, bottom=50, left=60, right=60)
            set_cell_border(cell, top={"sz":2,"color":"E2E8F0"}, bottom={"sz":2,"color":"E2E8F0"},
                                  left={"sz":2,"color":"E2E8F0"}, right={"sz":2,"color":"E2E8F0"})
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.size = Pt(8)
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(15, 23, 42)
            else:
                r.font.color.rgb = RGBColor(51, 65, 85)

    # ================= 任務三：搜尋密技速查工具箱 =================
    sec3_p = doc.add_paragraph()
    sec3_p.paragraph_format.space_before = Pt(6)
    sec3_p.paragraph_format.space_after = Pt(3)
    r3 = sec3_p.add_run("任務三：進階搜尋密技速查工具箱（Boolean Operators）")
    r3.font.size = Pt(11)
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(30, 58, 138)

    t3 = doc.add_table(rows=2, cols=4)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    t3.autofit = False
    t3_widths = [Inches(1.75), Inches(1.75), Inches(1.75), Inches(1.75)]

    tips = [
        ("① 雙引號 \" \" (精確比對)", "\"臺北老屋活化\"", "不把字詞拆散，鎖定完整連續名詞。"),
        ("② 減號 - (排除雜訊)", "大稻埕歷史 -美食 -打卡", "過濾浮濫業配、觀光打卡食記干擾。"),
        ("③ site: (限定權威網域)", "site:gov.tw 臺北大巨蛋", "專門抓取政府公報、法規與審查成果。"),
        ("④ filetype: (抓取研究報告)", "臺北夜市 演變 filetype:pdf", "直接挖出專家學者的論文或白皮書。")
    ]

    for j, (title, eg, desc) in enumerate(tips):
        # Header cell
        c_head = t3.rows[0].cells[j]
        c_head.width = t3_widths[j]
        set_cell_background(c_head, "EFF6FF")
        set_cell_margins(c_head, top=50, bottom=50, left=60, right=60)
        set_cell_border(c_head, top={"sz":4,"color":"93C5FD"}, bottom={"sz":4,"color":"93C5FD"},
                                left={"sz":2,"color":"CBD5E1"}, right={"sz":2,"color":"CBD5E1"})
        p_h = c_head.paragraphs[0]
        p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_h.paragraph_format.space_before = Pt(0)
        p_h.paragraph_format.space_after = Pt(0)
        rh = p_h.add_run(title)
        rh.font.size = Pt(8)
        rh.font.bold = True
        rh.font.color.rgb = RGBColor(30, 58, 138)

        # Content cell
        c_body = t3.rows[1].cells[j]
        c_body.width = t3_widths[j]
        set_cell_background(c_body, "FFFFFF")
        set_cell_margins(c_body, top=50, bottom=50, left=60, right=60)
        set_cell_border(c_body, top={"sz":2,"color":"E2E8F0"}, bottom={"sz":4,"color":"CBD5E1"},
                                left={"sz":2,"color":"CBD5E1"}, right={"sz":2,"color":"CBD5E1"})
        p_b = c_body.paragraphs[0]
        p_b.paragraph_format.space_before = Pt(0)
        p_b.paragraph_format.space_after = Pt(0)
        r_eg = p_b.add_run(f"範例：{eg}\n")
        r_eg.font.size = Pt(7.5)
        r_eg.font.bold = True
        r_eg.font.color.rgb = RGBColor(220, 38, 38)
        r_desc = p_b.add_run(desc)
        r_desc.font.size = Pt(7.5)
        r_desc.font.color.rgb = RGBColor(71, 85, 105)

    # ================= 任務四：自主主題探究任務單 (學生實作區) =================
    doc.add_page_break()

    sec4_p = doc.add_paragraph()
    sec4_p.paragraph_format.space_before = Pt(0)
    sec4_p.paragraph_format.space_after = Pt(3)
    r4 = sec4_p.add_run("任務四：自主主題網羅與五星評鑑探究單（實作產出）")
    r4.font.size = Pt(12)
    r4.font.bold = True
    r4.font.color.rgb = RGBColor(30, 58, 138)

    desc_p = doc.add_paragraph()
    desc_p.paragraph_format.space_after = Pt(4)
    r_desc4 = desc_p.add_run("請選擇一個感興趣的臺北主題，使用進階搜尋語法找到 2 篇不同出處的數位文本，依五星評鑑指標打分並完成分析：")
    r_desc4.font.size = Pt(8.5)
    r_desc4.font.color.rgb = RGBColor(100, 116, 139)

    # 主題與關鍵字選擇區
    meta_table = doc.add_table(rows=2, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    meta_widths = [Inches(3.5), Inches(3.5)]

    # Row 1: 主題
    c0 = meta_table.rows[0].cells[0]
    c0.width = meta_widths[0]
    set_cell_background(c0, "F8FAFC")
    set_cell_margins(c0, top=60, bottom=60, left=70, right=70)
    set_cell_border(c0, top={"sz":4,"color":"CBD5E1"}, bottom={"sz":2,"color":"CBD5E1"}, left={"sz":4,"color":"CBD5E1"}, right={"sz":2,"color":"CBD5E1"})
    p0 = c0.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(0)
    p0.add_run("【探究主題】（請勾選或自訂）：\n").font.bold = True
    p0.runs[0].font.size = Pt(8.5)
    p0.add_run("□ 大稻埕老屋文創與保存　□ 臺北捷運與生活變遷\n□ 臺北夜市庶民飲食文化　□ 大巨蛋公共安全疏散\n□ 自訂主題：___________________________________").font.size = Pt(8)

    # Row 1: 關鍵字
    c1 = meta_table.rows[0].cells[1]
    c1.width = meta_widths[1]
    set_cell_background(c1, "F8FAFC")
    set_cell_margins(c1, top=60, bottom=60, left=70, right=70)
    set_cell_border(c1, top={"sz":4,"color":"CBD5E1"}, bottom={"sz":2,"color":"CBD5E1"}, left={"sz":2,"color":"CBD5E1"}, right={"sz":4,"color":"CBD5E1"})
    p1 = c1.paragraphs[0]
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(0)
    p1.add_run("【進階搜尋語法組合】（如：\" \"、-、site:）：\n").font.bold = True
    p1.runs[0].font.size = Pt(8.5)
    p1.add_run("我使用的關鍵字：\n\n________________________________________________").font.size = Pt(8)

    # 合併 Row 1 下方空列作為備註提示
    meta_table.rows[1].cells[0].merge(meta_table.rows[1].cells[1])
    c_all = meta_table.rows[1].cells[0]
    set_cell_background(c_all, "EFF6FF")
    set_cell_margins(c_all, top=40, bottom=40, left=70, right=70)
    set_cell_border(c_all, top={"sz":2,"color":"CBD5E1"}, bottom={"sz":4,"color":"CBD5E1"}, left={"sz":4,"color":"CBD5E1"}, right={"sz":4,"color":"CBD5E1"})
    p_all = c_all.paragraphs[0]
    p_all.paragraph_format.space_before = Pt(0)
    p_all.paragraph_format.space_after = Pt(0)
    r_hint = p_all.add_run("※ 提示：請盡量挑選不同出處（例如：一篇官方/新聞專題，對比一篇社群部落格或內容農場），評鑑對比將更鮮明！")
    r_hint.font.size = Pt(8)
    r_hint.font.bold = True
    r_hint.font.color.rgb = RGBColor(29, 78, 216)

    # 兩篇文本雙欄評鑑表
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    compare_table = doc.add_table(rows=8, cols=3)
    compare_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    compare_table.autofit = False
    cmp_widths = [Inches(1.6), Inches(2.7), Inches(2.7)]

    cmp_headers = ["評鑑項目", "數位文本 A", "數位文本 B"]
    for j, cell in enumerate(compare_table.rows[0].cells):
        cell.width = cmp_widths[j]
        set_cell_background(cell, "DBEAFE")
        set_cell_margins(cell, top=60, bottom=60, left=70, right=70)
        set_cell_border(cell, top={"sz":4,"color":"93C5FD"}, bottom={"sz":6,"color":"3B82F6"},
                              left={"sz":4,"color":"CBD5E1"}, right={"sz":4,"color":"CBD5E1"})
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(cmp_headers[j])
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(30, 58, 138)

    rows_def = [
        ("1. 文本標題與網址", "標題：\n網址/平台：", "標題：\n網址/平台："),
        ("2. 出處類型歸屬", "□ 官方政府　□ 協作百科\n□ 專業新聞　□ 社群網紅\n□ 內容農場", "□ 官方政府　□ 協作百科\n□ 專業新聞　□ 社群網紅\n□ 內容農場"),
        ("3. 作者身分與權威度\n(1~5★)", "作者：\n專業背景：\n給分：□1★ □2★ □3★ □4★ □5★", "作者：\n專業背景：\n給分：□1★ □2★ □3★ □4★ □5★"),
        ("4. 引用佐證充分性\n(1~5★)", "是否有數據、採訪或文獻？\n\n給分：□1★ □2★ □3★ □4★ □5★", "是否有數據、採訪或文獻？\n\n給分：□1★ □2★ □3★ □4★ □5★"),
        ("5. 立場動機與客觀度\n(1~5★)", "目的（傳播/宣傳/業配/情緒）：\n\n給分：□1★ □2★ □3★ □4★ □5★", "目的（傳播/宣傳/業配/情緒）：\n\n給分：□1★ □2★ □3★ □4★ □5★"),
        ("6. 時效歷程與避雷檢視\n(1~5★)", "發布日期：\n是否為聳動騙點閱/廣告滿版？\n給分：□1★ □2★ □3★ □4★ □5★", "發布日期：\n是否為聳動騙點閱/廣告滿版？\n給分：□1★ □2★ □3★ □4★ □5★"),
        ("7. 該文本最大優勢與盲點", "優勢：\n\n盲點：", "優勢：\n\n盲點：")
    ]

    for row_idx, (col0, col1, col2) in enumerate(rows_def, start=1):
        row = compare_table.rows[row_idx]
        set_row_cant_split(row)
        for col_idx, text in enumerate([col0, col1, col2]):
            cell = row.cells[col_idx]
            cell.width = cmp_widths[col_idx]
            bg_color = "F8FAFC" if col_idx == 0 else "FFFFFF"
            set_cell_background(cell, bg_color)
            set_cell_margins(cell, top=50, bottom=50, left=60, right=60)
            set_cell_border(cell, top={"sz":2,"color":"E2E8F0"}, bottom={"sz":2,"color":"E2E8F0"},
                                  left={"sz":2,"color":"E2E8F0"}, right={"sz":2,"color":"E2E8F0"})
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.size = Pt(8)
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(15, 23, 42)
            else:
                r.font.color.rgb = RGBColor(71, 85, 105)

    # 綜合心得反思區
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    ref_table = doc.add_table(rows=1, cols=1)
    ref_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    ref_table.autofit = False
    ref_cell = ref_table.rows[0].cells[0]
    ref_cell.width = Inches(7.0)
    set_cell_background(ref_cell, "F8FAFC")
    set_cell_margins(ref_cell, top=70, bottom=70, left=80, right=80)
    set_cell_border(ref_cell, top={"sz":4,"color":"93C5FD"}, bottom={"sz":4,"color":"93C5FD"},
                              left={"sz":4,"color":"93C5FD"}, right={"sz":4,"color":"93C5FD"})
    p_ref = ref_cell.paragraphs[0]
    p_ref.paragraph_format.space_before = Pt(0)
    p_ref.paragraph_format.space_after = Pt(0)
    p_ref.add_run("【跨文本思辨與反思總結】（約 100~150 字）：\n").font.bold = True
    p_ref.runs[0].font.size = Pt(8.5)
    p_ref.runs[0].font.color.rgb = RGBColor(30, 58, 138)
    p_ref.add_run("問題引導：比較這兩篇文本後，你發現了哪些視角上的差異？若只看其中一篇可能會產生什麼誤解？我們以後上網搜尋資料時，應該如何避免被片面訊息牽著鼻子走？\n\n\n\n\n\n").font.size = Pt(8)

    # 頁尾註記
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer_p.paragraph_format.space_before = Pt(4)
    r_foot = footer_p.add_run("【學思軒・閱思龍門教學平台】數位媒體識讀系列教材 © 2026 學思軒\n><> ><> ><> ˙˙˙本網頁由黃于珊老師利用Google antigravity 2.0生成製作，僅供教學使用。˙˙˙<>< <>< <><")
    r_foot.font.size = Pt(7.5)
    r_foot.font.color.rgb = RGBColor(148, 163, 184)

    # 儲存檔案
    output_filename = "c:/2026_網羅臺北/[八上閱思龍門]網羅臺北——數位媒體識讀_簡要學習單.docx"
    doc.save(output_filename)
    print(f"成功產出學習單：{output_filename}")

if __name__ == "__main__":
    create_worksheet()
