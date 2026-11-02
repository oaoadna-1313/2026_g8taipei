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

def add_highlight_run(p, text, color_idx="yellow", bold=True):
    # color_idx: yellow, cyan, magenta, green
    r = p.add_run(text)
    r.font.size = Pt(8.5)
    r.font.bold = bold
    rPr = r._r.get_or_add_rPr()
    hl = parse_xml(f'<w:highlight {nsdecls("w")} w:val="{color_idx}"/>')
    rPr.append(hl)
    return r

def create_worksheet():
    doc = docx.Document()
    
    # 邊界設定 (2 cm)
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    # 樣式基礎字型
    normal_style = doc.styles['Normal']
    normal_style.font.name = '微軟正黑體'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')

    # ================= 1. 表頭大標題 =================
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    
    run_badge = title_p.add_run("[閱思龍門]透析臺北\n")
    run_badge.font.size = Pt(9.5)
    run_badge.font.bold = True
    run_badge.font.color.rgb = RGBColor(16, 128, 67)

    run_title = title_p.add_run("《透析臺北——文本比較與觀點思辨》學習單")
    run_title.font.size = Pt(16.5)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(15, 23, 42)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_after = Pt(8)
    run_sub = sub_p.add_run("文本證據標註 ❶~❺ × 空白五維雷達圖繪製 × 跨文本矩陣統整 × 觀點思辨天平")
    run_sub.font.size = Pt(9)
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
        set_cell_background(cell, "F8FAFC")
        set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
        set_cell_border(cell, top={"sz":4,"color":"CBD5E1"}, bottom={"sz":4,"color":"CBD5E1"},
                              left={"sz":4,"color":"CBD5E1"}, right={"sz":4,"color":"CBD5E1"})
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(headers[i])
        r.font.bold = (i >= 3)
        r.font.size = Pt(9)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 學習目標與證據標示指引 Box
    target_table = doc.add_table(rows=1, cols=1)
    target_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_cell = target_table.rows[0].cells[0]
    t_cell.width = Inches(7.0)
    set_cell_background(t_cell, "F0FDF4")
    set_cell_margins(t_cell, top=80, bottom=80, left=120, right=120)
    set_cell_border(t_cell, top={"sz":6,"color":"86EFAC"}, bottom={"sz":6,"color":"86EFAC"},
                           left={"sz":16,"color":"16A34A"}, right={"sz":6,"color":"86EFAC"})
    tp = t_cell.paragraphs[0]
    tp.paragraph_format.space_after = Pt(1)
    r1 = tp.add_run("🎯 閱讀理解核心任務（請善用文本螢光標記）：\n")
    r1.font.bold = True
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = RGBColor(22, 101, 52)
    r2 = tp.add_run(
        "閱讀四大文本時，請依據編號對照各要點之【文本證據】：\n"
        "❶ 觀察視角（文人／都會／生態／青年）  ❷ 空間焦點（老巷／商圈套房／群山公園／捷運）\n"
        "❸ 生活節奏（閒散／高速／四時自然／緊繃） ❹ 人際溫度（常民／心牆／物我／冷漠）  ❺ 核心觀點（一言以蔽之）"
    )
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # ================= 2. 任務一：四大文本精讀與螢光筆證據 =================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(6)
    h1.paragraph_format.space_after = Pt(4)
    run_h1 = h1.add_run("任務一：四大選篇文本精讀與【文本證據編號標註】")
    run_h1.font.size = Pt(11.5)
    run_h1.font.bold = True
    run_h1.font.color.rgb = RGBColor(30, 41, 59)

    # 文本一：舒國治
    tbox1 = doc.add_table(rows=1, cols=1)
    tbox1.alignment = WD_TABLE_ALIGNMENT.CENTER
    c1 = tbox1.rows[0].cells[0]
    c1.width = Inches(7.0)
    set_cell_background(c1, "FFFBEB")
    set_cell_margins(c1, top=70, bottom=70, left=100, right=100)
    set_cell_border(c1, top={"sz":4,"color":"D97706"}, bottom={"sz":4,"color":"D97706"},
                        left={"sz":12,"color":"D97706"}, right={"sz":4,"color":"D97706"})
    p1 = c1.paragraphs[0]
    r_t1 = p1.add_run("📌 文本一：舒國治〈水渠與老巷弄的浮生半日〉（選自《水城臺北》/《理想的下午》）\n")
    r_t1.font.bold = True
    r_t1.font.size = Pt(9)
    p1.add_run("「台北的好，在於它有許多讓人可以隨意散步的小巷。早晨陽光斜照在灰撲撲的公寓陽台上，九重葛垂落下來。走進").font.size = Pt(8.5)
    add_highlight_run(p1, "❷ 永康街或溫州街的巷弄", "yellow")
    p1.add_run("裡，你甚至可以看見昔日").font.size = Pt(8.5)
    add_highlight_run(p1, "❷ 水圳留下來的彎曲路徑", "yellow")
    p1.add_run("……台北是個盆地，古早時候是水城。即便今日被柏油路封蓋，但那股濕潤、").font.size = Pt(8.5)
    add_highlight_run(p1, "❸ 隨興、容許人無所事事", "green")
    p1.add_run("吃一碗").font.size = Pt(8.5)
    add_highlight_run(p1, "❹ 牛肉麵", "cyan")
    p1.add_run("或在").font.size = Pt(8.5)
    add_highlight_run(p1, "❷ 樹蔭下", "yellow")
    add_highlight_run(p1, "❸ 打盹", "green")
    p1.add_run("的氣質，").font.size = Pt(8.5)
    add_highlight_run(p1, "❺ 依舊是它最迷人的靈魂", "magenta")
    p1.add_run("。").font.size = Pt(8.5)
    add_highlight_run(p1, "❸ 不必趕著去哪裡", "green")
    p1.add_run("，在台北散步，").font.size = Pt(8.5)
    add_highlight_run(p1, "❶ 走迷路了也是一種風雅", "yellow")
    p1.add_run("。」").font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # 文本二：韓良露
    tbox2 = doc.add_table(rows=1, cols=1)
    tbox2.alignment = WD_TABLE_ALIGNMENT.CENTER
    c2 = tbox2.rows[0].cells[0]
    c2.width = Inches(7.0)
    set_cell_background(c2, "FFF1F2")
    set_cell_margins(c2, top=70, bottom=70, left=100, right=100)
    set_cell_border(c2, top={"sz":4,"color":"E11D48"}, bottom={"sz":4,"color":"E11D48"},
                        left={"sz":12,"color":"E11D48"}, right={"sz":4,"color":"E11D48"})
    p2 = c2.paragraphs[0]
    r_t2 = p2.add_run("📌 文本二：韓良露〈這座城市的甜美與殘酷〉（選自《台北的甜美與殘酷》）\n")
    r_t2.font.bold = True
    r_t2.font.size = Pt(9)
    p2.add_run("「台北是甜美的。它有全世界最密集的").font.size = Pt(8.5)
    add_highlight_run(p2, "❷ 二十四小時書店與超商", "yellow")
    p2.add_run("、深夜不熄的").font.size = Pt(8.5)
    add_highlight_run(p2, "❷ 夜市小吃", "yellow")
    p2.add_run("、四通八達的捷運，以及一種").font.size = Pt(8.5)
    add_highlight_run(p2, "❶ 高度尊重個人隱私的文明疏離", "yellow")
    p2.add_run("。你在這裡可以做任何你想做的人，").font.size = Pt(8.5)
    add_highlight_run(p2, "❹ 沒有人會過度干涉你", "cyan")
    p2.add_run("。但台北同時也是殘酷的。高昂的房價、狹窄如火柴盒的").font.size = Pt(8.5)
    add_highlight_run(p2, "❷ 出租套房", "yellow")
    p2.add_run("、冷酷的職場競爭，以及人與人之間").font.size = Pt(8.5)
    add_highlight_run(p2, "❹ 客氣卻難以穿透的心牆", "cyan")
    p2.add_run("。它像一座").font.size = Pt(8.5)
    add_highlight_run(p2, "❸ 巨大的旋轉舞台，只獎勵跑得夠快的人", "green")
    p2.add_run("。這座城市").font.size = Pt(8.5)
    add_highlight_run(p2, "❺ 擁抱你的夢想，卻也隨時冷眼看著你被現實吞沒", "magenta")
    p2.add_run("。」").font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # 文本三：劉克襄
    tbox3 = doc.add_table(rows=1, cols=1)
    tbox3.alignment = WD_TABLE_ALIGNMENT.CENTER
    c3 = tbox3.rows[0].cells[0]
    c3.width = Inches(7.0)
    set_cell_background(c3, "F0FDF4")
    set_cell_margins(c3, top=70, bottom=70, left=100, right=100)
    set_cell_border(c3, top={"sz":4,"color":"16A34A"}, bottom={"sz":4,"color":"16A34A"},
                        left={"sz":12,"color":"16A34A"}, right={"sz":4,"color":"16A34A"})
    p3 = c3.paragraphs[0]
    r_t3 = p3.add_run("📌 文本三：劉克襄〈群山環抱的綠意秘密〉（選自《安靜的台北》）\n")
    r_t3.font.bold = True
    r_t3.font.size = Pt(9)
    p3.add_run("「許多人以為台北只是一座被混凝土包覆的冷漠巨獸，但只要你搭上一班捷運或公車，").font.size = Pt(8.5)
    add_highlight_run(p3, "❸ 不到半小時", "green")
    p3.add_run("，你就能從最喧囂的信義商圈鑽進").font.size = Pt(8.5)
    add_highlight_run(p3, "❷ 象山、四獸山的林蔭深處", "yellow")
    p3.add_run("。台北最得天獨厚之處，在於它是一座").font.size = Pt(8.5)
    add_highlight_run(p3, "❺ 被群山與雙河擁抱的盆地", "magenta")
    p3.add_run("。當").font.size = Pt(8.5)
    add_highlight_run(p3, "❷ 大安森林公園的黑冠麻鷺", "yellow")
    p3.add_run("在草坪上").font.size = Pt(8.5)
    add_highlight_run(p3, "❸ 悠閒漫步", "green")
    p3.add_run("，當").font.size = Pt(8.5)
    add_highlight_run(p3, "❷ 基隆河岸的紅樹林", "yellow")
    p3.add_run("迎來冬候鳥，你才會猛然發覺：").font.size = Pt(8.5)
    add_highlight_run(p3, "❺ 自然從未離開台北", "magenta")
    p3.add_run("，它始終在城市的屋簷下與水泥縫隙中").font.size = Pt(8.5)
    add_highlight_run(p3, "❹ 頑強呼吸", "cyan")
    p3.add_run("。學會走出商場、").font.size = Pt(8.5)
    add_highlight_run(p3, "❶ 抬頭看山，你才算真正認識了台北", "yellow")
    p3.add_run("。」").font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # 文本四：李維
    tbox4 = doc.add_table(rows=1, cols=1)
    tbox4.alignment = WD_TABLE_ALIGNMENT.CENTER
    c4 = tbox4.rows[0].cells[0]
    c4.width = Inches(7.0)
    set_cell_background(c4, "F0F9FF")
    set_cell_margins(c4, top=70, bottom=70, left=100, right=100)
    set_cell_border(c4, top={"sz":4,"color":"0284C7"}, bottom={"sz":4,"color":"0284C7"},
                        left={"sz":12,"color":"0284C7"}, right={"sz":4,"color":"0284C7"})
    p4 = c4.paragraphs[0]
    r_t4 = p4.add_run("📌 文本四：李維〈在捷運尖峰時刻，我找不到自己的影子〉（選自《北漂青年生存記事》）\n")
    r_t4.font.bold = True
    r_t4.font.size = Pt(9)
    p4.add_run("「").font.size = Pt(8.5)
    add_highlight_run(p4, "❷ 捷運忠孝復興站", "yellow")
    p4.add_run("的手扶梯像一條").font.size = Pt(8.5)
    add_highlight_run(p4, "❸ 永不停歇的輸送帶", "green")
    p4.add_run("，所有人").font.size = Pt(8.5)
    add_highlight_run(p4, "❹ 面無表情地盯著手機螢幕", "cyan")
    p4.add_run("，以").font.size = Pt(8.5)
    add_highlight_run(p4, "❸ 精密計算過的步伐快速穿梭", "green")
    p4.add_run("。每個月五號薪水入帳，轉身就有一半付給了").font.size = Pt(8.5)
    add_highlight_run(p4, "❷ 頂樓加蓋、沒有對外窗的小套房", "yellow")
    p4.add_run("。在文青咖啡館的落地窗外，").font.size = Pt(8.5)
    add_highlight_run(p4, "❷ 信義區的霓虹燈", "yellow")
    p4.add_run("璀璨得像一場不真實的夢境。台北給了我們最豐富的藝文展覽、最多元的人生可能，但它也每天都在質問我們：你配得上這座城市的繁華嗎？").font.size = Pt(8.5)
    add_highlight_run(p4, "❶ 對於北漂族而言", "yellow")
    p4.add_run("，台北是造夢的天堂，也是讓人").font.size = Pt(8.5)
    add_highlight_run(p4, "❺ 精疲力竭的修煉場", "magenta")
    p4.add_run("。」").font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 任務一填空表格（加大第一欄寬度，防止折字）
    q1_p = doc.add_paragraph()
    q1_p.add_run("【核心要點與文本證據填答表】請將摘錄要點與對應的文本證據引文填入表中：").font.bold = True
    q1_p.runs[0].font.size = Pt(9.5)

    table_extract = doc.add_table(rows=6, cols=5)
    table_extract.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_extract.autofit = False

    t_widths = [Inches(1.5), Inches(1.37), Inches(1.37), Inches(1.37), Inches(1.37)]
    headers_ext = ["分析向度與證據編號", "文本一・舒國治", "文本二・韓良露", "文本三・劉克襄", "文本四・李維"]
    rows_ext_labels = [
        "❶ 觀察視角\n(身分與態度)",
        "❷ 空間焦點\n(關鍵地景場域)",
        "❸ 生活步調\n(節奏快慢與氛圍)",
        "❹ 人際溫度\n(人際互動模式)",
        "❺ 核心觀點\n(一言以蔽之)"
    ]

    hdr_row = table_extract.rows[0]
    set_row_cant_split(hdr_row)
    for j, cell in enumerate(hdr_row.cells):
        cell.width = t_widths[j]
        set_cell_background(cell, "E2E8F0")
        set_cell_margins(cell, top=70, bottom=70, left=70, right=70)
        set_cell_border(cell, top={"sz":6,"color":"94A3B8"}, bottom={"sz":6,"color":"94A3B8"},
                              left={"sz":4,"color":"CBD5E1"}, right={"sz":4,"color":"CBD5E1"})
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(headers_ext[j])
        r.font.bold = True
        r.font.size = Pt(8.5)

    for i in range(1, 6):
        r_row = table_extract.rows[i]
        set_row_cant_split(r_row)
        for j, cell in enumerate(r_row.cells):
            cell.width = t_widths[j]
            set_cell_margins(cell, top=90, bottom=90, left=70, right=70)
            set_cell_border(cell, top={"sz":4,"color":"E2E8F0"}, bottom={"sz":4,"color":"E2E8F0"},
                                  left={"sz":4,"color":"E2E8F0"}, right={"sz":4,"color":"E2E8F0"})
            p = cell.paragraphs[0]
            if j == 0:
                set_cell_background(cell, "F8FAFC")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(rows_ext_labels[i-1])
                r.font.bold = True
                r.font.size = Pt(8)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run("")
                r.font.size = Pt(8)

    doc.add_page_break()

    # ================= 3. 任務二：空白五維雷達圖手繪與矩陣對照 =================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(4)
    h2.paragraph_format.space_after = Pt(4)
    run_h2 = h2.add_run("任務二：圖表統整——五維度雷達圖手繪與矩陣對照")
    run_h2.font.size = Pt(11.5)
    run_h2.font.bold = True
    run_h2.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph("【五維度城市感知量化評估】請依據文本閱讀印象，在左表給予評分（1~10分），並在右側空白雷達網格上點標連線（請用四種顏色筆標註四大文本，亦可畫出自己的日常生活折線）：").runs[0].font.size = Pt(9)

    chart_area_table = doc.add_table(rows=1, cols=2)
    chart_area_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_left = chart_area_table.rows[0].cells[0]
    c_right = chart_area_table.rows[0].cells[1]
    c_left.width = Inches(3.4)
    c_right.width = Inches(3.6)
    set_cell_margins(c_left, top=70, bottom=70, left=70, right=70)
    set_cell_margins(c_right, top=70, bottom=70, left=70, right=70)

    # 左側：評分表格
    p_cl = c_left.paragraphs[0]
    p_cl.paragraph_format.space_after = Pt(2)
    p_cl.add_run("📊 感知指標評分參考表（滿分10分）：\n").font.bold = True
    p_cl.runs[0].font.size = Pt(8.5)
    
    score_sub_table = c_left.add_table(rows=6, cols=6)
    score_sub_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_headers = ["指標", "舒", "韓", "劉", "李", "我"]
    s_rows = [
        ("悠活慢調度", "9.5", "5.0", "8.0", "2.0", "   "),
        ("便利繁華度", "4.0", "9.5", "5.0", "9.0", "   "),
        ("生態自然感", "7.5", "4.0", "9.5", "2.5", "   "),
        ("生存壓力感", "2.0", "7.5", "3.0", "9.5", "   "),
        ("人際親密感", "8.5", "4.5", "7.0", "3.0", "   ")
    ]
    for j, h in enumerate(s_headers):
        sc = score_sub_table.rows[0].cells[j]
        set_cell_background(sc, "F1F5F9")
        p = sc.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8)
    for i, row_data in enumerate(s_rows):
        for j, val in enumerate(row_data):
            sc = score_sub_table.rows[i+1].cells[j]
            p = sc.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8)
            if j == 0:
                r.font.bold = True
                set_cell_background(sc, "F8FAFC")

    # 右側：空白雷達圖網格手繪區（插入高清圖檔，版面不跑版）
    p_cr = c_right.paragraphs[0]
    p_cr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_rt = p_cr.add_run("【五維度雷達網格手繪區（空白待點繪）】\n")
    r_rt.font.bold = True
    r_rt.font.size = Pt(8.5)
    
    radar_img_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "radar_grid_blank.png")
    if os.path.exists(radar_img_path):
        p_cr.add_run().add_picture(radar_img_path, width=Inches(3.3))
    else:
        r_diag = p_cr.add_run(
            "             [生態自然感 (10分)]\n"
            "                     ▲ 10\n"
            "                     |  8\n"
            "             /\\      |  6     /\\\n"
            "            /  \\     |  4    /  \\\n"
            "[悠活慢調度] ------+--2--+------ [便利繁華度]\n"
            " 10  8  6  4  2     | 0    2  4  6  8  10\n"
            "            \\  /     |       \\  /\n"
            "             \\/      |        \\/\n"
            "    [人際親密感]     ▼       [生存壓力感]\n\n"
        )
        r_diag.font.size = Pt(7.5)
        r_diag.font.name = 'Consolas'
        r_diag.font.color.rgb = RGBColor(71, 85, 105)

    p_cr_leg = c_right.add_paragraph()
    p_cr_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_leg = p_cr_leg.add_run("折線圖例： [ ] 舒國治(黃)  [ ] 韓良露(紅)  [ ] 劉克襄(綠)  [ ] 李維(藍)  [ ] 我的日常(紫)")
    r_leg.font.size = Pt(7.5)
    r_leg.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # 圖表解碼問題
    q_chart = doc.add_paragraph()
    q_chart.add_run("🔍 圖表解讀探究思考（請寫出你的觀察）：\n").font.bold = True
    q_chart.runs[0].font.size = Pt(9)
    q_chart.add_run("1. 比較「舒國治」與「李維」的折線形狀，你發現了什麼鏡像反差？背後的社會階級或生活條件有何不同？\n"
                    "答：___________________________________________________________________________________\n"
                    "2. 如果畫上你自己的生活折線，形狀最接近誰？為什麼？\n"
                    "答：___________________________________________________________________________________").font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ================= 4. 任務三：觀點思辨天平 =================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(4)
    h3.paragraph_format.space_after = Pt(4)
    run_h3 = h3.add_run("任務三：觀點思辨天平——我認同作者的說法嗎？")
    run_h3.font.size = Pt(11.5)
    run_h3.font.bold = True
    run_h3.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph("閱讀是心靈的交鋒。請針對以下兩項核心爭鳴議題，勾選你的立場，並寫下具體理由與生活例證：").runs[0].font.size = Pt(9)

    issues = [
        ("思辨議題一：舒國治說「在臺北巷弄隨興晃蕩打盹，是生活最迷人的靈魂」",
         "有評論指出：這種慢活只屬於有閒階級或特定文人的浪漫特權；真實台北人每天背負房貸、通勤塞車，根本無從奢侈慢活。",
         "□ 非常認同舒國治（安頓身心之必要）   □ 部分認同（心靈追求但現實殘酷）   □ 不認同（文人浪漫濾鏡）"),
        ("思辨議題二：韓良露說「人與人之間客氣卻難以穿透的心牆，是一種尊重隱私的文明自由」",
         "有人讚許這種距離讓人自由不受管束；也有人認為這是現代都會人心理孤立、寂寞冷漠的元凶。",
         "□ 認同是文明的自由與邊界   □ 矛盾共存（既自由又孤單）   □ 不認同（實為冷漠與疏離）")
    ]

    for title, desc, opt in issues:
        ibox = doc.add_table(rows=1, cols=1)
        ibox.alignment = WD_TABLE_ALIGNMENT.CENTER
        ic = ibox.rows[0].cells[0]
        ic.width = Inches(7.0)
        set_cell_background(ic, "F8FAFC")
        set_cell_margins(ic, top=70, bottom=70, left=100, right=100)
        set_cell_border(ic, top={"sz":4,"color":"CBD5E1"}, bottom={"sz":4,"color":"CBD5E1"},
                            left={"sz":12,"color":"3B82F6"}, right={"sz":4,"color":"CBD5E1"})
        ip = ic.paragraphs[0]
        ip.paragraph_format.space_after = Pt(2)
        r_it = ip.add_run(f"⚖️ {title}\n")
        r_it.font.bold = True
        r_it.font.size = Pt(9)
        r_it.font.color.rgb = RGBColor(30, 41, 59)
        
        r_id = ip.add_run(f"【爭鳴焦點】{desc}\n")
        r_id.font.size = Pt(8)
        r_id.font.color.rgb = RGBColor(100, 116, 139)
        
        r_io = ip.add_run(f"【我的立場】{opt}\n")
        r_io.font.bold = True
        r_io.font.size = Pt(8.5)
        r_io.font.color.rgb = RGBColor(29, 78, 216)
        
        r_ir = ip.add_run("【我認同／質疑的理由與生活觀察】（100字以內）：\n"
                          "_______________________________________________________________________________________\n"
                          "_______________________________________________________________________________________")
        r_ir.font.size = Pt(8)
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # ================= 5. 任務四：我的臺北微書寫 =================
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(4)
    h4.paragraph_format.space_after = Pt(4)
    run_h4 = h4.add_run("任務四：我的臺北微書寫——城事心聲")
    run_h4.font.size = Pt(11.5)
    run_h4.font.bold = True
    run_h4.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph("【個人觀點產出】如果換你寫一篇臺北，你最想描寫哪一個角落？你的核心意象與情感是什麼？請寫一段 150～200 字的短文：").runs[0].font.size = Pt(9)

    write_table = doc.add_table(rows=1, cols=1)
    write_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    wc = write_table.rows[0].cells[0]
    wc.width = Inches(7.0)
    set_cell_background(wc, "FAFAF9")
    set_cell_margins(wc, top=80, bottom=80, left=100, right=100)
    set_cell_border(wc, top={"sz":4,"color":"D6D3D1"}, bottom={"sz":4,"color":"D6D3D1"},
                        left={"sz":12,"color":"F59E0B"}, right={"sz":4,"color":"D6D3D1"})
    wp = wc.paragraphs[0]
    wp.paragraph_format.space_after = Pt(2)
    wp.add_run("題目自訂：【                                          】\n\n"
               "_________________________________________________________________________________________________\n\n"
               "_________________________________________________________________________________________________\n\n"
               "_________________________________________________________________________________________________\n\n"
               "_________________________________________________________________________________________________").font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ================= 6. 學習評量規準 (Rubric) =================
    rub_p = doc.add_paragraph()
    rub_p.add_run("📋 學生自評與教師多元評量量規（Rubric）：").font.bold = True
    rub_p.runs[0].font.size = Pt(9)

    rub_table = doc.add_table(rows=4, cols=4)
    rub_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    r_widths = [Inches(1.5), Inches(2.25), Inches(2.25), Inches(1.0)]
    r_hdrs = ["評量向度", "優異 (4~5分)", "尚可 (2~3分)", "得分"]
    r_contents = [
        ("要點統整與證據標註 (40%)", "能精確摘錄四大文本五向度，且能正確援引 ❶~❺ 文本原句作為有力論據。", "能大致摘錄主要重點，但部分文本證據引用較模糊。"),
        ("圖表解讀能力 (30%)", "能客觀給予向度評分，完成空白雷達圖手繪，精準解讀作者幾何形狀特徵。", "能完成雷達圖繪製，但對圖表意涵之分析較表面。"),
        ("觀點思辨論述 (30%)", "能清楚表達立場，論據充分且能結合生活經驗，展現獨立批判思維。", "能表達認同與否，但論述理由較為簡略缺乏例證。")
    ]

    for j, h in enumerate(r_hdrs):
        rc = rub_table.rows[0].cells[j]
        set_cell_background(rc, "F1F5F9")
        p = rc.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8)

    for i, (dim, good, fair) in enumerate(r_contents):
        row = rub_table.rows[i+1]
        for j, txt in enumerate([dim, good, fair, ""]):
            rc = row.cells[j]
            p = rc.paragraphs[0]
            r = p.add_run(txt)
            r.font.size = Pt(7.5)
            if j == 0:
                set_cell_background(rc, "F8FAFC")
                r.font.bold = True
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif j == 3:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # 儲存檔案（防範 Word 開啟鎖定衝突）
    output_path = r"c:\2026_透析臺北──文本比較\透析臺北——文本比較與觀點思辨_學生學習單.docx"
    try:
        doc.save(output_path)
        print(f"Successfully generated: {output_path}")
    except PermissionError:
        output_path_alt = r"c:\2026_透析臺北──文本比較\透析臺北——文本比較與觀點思辨_學生學習單(新版).docx"
        doc.save(output_path_alt)
        print(f"原檔案正被 Word 開啟鎖定，已成功另存至新版：{output_path_alt}")

if __name__ == "__main__":
    create_worksheet()
