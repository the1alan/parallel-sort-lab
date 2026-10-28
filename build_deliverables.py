from __future__ import annotations
import csv, os, platform, sys, multiprocessing, textwrap
from pathlib import Path
from datetime import date

from docx import Document
from docx.shared import Inches, Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGB
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from PIL import Image

ROOT = Path(__file__).resolve().parent
RESULTS = ROOT / 'results'
DOCS = ROOT / 'docs'
FIGURES = DOCS / 'figures'
FIGURES.mkdir(parents=True, exist_ok=True)
CSV_PATH = RESULTS / 'benchmark_results.csv'

ALG_MAP = {
    'Merge Sort (последовательный)': 'merge',
    'Parallel Merge Sort': 'parallel',
    'sorted() (встроенная)': 'builtin',
}
with CSV_PATH.open(encoding='utf-8-sig', newline='') as f:
    records = list(csv.DictReader(f))

def get_row(size, algorithm, workers=1, data_type='random'):
    for r in records:
        if int(r['size']) == size and ALG_MAP[r['algorithm']] == algorithm and int(r['workers']) == workers and r['data_type'] == data_type:
            return r
    raise KeyError((size, algorithm, workers, data_type))


def regenerate_charts():
    # Rebuild UI charts only from saved measured values; never synthesize data.
    from benchmark.runner import ResultRow
    from ui.charts import save_charts
    rows = []
    for x in records:
        rows.append(ResultRow(
            algorithm=ALG_MAP[x['algorithm']], size=int(x['size']), workers=int(x['workers']),
            data_type=x['data_type'], repeats=int(x['repeats']), median_time=float(x['median_time_s']),
            speedup=float(x['speedup']) if x['speedup'] else None,
            efficiency=float(x['efficiency']) if x['efficiency'] else None,
            correct=x['correct'].lower() == 'true'))
    charts_dir = RESULTS / 'charts'
    charts_dir.mkdir(parents=True, exist_ok=True)
    for old in charts_dir.glob('*.png'):
        old.unlink()
    return save_charts(rows, charts_dir)


def create_input_type_chart():
    import matplotlib.pyplot as plt
    types = ['random', 'sorted', 'reversed']
    labels = ['Случайные', 'Отсортированные', 'Обратный порядок']
    series = [
        ('Merge Sort', 'merge', 1, '#2563EB'),
        ('sorted()', 'builtin', 1, '#F59E0B'),
        ('Parallel p=2', 'parallel', 2, '#DC2626'),
        ('Parallel p=4', 'parallel', 4, '#7C3AED'),
    ]
    fig, ax = plt.subplots(figsize=(9.4, 4.0), dpi=180)
    x = list(range(len(types))); width = 0.18
    for idx, (label, algo, workers, color) in enumerate(series):
        vals = [float(get_row(50000, algo, workers, kind)['median_time_s']) * 1000 for kind in types]
        offset = (idx - (len(series)-1)/2) * width
        ax.bar([v+offset for v in x], vals, width=width, label=label, color=color)
    ax.set_yscale('log'); ax.set_xticks(x, labels)
    ax.set_ylabel('Медианное время, мс (лог. шкала)')
    ax.set_title('Влияние порядка входных данных (n = 50 000)')
    ax.grid(axis='y', which='both', alpha=.25); ax.set_axisbelow(True)
    ax.legend(ncol=2, frameon=False, fontsize=8)
    fig.tight_layout(); out=FIGURES/'input_types_50000.png'; fig.savefig(out, bbox_inches='tight'); plt.close(fig)
    return out

sizes = sorted({int(r['size']) for r in records})
all_correct = all(r['correct'].lower() == 'true' for r in records)
python_version = sys.version.split()[0]
platform_name = platform.platform()
logical_cpus = multiprocessing.cpu_count()
matplotlib_version = '3.10.8'

# Colors
NAVY = '111827'
PANEL = '1E293B'
PANEL2 = '243247'
TEAL = '5EEAD4'
BLUE = '8AB4FF'
WHITE = 'F8FAFC'
MUTED = 'B8C4D6'
INK = '152235'
LIGHT = 'F3F6FA'
BORDER = 'D9E2EC'
GOLD = 'F8C66D'
RED = 'FB7185'


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), fill); tcPr.append(shd)


def set_cell_margins(cell, top=70, start=100, bottom=70, end=100):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in('w:tcMar')
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar'); tcPr.append(tcMar)
    for m, v in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        node = tcMar.find(qn('w:' + m))
        if node is None:
            node = OxmlElement('w:' + m); tcMar.append(node)
        node.set(qn('w:w'), str(v)); node.set(qn('w:type'), 'dxa')


def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader'); tblHeader.set(qn('w:val'), 'true'); trPr.append(tblHeader)


def set_doc_style(doc):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0); sec.page_height = Cm(29.7)
    sec.top_margin = Cm(1.45); sec.bottom_margin = Cm(1.4)
    sec.left_margin = Cm(1.7); sec.right_margin = Cm(1.7)
    styles = doc.styles
    normal = styles['Normal']; normal.font.name = 'Aptos'; normal.font.size = Pt(9)
    normal.font.color.rgb = RGBColor.from_string(INK)
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.04
    for name, size, color in [('Title', 26, NAVY), ('Heading 1', 16, NAVY), ('Heading 2', 11, '0F766E'), ('Heading 3', 10, NAVY)]:
        st = styles[name]; st.font.name = 'Aptos Display' if name in ['Title','Heading 1'] else 'Aptos'
        st.font.size = Pt(size); st.font.bold = True; st.font.color.rgb = RGBColor.from_string(color)
        st.paragraph_format.space_before = Pt(6); st.paragraph_format.space_after = Pt(4)
    # Header and footer
    header = sec.header.paragraphs[0]
    header.text = 'PARALLEL SORT LAB  /  АЛГОРИТМЫ И СТРУКТУРЫ ДАННЫХ'
    header.style = styles['Caption']; header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in header.runs:
        run.font.name = 'Aptos'; run.font.size = Pt(7); run.font.bold = True; run.font.color.rgb = RGBColor.from_string('64748B')
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = footer.add_run('Учебный проект  •  28 октября 2026  •  ')
    r.font.size = Pt(7); r.font.color.rgb = RGBColor.from_string('64748B')
    fld = OxmlElement('w:fldSimple'); fld.set(qn('w:instr'), 'PAGE'); footer._p.append(fld)


def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(style='List Bullet' if level == 0 else 'List Bullet 2')
    p.paragraph_format.space_after = Pt(2)
    p.add_run(text)
    return p


def add_callout(doc, title, text, fill='E6FFFA'):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0,0); shade(cell, fill); set_cell_margins(cell, 120, 160, 120, 160)
    p = cell.paragraphs[0]
    r = p.add_run(title + '  '); r.bold = True; r.font.color.rgb = RGBColor.from_string('0F766E')
    p.add_run(text)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def add_report_table(doc, headers, rows, font_size=8):
    t = doc.add_table(rows=1, cols=len(headers)); t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.style = 'Table Grid'
    t.autofit = True
    for i, h in enumerate(headers):
        c=t.rows[0].cells[i]; c.text=str(h); shade(c, NAVY); set_cell_margins(c)
        for p in c.paragraphs:
            for r in p.runs:
                r.font.bold=True; r.font.size=Pt(font_size); r.font.color.rgb=RGBColor(255,255,255)
    set_repeat_table_header(t.rows[0])
    for ri, row in enumerate(rows):
        cells=t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text=str(val); set_cell_margins(cells[i])
            cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if ri % 2 == 1: shade(cells[i], 'F1F5F9')
            for p in cells[i].paragraphs:
                p.paragraph_format.space_after=Pt(0)
                for r in p.runs: r.font.size=Pt(font_size)
    return t


def sec_heading(doc, title, sub=None):
    p = doc.add_paragraph(style='Heading 1'); p.add_run(title)
    if sub:
        p2=doc.add_paragraph(sub); p2.paragraph_format.space_after=Pt(5)
        for r in p2.runs: r.font.size=Pt(8); r.font.color.rgb=RGBColor.from_string('64748B')


def create_report_docx():
    doc=Document(); set_doc_style(doc)
    # Page 1
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(18); p.paragraph_format.space_after=Pt(2)
    r=p.add_run('ИССЛЕДОВАТЕЛЬСКИЙ ПРОЕКТ'); r.bold=True; r.font.size=Pt(9); r.font.color.rgb=RGBColor.from_string('0F766E')
    p=doc.add_paragraph(style='Title'); p.paragraph_format.space_after=Pt(2); p.add_run('Parallel Sort Lab')
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(10)
    r=p.add_run('Параллельные алгоритмы сортировки'); r.font.size=Pt(16); r.bold=True; r.font.color.rgb=RGBColor.from_string('334155')
    meta=doc.add_table(rows=3, cols=2); meta.style='Table Grid'; meta.alignment=WD_TABLE_ALIGNMENT.LEFT
    meta_data=[('Дисциплина','Алгоритмы и структуры данных'),('Команда','5 участников — ФИО добавить перед сдачей'),('Технологии','Python 3 · Tkinter · Matplotlib · ProcessPoolExecutor')]
    for i,(a,b) in enumerate(meta_data):
        meta.cell(i,0).text=a; meta.cell(i,1).text=b; shade(meta.cell(i,0),'E2E8F0')
        for c in meta.rows[i].cells:
            set_cell_margins(c,80,110,80,110)
            for pp in c.paragraphs:
                for rr in pp.runs: rr.font.size=Pt(8)
        for rr in meta.cell(i,0).paragraphs[0].runs: rr.bold=True
    doc.add_paragraph()
    sec_heading(doc,'1. Цель и задачи')
    doc.add_paragraph('Цель проекта — исследовать параллельную сортировку слиянием, реализовать сравнимые варианты сортировки и измерить, при каких условиях параллельная обработка имеет практический смысл.')
    add_bullet(doc,'Реализовать последовательный Merge Sort и Parallel Merge Sort на отдельных процессах.')
    add_bullet(doc,'Сравнить обе реализации со встроенной функцией sorted() на одинаковых входных данных.')
    add_bullet(doc,'Измерить время, рассчитать ускорение и эффективность, проверить корректность результатов.')
    add_bullet(doc,'Разработать настольный интерфейс с таблицей результатов, графиками и экспортом CSV.')
    sec_heading(doc,'2. Теоретическая основа')
    doc.add_paragraph('Merge Sort использует принцип «разделяй и властвуй»: массив делится на две части, каждая часть рекурсивно сортируется, затем две упорядоченные последовательности объединяются линейным слиянием. При равенстве элементов сначала выбирается элемент левой части, поэтому алгоритм устойчив.')
    add_callout(doc,'Рекуррентность','T(n) = 2T(n/2) + O(n)  ⇒  T(n) = O(n log n). Дополнительная память — O(n).')
    doc.add_page_break()
    # Page 2
    sec_heading(doc,'3. Реализации и архитектура')
    doc.add_paragraph('В проекте сравниваются три варианта. Все они получают одинаковый список целых чисел; результат каждого измерения сверяется с sorted(входные_данные). Исходный список не должен изменяться.')
    add_report_table(doc,['Вариант','Реализация','Зачем нужен'],[
        ('Merge Sort','Собственная рекурсивная реализация на Python','Базовая последовательная версия; O(n log n), O(n) памяти.'),
        ('Parallel Merge Sort','Список делится на p частей; части сортируются через ProcessPoolExecutor, затем сливаются в главном процессе','Исследование распараллеливания и накладных расходов.'),
        ('sorted()','Встроенная сортировка Python (Timsort)','Практический ориентир; это не собственная реализация проекта.')], font_size=8)
    sec_heading(doc,'4. Параллельная обработка')
    doc.add_paragraph('Параллельный вариант создаёт пул один раз на вызов, передаёт рабочие части в процессы и выполняет заключительное попарное слияние в главном процессе. Для малых массивов ниже заданного порога используется последовательная сортировка. Для CPU-bound кода применяются процессы: потоки CPython ограничены GIL.')
    add_bullet(doc,'Полное измеренное время включает старт пула, сериализацию и передачу данных, сортировку частей и финальное слияние.')
    add_bullet(doc,'На macOS применяется режим запуска spawn; поэтому код запуска защищён условием if __name__ == "__main__".')
    add_bullet(doc,'Ускорение не гарантировано: фиксированные расходы могут превышать выигрыш от параллельной работы.')
    sec_heading(doc,'5. Анализ сложности')
    add_report_table(doc,['Алгоритм','Время','Дополнительная память'],[
        ('Merge Sort','O(n log n) во всех случаях','O(n)'),
        ('Parallel Merge Sort','Сортировка частей примерно O((n/p) log(n/p)); финальное слияние O(n log p) + расходы процессов','O(n) на главном процессе плюс копии частей'),
        ('sorted() / Timsort','O(n log n); на уже упорядоченных данных возможно O(n)','O(n)')], font_size=8)
    sec_heading(doc,'6. Основные модули')
    doc.add_paragraph('algorithms/ — сортировки; benchmark/ — генерация входа, замеры, метрики и CSV; ui/ — интерфейс и графики; main.py — запуск GUI; cli.py — консольный эксперимент; tests/ — автоматические проверки.')
    doc.add_page_break()
    # Page 3
    sec_heading(doc,'7. Методика эксперимента')
    doc.add_paragraph('Проведено 75 измеряемых запусков: полный набор для случайных данных (3 размера × 5 конфигураций × 3 повтора = 45 запусков) и дополнительные сравнения отсортированных и обратных данных при n = 50 000 (по 15 запусков на каждый тип). Входные данные генерируются с фиксированным seed = 42. Таймер — time.perf_counter(); генерация и проверка корректности не входят в замер. Итог — медиана трёх повторов.')
    add_report_table(doc,['Параметр','Значение'],[
        ('Размеры массива','10 000; 50 000; 100 000 элементов'),
        ('Типы данных','random на трёх размерах; sorted и reversed при n = 50 000'),
        ('Алгоритмы','Merge Sort; sorted(); Parallel Merge Sort при p = 1, 2, 4'),
        ('Повторы','3 на каждую конфигурацию'),
        ('Метрики','S(p) = T_merge / T_parallel(p); E(p) = S(p) / p'),
        ('Проверка','Все 25 строк итогового CSV имеют correct = True')], font_size=8)
    sec_heading(doc,'8. Результаты измерений')
    result_rows=[]
    for n in sizes:
        merge=float(get_row(n,'merge')['median_time_s'])
        builtin=float(get_row(n,'builtin')['median_time_s'])
        p2=float(get_row(n,'parallel',2)['median_time_s'])
        p4=float(get_row(n,'parallel',4)['median_time_s'])
        result_rows.append((f'{n:,}'.replace(',',' '),f'{merge*1000:.2f}',f'{builtin*1000:.2f}',f'{p2*1000:.1f}',f'{p4*1000:.1f}'))
    add_report_table(doc,['n','Merge Sort, мс','sorted(), мс','Parallel p=2, мс','Parallel p=4, мс'],result_rows,font_size=8)
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(2)
    r=p.add_run('Единицы измерения — миллисекунды; показатели являются медианами трёх повторов.'); r.italic=True; r.font.size=Pt(7.5); r.font.color.rgb=RGBColor.from_string('64748B')
    if (RESULTS/'charts'/'time_random.png').exists():
        doc.add_picture(str(RESULTS/'charts'/'time_random.png'), width=Cm(16.3))
        doc.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()
    # Page 4: compare input order for a common size
    sec_heading(doc,'9. Влияние порядка входных данных')
    doc.add_paragraph('Для дополнительного сравнения проведены отдельные замеры на отсортированных и обратно отсортированных массивах длиной 50 000. Каждый тип сравнивался с тем же размером случайных данных, по три повтора на конфигурацию. Это сравнение не расширяет полный размерный набор: оно отвечает только на вопрос о влиянии исходного порядка при фиксированном n.')
    type_rows=[]
    for kind,label in [('random','Случайные'),('sorted','Отсортированные'),('reversed','Обратный порядок')]:
        type_rows.append((label,
            f"{float(get_row(50000,'merge',1,kind)['median_time_s'])*1000:.2f}",
            f"{float(get_row(50000,'builtin',1,kind)['median_time_s'])*1000:.3f}",
            f"{float(get_row(50000,'parallel',2,kind)['median_time_s'])*1000:.1f}",
            f"{float(get_row(50000,'parallel',4,kind)['median_time_s'])*1000:.1f}"))
    add_report_table(doc,['Тип данных (n = 50 000)','Merge Sort, мс','sorted(), мс','Parallel p=2, мс','Parallel p=4, мс'],type_rows,font_size=8)
    p=doc.add_paragraph('Для встроенной sorted() порядок заметно влияет на время: Timsort умеет эффективно использовать уже упорядоченные фрагменты. Для собственной Merge Sort временная сложность остаётся O(n log n), поэтому изменение порядка не должно менять класс сложности.')
    p.paragraph_format.space_before=Pt(4)
    if (FIGURES/'input_types_50000.png').exists():
        doc.add_picture(str(FIGURES/'input_types_50000.png'), width=Cm(16.3))
        doc.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()
    # Page 5: interpret and conclude
    sec_heading(doc,'10. Интерпретация ускорения')
    doc.add_paragraph('В данном окружении Parallel Merge Sort с p = 2 и p = 4 медленнее последовательной версии на всех протестированных конфигурациях. Каждый замер создаёт новый пул процессов; на выбранных размерах запуск процессов, сериализация/передача списков и последовательное слияние перекрывают выигрыш от распараллеливания.')
    t=doc.add_table(rows=1,cols=2); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.autofit=True
    for i,name in enumerate(['speedup_random.png','efficiency_random.png']):
        f=RESULTS/'charts'/name
        if f.exists():
            para=t.cell(0,i).paragraphs[0]; para.alignment=WD_ALIGN_PARAGRAPH.CENTER
            para.add_run().add_picture(str(f), width=Cm(8.0))
        set_cell_margins(t.cell(0,i),50,50,50,50)
    r2=float(get_row(100000,'parallel',2)['speedup']); e2=float(get_row(100000,'parallel',2)['efficiency'])
    p=doc.add_paragraph(f'При n = 100 000 наблюдается S(2) = {r2:.3f} и E(2) = {e2:.3f}. Так как S(p) < 1, параллельный вызов медленнее базового. Для p = 4 S(4) ≈ {float(get_row(100000,"parallel",4)["speedup"]):.3f}.')
    p.paragraph_format.space_before=Pt(2)
    sec_heading(doc,'11. Выводы и ограничения')
    add_bullet(doc,'Корректность подтверждена для всех 25 итоговых строк CSV: каждый результат совпал с sorted(вход).')
    add_bullet(doc,'При n = 100 000 sorted() заняла 13,35 мс против 172,23 мс у Python Merge Sort; это сравнение оптимизированной встроенной реализации с учебной реализацией на Python.')
    add_bullet(doc,'Параллельный вариант не дал ускорения на текущем стенде. Возможное улучшение — переиспользовать пул между повторами, исследовать большие массивы и сокращать стоимость передачи данных.')
    add_bullet(doc,'Стенд: Python '+python_version+'; '+platform_name+'; '+str(logical_cpus)+' логических CPU. Это ограниченная Linux-среда, а не Mac пользователя; перед защитой рекомендуется повторить замеры локально.')
    add_callout(doc,'Практический вывод','Параллелизм имеет смысл только тогда, когда объём работы на процесс превышает накладные расходы запуска и обмена данными.', fill='FFF7E6')
    doc.add_paragraph('Фактические данные: results/benchmark_results.csv и results/charts/*.png. ФИО участников и реальный личный вклад заполнить перед сдачей.')
    out=DOCS/'report.docx'; doc.save(out); return out

# ------ PPTX helpers ------
def rgb(hexstr): return PptRGB.from_string(hexstr)

def add_bg(slide, color=NAVY):
    sh=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,PptInches(13.333),PptInches(7.5))
    sh.fill.solid(); sh.fill.fore_color.rgb=rgb(color); sh.line.fill.background()
    slide.shapes._spTree.remove(sh._element); slide.shapes._spTree.insert(2, sh._element)

def add_text(slide, text, x, y, w, h, size=16, color=WHITE, bold=False, font='Aptos', align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, margin=0.03):
    shape=slide.shapes.add_textbox(PptInches(x),PptInches(y),PptInches(w),PptInches(h))
    tf=shape.text_frame; tf.clear(); tf.word_wrap=True
    tf.margin_left=PptInches(margin); tf.margin_right=PptInches(margin); tf.margin_top=PptInches(margin); tf.margin_bottom=PptInches(margin)
    tf.vertical_anchor=valign
    for i,line in enumerate(str(text).split('\n')):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.text=line; p.alignment=align
        p.font.name=font; p.font.size=PptPt(size); p.font.bold=bold; p.font.color.rgb=rgb(color)
        p.space_after=PptPt(4)
    return shape

def add_box(slide,x,y,w,h,fill=PANEL,line=None, radius=True):
    shape=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,PptInches(x),PptInches(y),PptInches(w),PptInches(h))
    shape.fill.solid(); shape.fill.fore_color.rgb=rgb(fill)
    if line:
        shape.line.color.rgb=rgb(line); shape.line.width=PptPt(1)
    else: shape.line.fill.background()
    return shape

def add_line(slide,x1,y1,x2,y2,color=TEAL,width=2):
    sh=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,PptInches(x1),PptInches(y1),PptInches(x2),PptInches(y2))
    sh.line.color.rgb=rgb(color); sh.line.width=PptPt(width); return sh

def add_header(slide, section, title, subtitle=None):
    add_text(slide, section.upper(), 0.7, 0.35, 4.5, 0.28, 10, TEAL, True)
    add_text(slide, title, 0.7, 0.72, 12.0, 0.58, 29, WHITE, True)
    if subtitle: add_text(slide, subtitle, 0.72, 1.34, 11.8, 0.45, 12, MUTED)

def add_footer(slide, n):
    add_line(slide,0.7,7.12,12.65,7.12,PANEL2,1)
    add_text(slide,'PARALLEL SORT LAB  •  АЛГОРИТМЫ И СТРУКТУРЫ ДАННЫХ',0.72,7.18,8,0.17,7,MUTED,True)
    add_text(slide,f'{n:02d} / 12',11.65,7.16,0.95,0.2,8,TEAL,True,align=PP_ALIGN.RIGHT)

def new_slide(prs, section, title, subtitle=None):
    s=prs.slides.add_slide(prs.slide_layouts[6]); add_bg(s); add_header(s,section,title,subtitle); add_footer(s,len(prs.slides)); return s

def add_card(slide,x,y,w,h,title,body,accent=TEAL,kicker=None):
    add_box(slide,x,y,w,h,PANEL,PANEL2)
    add_box(slide,x,y,0.055,h,accent,accent,False)
    if kicker: add_text(slide,kicker.upper(),x+0.22,y+0.17,w-0.42,0.2,8,accent,True)
    add_text(slide,title,x+0.22,y+(0.42 if kicker else 0.23),w-0.42,0.4,16,WHITE,True)
    add_text(slide,body,x+0.22,y+(0.9 if kicker else 0.75),w-0.42,h-(1.02 if kicker else 0.86),11,MUTED)

def add_pill(slide,text,x,y,w,fill=PANEL2,fg=TEAL):
    add_box(slide,x,y,w,0.36,fill,fill)
    add_text(slide,text,x+0.08,y+0.04,w-0.16,0.22,9,fg,True,align=PP_ALIGN.CENTER,valign=MSO_ANCHOR.MIDDLE)

def add_img(slide,path,x,y,w,h=None):
    if not Path(path).exists(): return
    slide.shapes.add_picture(str(path),PptInches(x),PptInches(y),width=PptInches(w),height=PptInches(h) if h else None)

def fmt_ms(size, algo, workers=1):
    return f"{float(get_row(size,algo,workers)['median_time_s'])*1000:.2f}"


def create_pptx():
    prs=Presentation(); prs.slide_width=PptInches(13.333); prs.slide_height=PptInches(7.5)
    prs.core_properties.title='Parallel Sort Lab — Параллельные алгоритмы сортировки'
    prs.core_properties.subject='Учебный проект по алгоритмам и структурам данных'
    prs.core_properties.author='Команда Parallel Sort Lab'
    # 1 cover
    s=prs.slides.add_slide(prs.slide_layouts[6]); add_bg(s)
    add_pill(s,'УЧЕБНЫЙ ПРОЕКТ  /  PYTHON',0.78,0.65,2.5)
    add_text(s,'PARALLEL\nSORT LAB',0.78,1.38,6.4,1.55,38,WHITE,True)
    add_text(s,'Параллельные алгоритмы сортировки',0.82,3.1,6.0,0.52,22,TEAL,True)
    add_text(s,'Сравнение Merge Sort, Parallel Merge Sort и sorted()\nчерез измерения, метрики и визуализацию',0.82,3.8,5.3,0.85,15,MUTED)
    add_pill(s,'5 УЧАСТНИКОВ',0.82,5.26,1.75,fill=PANEL2,fg=WHITE)
    add_pill(s,'TKINTER + MATPLOTLIB',2.72,5.26,2.55,fill=PANEL2,fg=WHITE)
    add_text(s,'Сдача: 29 октября 2026',0.82,6.48,4.0,0.3,10,MUTED)
    add_img(s,FIGURES/'app_results.png',7.0,1.0,5.65,3.45)
    add_box(s,7.0,4.72,5.65,1.3,PANEL,PANEL2)
    add_text(s,'ОТ ИДЕИ — К ИЗМЕРЕНИЮ',7.3,4.98,4.9,0.28,10,TEAL,True)
    add_text(s,'Код • реальные тесты • CSV • графики • отчёт',7.3,5.38,4.95,0.35,13,WHITE,True)
    add_footer(s,1)
    # 2 comparison
    s=new_slide(prs,'01 / постановка','Три реализации — одна задача','Сортируем одинаковые массивы, измеряем полное время вызова и проверяем результат.')
    add_card(s,0.72,2.05,3.85,2.65,'Merge Sort','Собственная последовательная реализация.\nO(n log n) по времени, O(n) дополнительной памяти.',TEAL,'Базовый алгоритм')
    add_card(s,4.73,2.05,3.85,2.65,'Parallel Merge Sort','Разбиение на p частей → отдельные процессы → слияние в главном процессе.',BLUE,'Что исследуем')
    add_card(s,8.74,2.05,3.85,2.65,'sorted()','Встроенный высокооптимизированный Timsort. Практический эталон, а не наш алгоритм.',GOLD,'Ориентир')
    add_box(s,0.72,5.15,11.87,0.8,PANEL)
    add_text(s,'Критерий корректности',0.98,5.35,2.4,0.25,11,TEAL,True)
    add_text(s,'Каждый результат должен совпадать с sorted(исходный_массив)',3.4,5.33,8.6,0.28,14,WHITE,True)
    # 3 merge sort
    s=new_slide(prs,'02 / алгоритм','Merge Sort: «разделяй и властвуй»','Рекурсивное деление уменьшает размер задачи; слияние объединяет отсортированные части линейно.')
    # visual top row arrays
    labels=[('8  3  7  4  9  2  5  1',0.9,2.15,11.4),('8  3  7  4',1.3,3.35,4.65),('9  2  5  1',7.35,3.35,4.65)]
    for text,x,y,w in labels:
        add_box(s,x,y,w,0.62,PANEL,PANEL2)
        add_text(s,text,x+0.1,y+0.15,w-0.2,0.24,18,WHITE,True,align=PP_ALIGN.CENTER)
    add_line(s,6.65,2.78,3.6,3.3,TEAL,2); add_line(s,6.65,2.78,9.65,3.3,TEAL,2)
    add_text(s,'Деление',0.94,4.32,2.0,0.3,12,TEAL,True)
    add_text(s,'до массивов длины 1',0.94,4.7,3.1,0.28,11,MUTED)
    add_text(s,'Рекурсивная сортировка',4.3,4.32,3.6,0.3,12,BLUE,True)
    add_text(s,'каждая половина сортируется отдельно',4.3,4.7,4.1,0.28,11,MUTED)
    add_text(s,'Слияние',9.1,4.32,2.1,0.3,12,GOLD,True)
    add_text(s,'два указателя, линейный проход',9.1,4.7,3.0,0.35,11,MUTED)
    add_box(s,0.9,5.48,11.4,0.65,'0F766E')
    add_text(s,'Результат: 1  2  3  4  5  7  8  9',1.1,5.65,11.0,0.25,18,WHITE,True,align=PP_ALIGN.CENTER)
    # 4 complexity
    s=new_slide(prs,'03 / анализ','Сложность: что гарантирует алгоритм','Скорость растёт не только от алгоритма, но и от реализации и накладных расходов среды.')
    add_card(s,0.78,2.0,3.72,2.05,'O(n log n)','Временная сложность Merge Sort\nв лучшем, среднем и худшем случаях.',TEAL,'Время')
    add_card(s,4.8,2.0,3.72,2.05,'O(n)','Дополнительная память для\nвременных массивов при слиянии.',BLUE,'Память')
    add_card(s,8.82,2.0,3.72,2.05,'Устойчивость','Равные элементы сохраняют порядок,\nесли при равенстве брать левую часть.',GOLD,'Свойство')
    add_box(s,0.78,4.55,11.75,1.22,PANEL,PANEL2)
    add_text(s,'T(n) = 2T(n/2) + O(n)  ⇒  T(n) = O(n log n)',1.06,4.82,11.15,0.4,21,TEAL,True,align=PP_ALIGN.CENTER)
    add_text(s,'Глубина рекурсии — log₂(n); суммарная работа каждого уровня слияния — O(n).',1.06,5.34,11.15,0.25,11,MUTED,align=PP_ALIGN.CENTER)
    # 5 parallel architecture
    s=new_slide(prs,'04 / параллелизм','Как работает Parallel Merge Sort','Один пул процессов на вызов; финальное слияние в текущей реализации остаётся последовательным.')
    steps=[('Входной массив','n элементов',0.75,2.5,2.25),('Разбиение','p частей',3.35,2.5,2.0),('Процессы','сортировка частей',5.85,2.5,2.2),('Слияние','дерево слияния',8.55,2.5,2.0),('Результат','отсортировано',10.95,2.5,1.65)]
    for i,(title,body,x,y,w) in enumerate(steps):
        accent=[TEAL,BLUE,BLUE,GOLD,TEAL][i]
        add_box(s,x,y,w,1.25,PANEL,PANEL2)
        add_text(s,f'{i+1:02d}',x+0.15,y+0.15,0.35,0.22,9,accent,True)
        add_text(s,title,x+0.15,y+0.45,w-0.3,0.28,12,WHITE,True,align=PP_ALIGN.CENTER)
        add_text(s,body,x+0.1,y+0.84,w-0.2,0.25,8.5,MUTED,align=PP_ALIGN.CENTER)
        if i < 4: add_line(s,x+w+0.05,3.12,steps[i+1][2]-0.1,3.12,accent,2)
    for i,(title,body) in enumerate([('Число процессов','workers = 1, 2, 4'),('Порог','малые массивы — последовательно'),('Платформа','spawn совместим с macOS')]):
        x=0.8+i*4.08
        add_card(s,x,4.55,3.72,1.28,title,body,[TEAL,BLUE,GOLD][i])
    # 6 overheads
    s=new_slide(prs,'05 / ограничения','Параллелизм не бесплатный','Процессы дают возможность выполнять вычисления параллельно, но добавляют свои расходы.')
    items=[('01','Запуск','Создание процессов и их жизненный цикл.'),('02','Передача','Сериализация списков и передача данных через pickle.'),('03','Слияние','Объединение частей выполняется в главном процессе.'),('04','Ресурсы','Если процессов больше CPU, они конкурируют за время процессора.')]
    for i,(num,title,body) in enumerate(items):
        x=0.8+(i%2)*6.1; y=2.0+(i//2)*1.65
        add_box(s,x,y,5.7,1.32,PANEL,PANEL2)
        add_text(s,num,x+0.2,y+0.25,0.55,0.35,17,TEAL if i%2==0 else GOLD,True)
        add_text(s,title,x+0.9,y+0.22,4.3,0.3,15,WHITE,True)
        add_text(s,body,x+0.9,y+0.63,4.4,0.47,11,MUTED)
    add_box(s,0.8,5.6,11.9,0.72,'3B2632')
    add_text(s,'По закону Амдала последовательная часть ограничивает теоретический выигрыш: S(p) ≤ 1 / (f + (1 − f) / p).',1.05,5.83,11.4,0.27,12,'FDA4AF',True,align=PP_ALIGN.CENTER)
    # 7 methodology
    s=new_slide(prs,'06 / эксперимент','Сравнение без подмены условий','Одинаковые данные, таймер высокой точности и проверка результата вне измеряемого времени.')
    steps=[('1','Сгенерировать','Массив генерируется один раз.'),('2','Измерить','time.perf_counter(); полный вызов сортировки.'),('3','Повторить','3 повтора, итог — медиана.'),('4','Проверить','Сравнить с sorted(вход).'),('5','Сохранить','CSV + графики времени, S и E.')]
    for i,(num,title,body) in enumerate(steps):
        x=0.75+i*2.52
        add_box(s,x,2.35,2.25,2.15,PANEL,PANEL2)
        add_box(s,x+0.18,2.57,0.5,0.5,TEAL if i in (0,3) else BLUE)
        add_text(s,num,x+0.18,2.69,0.5,0.22,12,NAVY,True,align=PP_ALIGN.CENTER)
        add_text(s,title,x+0.18,3.25,1.9,0.3,13,WHITE,True)
        add_text(s,body,x+0.18,3.7,1.9,0.65,10,MUTED)
    add_box(s,0.75,5.1,11.9,1.08,PANEL2)
    add_text(s,'Параметры',1.0,5.33,1.3,0.25,11,TEAL,True)
    add_text(s,'random: 10k / 50k / 100k  •  sorted/reversed: 50k  •  3 повтора  •  p = 1 / 2 / 4',2.35,5.29,9.8,0.35,11,WHITE,True)
    add_text(s,'S(p) = T_merge / T_parallel(p)     |     E(p) = S(p) / p',2.35,5.73,8.8,0.25,11,MUTED)
    # 8 results table
    s=new_slide(prs,'07 / результаты','Что показали реальные измерения','Медиана времени, миллисекунды. Случайные данные, по 3 повтора; все результаты корректны.')
    table=s.shapes.add_table(4,5,PptInches(0.85),PptInches(2.05),PptInches(11.65),PptInches(2.55)).table
    widths=[1.65,2.25,2.05,2.7,2.6]
    for i,w in enumerate(widths): table.columns[i].width=PptInches(w)
    headers=['Размер n','Merge Sort','sorted()','Parallel p=2','Parallel p=4']
    for j,h in enumerate(headers):
        c=table.cell(0,j); c.text=h; c.fill.solid(); c.fill.fore_color.rgb=rgb(TEAL)
        c.vertical_anchor=MSO_ANCHOR.MIDDLE
        for p in c.text_frame.paragraphs:
            p.alignment=PP_ALIGN.CENTER
            for r in p.runs: r.font.name='Aptos'; r.font.size=PptPt(12); r.font.bold=True; r.font.color.rgb=rgb(NAVY)
    for i,n in enumerate(sizes, start=1):
        vals=[f'{n:,}'.replace(',',' '),fmt_ms(n,'merge'),fmt_ms(n,'builtin'),fmt_ms(n,'parallel',2),fmt_ms(n,'parallel',4)]
        for j,v in enumerate(vals):
            c=table.cell(i,j); c.text=v; c.fill.solid(); c.fill.fore_color.rgb=rgb(PANEL if i%2 else PANEL2); c.vertical_anchor=MSO_ANCHOR.MIDDLE
            for p in c.text_frame.paragraphs:
                p.alignment=PP_ALIGN.CENTER
                for r in p.runs: r.font.name='Aptos'; r.font.size=PptPt(13 if j==0 else 12); r.font.bold=(j==0); r.font.color.rgb=rgb(WHITE)
    add_card(s,0.9,5.0,5.55,1.1,'sorted() — быстрее','При n = 100 000: 13,35 мс против 172,23 мс у Merge Sort.',GOLD)
    add_card(s,6.78,5.0,5.55,1.1,'p = 2 и p = 4 — медленнее','Пул и передача данных доминируют на протестированных размерах.',RED)
    # 9 times chart
    s=new_slide(prs,'08 / визуализация','Время выполнения','Логарифмическая шкала показывает разницу между Python Merge Sort, sorted() и вызовом с новыми процессами.')
    add_img(s,RESULTS/'charts'/'time_random.png',0.9,1.9,11.5,4.8)
    # 10 speedup + efficiency
    s=new_slide(prs,'09 / масштабирование','Ускорение оказалось меньше единицы','Для данной реализации и стенда дополнительный процесс увеличил полное время вызова.')
    add_img(s,RESULTS/'charts'/'speedup_random.png',0.75,1.9,5.95,4.1)
    add_img(s,RESULTS/'charts'/'efficiency_random.png',6.8,1.9,5.75,4.1)
    add_box(s,0.9,6.16,11.5,0.52,'3B2632')
    r2=float(get_row(100000,'parallel',2)['speedup']); e2=float(get_row(100000,'parallel',2)['efficiency'])
    add_text(s,f"При n = 100 000: S(2) = {r2:.3f}; E(2) = {e2:.3f}. Значение S < 1 означает замедление, а не ускорение.",1.1,6.31,11.1,0.23,11,'FDA4AF',True,align=PP_ALIGN.CENTER)
    # 11 application & verification
    s=new_slide(prs,'10 / программная реализация','Интерфейс, тесты и экспорт','Пользователь задаёт эксперимент, видит таблицу, графики, корректность и выгружает CSV.')
    add_img(s,FIGURES/'app_results.png',0.8,1.85,7.45,4.55)
    add_card(s,8.55,1.98,3.9,1.05,'32 автоматических теста','Запуск через unittest; GUI проверен в виртуальном дисплее.',TEAL)
    add_card(s,8.55,3.23,3.9,1.05,'CSV + 3 графика','Время, ускорение и эффективность; графики строятся из измерений.',BLUE)
    add_card(s,8.55,4.48,3.9,1.05,'25 строк результатов','Все выходы эксперимента совпали с sorted(вход).',GOLD)
    add_text(s,'Запуск: python3 main.py     •     Тесты: xvfb-run -a python3 -m unittest discover -s tests -v (Linux)',0.92,6.45,11.5,0.28,10,MUTED)
    # 12 conclusions
    s=new_slide(prs,'11 / выводы','Результат исследования','Параллельность — инструмент, а не гарантия ускорения.')
    add_card(s,0.8,2.0,3.75,1.72,'1. Корректность','Все 25 итоговых строк CSV содержат корректный результат.',TEAL)
    add_card(s,4.8,2.0,3.75,1.72,'2. Накладные расходы','Запуск процессов и передача данных доминировали на выбранных размерах.',RED)
    add_card(s,8.8,2.0,3.75,1.72,'3. Следующий шаг','Повторить тест на целевом Mac и изучить переиспользование пула.',GOLD)
    add_box(s,0.8,4.25,11.75,1.45,PANEL,PANEL2)
    add_text(s,'РАСПРЕДЕЛЕНИЕ ЗАЩИТЫ',1.05,4.5,3.0,0.24,10,TEAL,True)
    add_text(s,'У1 — архитектура и итоги   •   У2 — теория и Merge Sort   •   У3 — параллельная версия\nУ4 — эксперимент и графики   •   У5 — интерфейс, тесты и инструкция запуска',1.05,4.92,10.9,0.52,12,WHITE)
    add_text(s,'Ограничение: измерения выполнены в Linux-среде с 3 логическими CPU; значения зависят от оборудования.',1.03,6.04,11.3,0.32,10,MUTED)
    out=DOCS/'presentation.pptx'; prs.save(out); return out


def create_report_md():
    type_table_rows=[]
    for kind,label in [('random','Случайные'),('sorted','Отсортированные'),('reversed','Обратный порядок')]:
        type_table_rows.append('| '+label+' | '+f"{float(get_row(50000,'merge',1,kind)['median_time_s'])*1000:.2f}"+' | '+f"{float(get_row(50000,'builtin',1,kind)['median_time_s'])*1000:.3f}"+' | '+f"{float(get_row(50000,'parallel',2,kind)['median_time_s'])*1000:.1f}"+' | '+f"{float(get_row(50000,'parallel',4,kind)['median_time_s'])*1000:.1f}"+' |')
    type_table='\n'.join(type_table_rows)
    result_table=[]
    for n in sizes:
        row=[f'{n:,}'.replace(',',' ')]
        for alg,workers in [('merge',1),('builtin',1),('parallel',1),('parallel',2),('parallel',4)]:
            row.append(f"{float(get_row(n,alg,workers)['median_time_s'])*1000:.2f}")
        result_table.append('| '+' | '.join(row)+' |')
    header='| n | Merge Sort, мс | sorted(), мс | Parallel p=1, мс | Parallel p=2, мс | Parallel p=4, мс |\n|---:|---:|---:|---:|---:|---:|'
    md=f'''# Parallel Sort Lab — отчёт по проекту

**Дисциплина:** Алгоритмы и структуры данных  
**Тема:** Параллельные алгоритмы сортировки  
**Команда:** 5 участников; ФИО и реальный вклад необходимо заполнить перед сдачей.

## 1. Цель и задачи

Цель — экспериментально сравнить последовательную и параллельную сортировки слиянием и встроенную `sorted()`, проверить корректность, рассчитать ускорение и эффективность, визуализировать измерения.

## 2. Теория и алгоритмы

Merge Sort работает по принципу «разделяй и властвуй»: массив делится на две части, каждая сортируется рекурсивно, затем части сливаются за линейное время. Рекуррентность `T(n) = 2T(n/2) + O(n)` даёт время `O(n log n)`; дополнительная память — `O(n)`. Алгоритм устойчив при выборе элемента левой части в случае равенства.

Parallel Merge Sort делит входной список на `p` частей, сортирует части через `ProcessPoolExecutor` и объединяет их в главном процессе. Пул создаётся один раз на вызов; небольшие массивы ниже порога сортируются последовательно. Время включает создание пула, сериализацию/передачу частей и финальное слияние. `sorted()` — встроенная реализация Timsort, используемая как практический ориентир, а не как собственный алгоритм.

## 3. Анализ сложности

| Алгоритм | Время | Память |
|---|---|---|
| Merge Sort | `O(n log n)` | `O(n)` |
| Parallel Merge Sort | `O((n/p) log(n/p))` для сортировки частей + `O(n log p)` на последовательное финальное слияние + накладные расходы | `O(n)` плюс копии частей |
| sorted() / Timsort | `O(n log n)` в общем случае; на упорядоченных данных возможно `O(n)` | `O(n)` |

## 4. Реализация

Интерфейс написан на Tkinter, графики — Matplotlib. `benchmark/` отвечает за генерацию данных, замеры, метрики и CSV. `algorithms/` содержит две сортировки. В GUI вычисления выполняются в фоновом потоке; сообщения передаются главному потоку через очередь. Запуск защищён `if __name__ == "__main__"` для корректной работы `multiprocessing`.

## 5. Методика эксперимента

Полный размерный набор (10 000, 50 000, 100 000) измерен на случайных данных; для проверки влияния порядка на n = 50 000 дополнительно измерены sorted и reversed. Везде 3 повтора; для параллельной реализации `p = 1, 2, 4`. Генерация и проверка не входят в измеряемое время. Таймер — `time.perf_counter()`, итог — медиана трёх замеров. Каждый результат сравнивается с `sorted(вход)`.

Ускорение: `S(p) = T_merge / T_parallel(p)`. Эффективность: `E(p) = S(p) / p`.

## 6. Результаты

{header}
{chr(10).join(result_table)}

Всего выполнено 75 измеряемых вызовов, объединённых в 25 строк CSV; все строки имеют `correct = True`. Диаграммы сохранены в `results/charts/`; фактические замеры находятся в `results/benchmark_results.csv`.

## 7. Дополнительное сравнение порядка данных при n = 50 000

| Тип данных | Merge Sort, мс | sorted(), мс | Parallel p=2, мс | Parallel p=4, мс |
|---|---:|---:|---:|---:|
{type_table}

## 8. Анализ результатов

В данном окружении Parallel Merge Sort с `p = 2` и `p = 4` оказался медленнее последовательного Merge Sort на всех трёх размерах. Например, при `n = 100 000` последовательный Merge Sort занял {float(get_row(100000,'merge')['median_time_s'])*1000:.2f} мс, а параллельная версия — {float(get_row(100000,'parallel',2)['median_time_s'])*1000:.1f} мс при `p = 2` и {float(get_row(100000,'parallel',4)['median_time_s'])*1000:.1f} мс при `p = 4`. Здесь ускорение меньше единицы. Причина — расходы запуска новых процессов, передачи данных и последовательного слияния; в среде доступно только {logical_cpus} логических CPU, а `p = 4` превышает это число.

Встроенная `sorted()` существенно быстрее Python-реализации Merge Sort. При `n = 100 000` её медиана составила {float(get_row(100000,'builtin')['median_time_s'])*1000:.2f} мс против {float(get_row(100000,'merge')['median_time_s'])*1000:.2f} мс. Это сравнение реализаций (Timsort на C против алгоритма на Python), а не чистое сравнение алгоритмических идей.

## 9. Выводы и ограничения

Проект реализует алгоритмы, GUI, CSV-экспорт, три вида графиков и автоматические тесты. Корректность всех измеренных результатов подтверждена. Эксперимент демонстрирует, что параллелизм не гарантирует ускорение: при выбранном размере задачи накладные расходы выше выигрыша. В качестве дальнейшего улучшения можно переиспользовать пул между повторениями, оценить более крупные массивы и повторить исследование на целевом компьютере.

**Стенд:** Python {python_version}; {platform_name}; {logical_cpus} логических CPU. Это ограниченная Linux-среда выполнения, а не Mac пользователя. Перед защитой желательно повторить замеры на компьютере команды и при необходимости обновить CSV, графики, выводы и презентацию.
'''
    (DOCS/'report.md').write_text(md,encoding='utf-8')


def create_presentation_plan():
    content='''# План презентации и заметки докладчиков — Parallel Sort Lab

## Распределение времени

Презентация рассчитана примерно на 10–12 минут. Каждый из пяти участников представляет свою часть не менее одной минуты.

1. **Участник 1 — архитектура и итоги (слайды 1–2, 11–12).** Представляет задачу, объясняет структуру проекта, показывает тестирование и выводы.
2. **Участник 2 — Merge Sort (слайды 3–4).** Объясняет принцип «разделяй и властвуй», слияние и сложность `O(n log n)`.
3. **Участник 3 — параллельная версия (слайды 5–6).** Рассказывает о разбиении на процессы, передаче данных, накладных расходах и законе Амдала.
4. **Участник 4 — эксперимент (слайды 7–10).** Объясняет методику, таблицу, графики времени, ускорения и эффективности.
5. **Участник 5 — интерфейс и документация (слайды 10–12).** Демонстрирует GUI, экспорт CSV и инструкцию запуска; озвучивает ограничения.

## Ключевые заметки докладчика

### Слайды 1–2: постановка и реализация

Наша цель — не объявить параллельную сортировку самой быстрой заранее, а проверить это на измерениях. Сравниваем собственную последовательную сортировку слиянием, параллельную реализацию через процессы и встроенную `sorted()`. Важное различие: `sorted()` — практический ориентир, а не наша реализация.

### Слайды 3–4: Merge Sort

Массив рекурсивно делится до единичных частей. Слияние двух упорядоченных частей линейно. Глубина деления равна `log₂ n`, а работа на каждом уровне пропорциональна n; поэтому время `O(n log n)`. Для временных массивов требуется `O(n)` памяти. Merge Sort устойчив при корректном выборе равных элементов.

### Слайды 5–6: параллельная версия

Список делится на p частей, части обрабатываются процессами, после чего результаты попарно сливаются в главном процессе. Процессы подходят для вычислительной нагрузки, ограниченной GIL. Но создание пула и сериализация данных имеют цену. Поэтому параллелизм может замедлить небольшой эксперимент; это не обязательно ошибка реализации.

### Слайды 7–10: методика и реальные числа

Полный размерный набор 10 000, 50 000 и 100 000 измерен на случайных данных; на размере 50 000 дополнительно сравнили sorted и reversed. Для каждого тестового случая использовали три повтора и медиану. Генерация и проверка исключены из таймера; вызов параллельной сортировки включает запуск пула. Каждый результат сравнивается с `sorted(вход)`. Дополнительно сравнили случайные, отсортированные и обратные данные при n=50 000. В основном стенде при n=100 000 Merge Sort занимает около 172,23 мс; `sorted()` — 13,35 мс; Parallel Merge Sort p=2 — 886,07 мс; p=4 — 1 611,87 мс. Следовательно, ускорения не получено: S(2)≈0,194, S(4)≈0,107.

### Слайды 11–12: GUI и вывод

Приложение позволяет выбрать размер данных, тип, алгоритм, число процессов и повторов; результаты появляются в таблице, графики строятся по фактическим измерениям, CSV можно экспортировать. Автоматические тесты проверяют алгоритмы, метрики, CSV и GUI-сценарий. Измерения выполнены в Linux-контейнере с тремя логическими CPU, поэтому числа нельзя прямо переносить на Mac команды; стоит повторить тест локально.

## Вероятные вопросы преподавателя

1. **Почему Merge Sort имеет `O(n log n)`?** Глубина рекурсии `log₂ n`, работа на каждом уровне — `O(n)`.
2. **Почему параллельная версия медленнее?** Запуск процессов, pickle/передача списков и последовательное слияние перекрывают выгоду на тестированных размерах.
3. **Почему процессы, а не потоки?** Из-за GIL в CPython CPU-bound Python-код обычно не ускоряется потоками.
4. **Что показывает S(p)?** `T_merge / T_parallel(p)`. Если S меньше 1, параллельный вызов медленнее базового.
5. **Что показывает E(p)?** `S(p)/p`; при идеальном масштабировании значение близко к 1.
6. **Почему сравниваем с sorted()?** Как практический ориентир оптимизированной встроенной реализации. Это не прямое сравнение одинакового кода.
7. **Почему используем медиану?** Она меньше чувствительна к случайным выбросам измерений.
8. **Что улучшить?** Переиспользовать пул между повторами, исследовать более крупные n, уменьшить передачу данных или использовать C/NumPy.
'''
    (DOCS/'presentation_plan.md').write_text(content,encoding='utf-8')


def create_speaker_notes():
    text='''# Краткий сценарий защиты (5 участников)

**Участник 1 — 1–1,5 минуты.** «Наш проект Parallel Sort Lab сравнивает последовательный Merge Sort, параллельный Merge Sort и встроенную sorted(). Мы реализовали графическое приложение на Tkinter, измерения через perf_counter, проверку корректности, экспорт CSV и построение графиков. Архитектура разделена на алгоритмы, экспериментальный модуль и интерфейс».

**Участник 2 — 1–1,5 минуты.** «Merge Sort реализует принцип “разделяй и властвуй”: деление, рекурсивная сортировка, слияние. Слияние линейно, глубина рекурсии log₂ n, поэтому время O(n log n), дополнительная память O(n). Алгоритм устойчив при выборе левого элемента при равенстве».

**Участник 3 — 1–1,5 минуты.** «Параллельный вариант делит массив на p частей и сортирует части через ProcessPoolExecutor. После этого главный процесс объединяет части попарно. Мы используем процессы для CPU-bound работы, но запуск пула и сериализация списков добавляют расходы. Поэтому ускорение не гарантируется».

**Участник 4 — 1,5–2 минуты.** «Мы тестировали n=10 000, 50 000 и 100 000, случайные данные, три повтора и брали медиану. Генерация и проверка не входят в измерение, а время параллельного вызова включает пул. При n=100 000 Merge Sort — около 172 мс; sorted() — около 13 мс; параллельный p=2 — около 886 мс, p=4 — около 1 612 мс. Это показывает, что на данном стенде накладные расходы доминируют».

**Участник 5 — 1–1,5 минуты.** «В интерфейсе можно настроить эксперимент, посмотреть таблицу результатов и три графика, экспортировать CSV. Тесты проверяют корректность сортировки, обработку граничных случаев, метрики и интерфейс. Эксперимент выполнен в ограниченной Linux-среде с тремя логическими CPU, поэтому перед переносом выводов на другой компьютер нужно повторить замеры».

**Важно перед защитой:** впишите реальные ФИО, убедитесь, что каждый участник умеет объяснить свою часть, и при возможности повторите benchmark на компьютере команды. Не утверждайте, что параллельная сортировка ускорила выполнение: в сохранённом эксперименте S(p) меньше 1 для p=2 и p=4.
'''
    (DOCS/'speaker_notes.md').write_text(text,encoding='utf-8')

if __name__ == '__main__':
    regenerate_charts(); create_input_type_chart()
    create_report_md(); create_presentation_plan(); create_speaker_notes()
    report=create_report_docx(); pptx=create_pptx()
    print('Created', report)
    print('Created', pptx)
