"""Rebuild the preliminary workbook and brief; prices are unverified allowances.
Requires Python 3, xlsxwriter, reportlab, Pillow. Run in this directory.
"""
from pathlib import Path
from collections import defaultdict
import json, math, csv
import xlsxwriter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

P = Path(__file__).resolve().parent
data = json.loads((P/'estimate_data.json').read_text())
rows = data['rows']
groups = defaultdict(float)
for r in rows: groups[r['group']] += r['qty'] * r['price']
subtotal = sum(groups.values())
reserve = data['reserve']
targets = [2050000, 480000, 620000, 500000, 500000, 200000, 450000, 300000, 1050000, 150000]
assert sum(targets) + reserve == data['budget']
totals = {key: sum(r['qty']*r[key] for r in rows) for key in ('low','price','high')}

inputs = [
 ('budget','Лимит с резервом',7000000,'₸','Без крупной техники'),
 ('reserve','Резерв',700000,'₸','10% лимита; не 10% любой итоговой суммы'),
 ('work_area','Площадь для оплаты работы',41,'м²','Допущение договора; полезная 38,8, учётная 40,7'),
 ('rate','Ставка бригады',50000,'₸/м²','От пользователя; материалы считаются отдельно'),
 ('living','Жилая комната',20.5,'м²','По экспликации'),
 ('kitchen','Кухня',8.9,'м²','По экспликации, остаётся на месте'),
 ('hall','Прихожая',4.3,'м²','По экспликации'),
 ('bath','Санузел',5.1,'м²','По экспликации'),
 ('wall_paint','Стены под покраску, нетто',96,'м²','Предварительно; проверить проёмы/развёртки'),
 ('dry_plaster','Сухие зоны под штукатурку',86,'м²','Неизвестно, насколько нужна новая штукатурка'),
 ('wet_plaster','Влажные зоны под штукатурку',25,'м²','Предварительная площадь'),
 ('bath_tiles','Облицовка стен ванной',23,'м²','Предварительная площадь'),
 ('backsplash','Кухонный фартук',2.5,'м²','Около 3 м кухни'),
 ('height','Условная чистая высота',2.8,'м','После отделки; требуется обмер'),
 ('sound_wall_length','Длина условной соседской стены',3.1,'м','Положение межквартирной стены не подтверждено'),
 ('floor_mm','Средний финишный ровнитель',5,'мм','На существующей пригодной базовой стяжке'),
 ('plaster_mm','Гипсовая штукатурка, средний слой',10,'мм','По обмеру отклонений'),
 ('wet_plaster_mm','Цементная штукатурка, средний слой',10,'мм','По обмеру отклонений'),
 ('waste','Запас материала',0.10,'доля','Уточнить по раскрою, системе и упаковкам'),
 ('paint_coverage','Укрывистость/площадь покрытия стены',10,'м²/л/слой','Условный расход, заменить паспортом'),
 ('paint_coats','Число слоёв краски',2,'слоя','По технологии продукта'),
 ('ceiling_coverage','Площадь покрытия потолка',8,'м²/л/слой','Условный расход'),
 ('waterproof_area','Площадь гидроизоляции',16,'м²','Пол + подъёмы + мокрые стены; проверить схему'),
]
cell = {key:f"'Исходные'!$B${i+5}" for i,(key,*_) in enumerate(inputs)}
vals = {key:value for key,_,value,*_ in inputs}
F = lambda key: cell[key]
area = '('+'+'.join(F(k) for k in ('living','kitchen','hall','bath'))+')'
waste = '(1+'+F('waste')+')'
qform = {
 0: '='+F('work_area'),
 1: '=ROUNDUP('+F('dry_plaster')+'*8.5*'+F('plaster_mm')+'/10*'+waste+'/30,0)',
 2: '=ROUNDUP('+F('wet_plaster')+'*12*'+F('wet_plaster_mm')+'/10*'+waste+'/25,0)',
 3: '=ROUNDUP(('+F('wall_paint')+'+'+area+')*2*'+waste+'/20,0)',
 5: '=ROUNDUP('+F('wall_paint')+'*'+waste+'/50,0)',
 7: '=CEILING('+F('wall_paint')+'*'+F('paint_coats')+'/'+F('paint_coverage')+'*'+waste+',10)',
 8: '=CEILING('+area+'*'+F('paint_coats')+'/'+F('ceiling_coverage')+'*'+waste+',5)',
 10: '=ROUNDUP('+area+'*'+F('floor_mm')+'*1.6*'+waste+'/25,0)',
 11: '=ROUNDUP('+F('living')+'*'+waste+',0)',
 12: '=ROUNDUP('+F('living')+'*'+waste+',0)',
 14: '=ROUNDUP(('+F('kitchen')+'+'+F('hall')+'+'+F('bath')+')*'+waste+',0)',
 15: '=ROUNDUP('+F('bath_tiles')+'*'+waste+',0)',
 16: '=ROUNDUP('+F('backsplash')+'*'+waste+',0)',
 17: '=ROUNDUP(('+F('kitchen')+'+'+F('hall')+'+'+F('bath')+'+'+F('bath_tiles')+'+'+F('backsplash')+')*4.5*'+waste+'/25,0)',
 19: '=CEILING('+F('waterproof_area')+'*1.5*'+waste+',10)',
 22: '=ROUNDUP('+F('living')+'*'+waste+',0)',
 23: '=ROUNDUP(2*'+F('living')+'*'+waste+',0)',
 25: '=ROUNDUP('+F('sound_wall_length')+'*'+F('height')+',0)',
 26: '='+area+'-'+F('living'),
}

wb = xlsxwriter.Workbook(P/'Renovation_estimate.xlsx')
wb.set_properties({'title':'Санды-кала 2: материалы и смета','subject':'Предварительные объёмы; товарные цены не проверены','author':'Design concept','comments':'13.10.2026; prices are planning assumptions, not verified retail quotes.'})
wb.set_calc_mode('auto')
fmt = {
 'title':wb.add_format({'bold':True,'font_size':18,'font_color':'#39432E'}),
 'head':wb.add_format({'bold':True,'bg_color':'#626B4E','font_color':'white','text_wrap':True,'valign':'vcenter','border':1}),
 'note':wb.add_format({'font_color':'#6A6A60','text_wrap':True,'valign':'top'}),
 'txt':wb.add_format({'text_wrap':True,'valign':'top','border':1,'border_color':'#DDDED4'}),
 'input':wb.add_format({'font_color':'#176AB1','num_format':'#,##0.##','bg_color':'#F2F8FF','border':1,'border_color':'#DDDED4','valign':'top'}),
 'number':wb.add_format({'num_format':'#,##0.##','border':1,'border_color':'#DDDED4','valign':'top'}),
 'money':wb.add_format({'num_format':'#,##0 "₸"','border':1,'border_color':'#DDDED4','valign':'top'}),
 'total':wb.add_format({'num_format':'#,##0 "₸"','bold':True,'bg_color':'#E9EEDF','border':1}),
 'link':wb.add_format({'font_color':'#176AB1','underline':1,'text_wrap':True,'border':1}),
 'pct':wb.add_format({'font_color':'#176AB1','num_format':'0%','bg_color':'#F2F8FF','border':1}),
}
def sheet(name,title,lastcol,subtitle):
 s=wb.add_worksheet(name);s.hide_gridlines(2);s.freeze_panes(4,0)
 s.merge_range(0,0,0,lastcol,title,fmt['title']);s.merge_range(1,0,2,lastcol,subtitle,fmt['note']);s.set_row(1,23);s.set_row(2,22)
 s.set_landscape();s.set_paper(9);s.fit_to_pages(1,0);s.repeat_rows(0,3);s.set_margins(.3,.3,.4,.4)
 s.set_footer('&LСанды-кала 2 · 13.10.2026&CПредварительно&RСтраница &P из &N')
 return s

s=sheet('Исходные','Исходные данные и допущения',3,'Синие ячейки редактируются. Товарные цены — допущения, не найденные прайсы. Изменение исходных величин пересчитывает связанные строки; фиксированные метражи трасс уточняются вручную.')
s.write_row(3,0,['Параметр','Значение','Единица','Основание/ограничение'],fmt['head'])
s.set_column(0,0,42);s.set_column(1,1,18);s.set_column(2,2,18);s.set_column(3,3,70)
for i,(key,label,value,unit,note) in enumerate(inputs,4):
 s.write(i,0,label,fmt['txt']);s.write_number(i,1,value,fmt['pct'] if key=='waste' else fmt['input']);s.write(i,2,unit,fmt['txt']);s.write(i,3,note,fmt['txt']);s.set_row(i,32)
 s.data_validation(i,1,i,1,{'validate':'decimal','criteria':'>=','value':0,'input_title':'Неотрицательное число','input_message':'Изменяйте только после обмера/договора.'})

s=sheet('Покупки','Предварительная ведомость материалов и работ',12,'Все товарные ставки — расчётные ориентиры. Низкая/базовая/высокая цена — чувствительность, не рыночные котировки. Синие ячейки редактируются. Монтажные работы повторно не добавлены к цене бригады.')
headers=['№','Группа','Материал / статья','Количество','Единица','Низкая цена','Базовая цена','Высокая цена','Низкая сумма','Базовая сумма','Высокая сумма','Расчёт / комплектация','Статус цены']
s.write_row(3,0,headers,fmt['head']);s.set_row(3,34);s.set_column(0,0,4);s.set_column(1,1,22);s.set_column(2,2,48);s.set_column(3,4,12);s.set_column(5,10,17);s.set_column(11,11,70);s.set_column(12,12,30)
s.freeze_panes(4,3);s.autofilter(3,0,3+len(rows),12)
for i,r in enumerate(rows):
 j=i+4;n=j+1;s.set_row(j,60)
 s.write_number(j,0,i+1,fmt['number']);s.write(j,1,r['group'],fmt['txt']);s.write(j,2,r['item'],fmt['txt'])
 if i in qform:s.write_formula(j,3,qform[i],fmt['number'],r['qty'])
 else:s.write_number(j,3,r['qty'],fmt['input'])
 s.write(j,4,r['unit'],fmt['txt'])
 for col,key in [(5,'low'),(6,'price'),(7,'high')]:
  if i==0:s.write_formula(j,col,'='+F('rate'),fmt['money'],r[key])
  else:s.write_number(j,col,r[key],fmt['input'])
 for col,key,pricecol in [(8,'low','F'),(9,'price','G'),(10,'high','H')]:s.write_formula(j,col,f'=D{n}*{pricecol}{n}',fmt['money'],r['qty']*r[key])
 s.write(j,11,r['note'],fmt['txt']);s.write(j,12,r['status'],fmt['txt'])
end=4+len(rows);s.write(end,2,'Без резерва',fmt['head'])
for col,key,letter in [(8,'low','I'),(9,'price','J'),(10,'high','K')]:s.write_formula(end,col,f'=SUM({letter}5:{letter}{end})',fmt['total'],totals[key])
s.print_area(0,0,end,12)

s=sheet('Итог','Смета: реальный расчёт и целевой лимит',4,'База: кухня на месте, ванна, ламинат, обычная кровать, потолок гостиной и одна условная стена с шумоизоляцией, две двери 2600 мм. Цели 7 млн — лимиты для запроса КП, а не подтверждённая корзина.')
s.write_row(3,0,['Статья','Низкий сценарий','Базовый сценарий','Высокий сценарий','Цель 7 млн'],fmt['head']);s.set_column(0,0,42);s.set_column(1,4,21)
for i,((group,value),target) in enumerate(zip(groups.items(),targets),4):
 s.write(i,0,group,fmt['txt']);s.set_row(i,32)
 for col,key,letter in [(1,'low','I'),(2,'price','J'),(3,'high','K')]:
  cached=sum(r['qty']*r[key] for r in rows if r['group']==group)
  s.write_formula(i,col,f'=SUMIF(\'Покупки\'!$B$5:$B${end},A{i+1},\'Покупки\'!${letter}$5:${letter}${end})',fmt['money'],cached)
 s.write_number(i,4,target,fmt['input'])
j=4+len(groups);s.write(j,0,'Расходы без резерва',fmt['head'])
for col,key,letter in [(1,'low','B'),(2,'price','C'),(3,'high','D'),(4,None,'E')]:s.write_formula(j,col,f'=SUM({letter}5:{letter}{j})',fmt['total'],totals[key] if key else sum(targets))
s.write(j+1,0,'Резерв',fmt['txt'])
for col in range(1,5):s.write_formula(j+1,col,'='+F('reserve'),fmt['money'],reserve)
s.write(j+2,0,'Всего с резервом',fmt['head'])
for col,key,letter in [(1,'low','B'),(2,'price','C'),(3,'high','D'),(4,None,'E')]:s.write_formula(j+2,col,f'={letter}{j+1}+{letter}{j+2}',fmt['total'],(totals[key] if key else sum(targets))+reserve)
s.write(j+3,0,'Остаток лимита (минус = превышение)',fmt['txt'])
for col,key,letter in [(1,'low','B'),(2,'price','C'),(3,'high','D'),(4,None,'E')]:s.write_formula(j+3,col,f'={F("budget")}-{letter}{j+3}',fmt['total'],data['budget']-(totals[key] if key else sum(targets))-reserve)
s.conditional_format(j+3,1,j+3,4,{'type':'cell','criteria':'<','value':0,'format':wb.add_format({'font_color':'#A63E2C','bg_color':'#FFF1EA'})})
s.merge_range(j+5,0,j+7,4,'Не включены: крупная техника, лоджия, перенос кухни/стояков, проёмы, стеклянная перепланировка, шкаф-кровать, паркет, новая входная дверь/окна/радиаторы, кондиционер и тёплый пол. Сложная новая вентиляция и неподтверждённые спецработы требуют отдельного КП. Финишный ровнитель предполагает пригодную базовую стяжку.',fmt['note'])

options=[
 ('Инженерная доска вместо ламината',436500,'23 × (22 000 − 6 500) + 80 000; цены условные, выбранная технология уточняется.'),
 ('Шкаф-кровать вместо обычного корпуса',550000,'650 000 − 100 000; матрас остаётся. Не вычитать автоматически обычный шкаф.'),
 ('Третья дверь 2600 мм',190000,'Если нужна дверь в кухню; лимит комплекта, труд по договору.'),
 ('Два новых водосчётчика при необходимости',24000,'2 × 12 000; в базовой смете не покупаются. Оформление и пломбировка не оценены.'),
 ('Каждый дополнительный м² акустической стены',12000,'Материалы условной системы; площадь определить после осмотра, труд проверить по договору.'),
 ('Базовая стяжка 50 мм, если её нет',249900,'Пример: 38,8 × 50 × 1,8 кг/м²/мм × 1,05 / 25 = 147 мешков × 1 700. Проверить конструкцию/нагрузку/рецептуру; не монтажное задание.'),
 ('Оплата работы по 38,8 вместо 41 м²',-110000,'(38,8 − 41) × 50 000; только если это согласованная договорная площадь.'),
 ('Частичный проём кухни и стекло',None,'Не включено; тип стены неизвестен. Сначала конструктивная проверка, затем отдельная спецификация и КП; цену не подтверждаю.'),
]
s=sheet('Опции','Изменения к базовому сценарию',2,'Эти строки не включены в итог и не суммируются автоматически. Положительные суммы — условные доплаты; непроверенная перепланировка не имеет подтверждённой цены.')
s.write_row(3,0,['Опция','Изменение, ₸','Основание / условие'],fmt['head']);s.set_column(0,0,48);s.set_column(1,1,20);s.set_column(2,2,100)
for i,(a,b,c) in enumerate(options,4):
 s.write(i,0,a,fmt['txt']);s.write(i,1,b if b is not None else 'Нужно КП',fmt['money'] if b is not None else fmt['txt']);s.write(i,2,c,fmt['txt']);s.set_row(i,60)

electrical=[
 ('Освещение','70 м: 3×1,5','Часто до 10 А','Уточнить трассы, нагрузки, прокладку и светильники'),
 ('Розетки + духовка + отдельные машины','180 м: 3×2,5','Обычно 16 А','Метры всех этих линий уже в одной строке; духовка повторно не добавляется'),
 ('Варочная панель 7,2 кВт','20 м: 3×6','32 А как кандидат','7 200 / 230 ≈ 31,3 А; проверить допустимый ток при конкретной прокладке'),
 ('Духовка 3,5 кВт','Входит в 180 м 3×2,5','16 А при подтверждении','3 500 / 230 ≈ 15,2 А; по паспорту устройства'),
 ('Совместная мощность панели и духовки','10,7 кВт','≈46,5 А при 230 В','Если ввод 32 А, нужна другая согласованная схема/ограничение мощности'),
 ('3×4 мм²','Не закупать по умолчанию','По расчёту','Если потребуется, заменяет соответствующий метраж 3×2,5'),
 ('Защитное отключение','Тип A, 30 мА — предмет проекта','По схеме/паспорту','Инверторным устройствам могут требоваться другие типы. Проверить PE, N и ввод.'),
 ('Счётчики','Проверить комплект застройщика','Без автоматической замены','Паспорта, поверка, номера, пломбы и порядок ввода; электрический/тепловой прибор не оценены'),
]
s=sheet('Инженерия','Задание электрику и сантехнику',3,'Предварительно 230 В, одна фаза. Это не монтажная схема. Сечения и защиту назначают по расчёту, способу прокладки, вводу и паспортам. Подтвердить выделенную мощность до покупки панели/духовки.')
s.write_row(3,0,['Линия / вопрос','Метраж / параметр','Защита / нагрузка','Проверка'],fmt['head']);s.set_column(0,0,40);s.set_column(1,2,28);s.set_column(3,3,85)
for i,r in enumerate(electrical,4):s.write_row(i,0,r,fmt['txt']);s.set_row(i,65)

sources=[
 ('Ставка работы','Сообщение пользователя','50 000 ₸/м²; состав основных работ подтверждён','Подтверждено пользователем'),
 ('Площади','Ранее предоставленная схема','20,5 + 8,9 + 4,3 + 5,1 = 38,8 м²','Экспликация; натурный обмер не выполнен'),
 ('Товары: все цены','Расчётные допущения','Ни один товарный прайс не прочитан; объёмы предварительные','Не проверено'),
 ('Каталог для запроса','https://12.kz/','CONNECT HTTP 403, повторная проверка без результата','Цена/наличие не получены'),
 ('Каталог для запроса','https://akvilon.kz/','CONNECT HTTP 403','Доставку в Астану проверить'),
 ('Каталог для запроса','https://leroymerlin.kz/','CONNECT HTTP 403','Актуальный ассортимент/условия не установлены'),
 ('Продавцы для сравнения','https://satu.kz/','CONNECT HTTP 403','Отдельные продавцы и документы не проверены'),
 ('Поиск','Google / Bing','CONNECT HTTP 403','Рыночное сравнение не выполнено'),
]
s=sheet('Источники','Источники и запрос коммерческих предложений',3,'13.10.2026. Отсутствие доступа не заменяется выдуманными прайсами. Адреса магазинов, остатки и минимальные цены не установлены. Примеры марок в записке — кандидаты для сравнения, не проверенные предложения.')
s.write_row(3,0,['Источник','Ссылка / происхождение','Что получено','Статус'],fmt['head']);s.set_column(0,0,30);s.set_column(1,1,45);s.set_column(2,2,65);s.set_column(3,3,46)
for i,r in enumerate(sources,4):
 for j,v in enumerate(r):
  if v.startswith('https://'):s.write_url(i,j,v,fmt['link'])
  else:s.write(i,j,v,fmt['txt'])
 s.set_row(i,50)
s.merge_range(14,0,17,3,'Текст запроса: «Прошу цену по приложенной ведомости для доставки в Астану. Укажите производителя/артикул, единицу и упаковку, фактическое количество, цену с налогами, наличие, срок поставки, доставку/подъём, гарантию и возврат. Замены согласовать отдельно; приложить паспорта/документы. Двери: полотно 2600 мм, цена полного комплекта. Акустика: спецификация всей системы. Работа и сборка: указать, что уже входит».',fmt['note'])
wb.close()

with (P/'Shopping_list.csv').open('w',encoding='utf-8-sig',newline='') as f:
 writer=csv.writer(f,delimiter=';',lineterminator='\n');writer.writerow(['Группа','Наименование','Количество','Единица','Цена-допущение, ₸','Сумма, ₸','Статус','Примечание'])
 for r in rows:writer.writerow([r['group'],r['item'],r['qty'],r['unit'],r['price'],r['qty']*r['price'],r['status'],r['note']])

# Small conceptual bathroom drawing: keep the original model dimensions explicit.
svg='''<svg xmlns="http://www.w3.org/2000/svg" width="1100" height="1000" viewBox="0 0 1100 1000">
<rect width="1100" height="1000" fill="#faf9f5"/><style>text{font-family:Arial,sans-serif;fill:#30352c} .s{font-size:16px}.l{font-size:20px}.d{stroke:#89917d;stroke-width:1.5;fill:none}</style>
<text x="55" y="55" font-size="28">Ванна в исходном санузле · предварительная расстановка</text>
<text x="55" y="87" class="s">5,1 м² по экспликации. Размеры 2,82 × 1,82 м взяты из условной модели — требуется обмер.</text>
<g transform="translate(155 185)">
<rect width="705" height="455" fill="#edece4" stroke="#5f6856" stroke-width="14"/>
<rect x="15" y="15" width="425" height="175" rx="12" fill="#fff" stroke="#8f9f97" stroke-width="3"/>
<rect x="38" y="38" width="379" height="129" rx="48" fill="#e6eef0" stroke="#c0d0d2"/>
<text x="227" y="100" text-anchor="middle" class="l">Ванна 170 × 70</text>
<text x="227" y="129" text-anchor="middle" class="s">доступ к сифону со стороны экрана</text>
<rect x="520" y="15" width="170" height="120" rx="5" fill="#c9b28c" stroke="#8c7b60" stroke-width="2"/>
<ellipse cx="605" cy="72" rx="57" ry="32" fill="#fff" stroke="#a7b4ae"/>
<text x="605" y="157" text-anchor="middle" class="s">Тумба 60–70 см</text>
<rect x="15" y="290" width="150" height="150" fill="#fff" stroke="#9ba59d" stroke-width="2"/>
<text x="90" y="350" text-anchor="middle" class="s">Стиральная</text><text x="90" y="375" text-anchor="middle" class="s">60 × 60</text>
<path d="M170 365 H250" class="d"/><path d="M250 365 l-10 -6 m10 6 l-10 6" class="d"/>
<text x="194" y="404" class="s">фронт →</text>
<rect x="648" y="266" width="42" height="110" rx="8" fill="#fff" stroke="#9ba59d"/>
<ellipse cx="605" cy="320" rx="62" ry="47" fill="#fff" stroke="#9ba59d" stroke-width="2"/>
<text x="604" y="411" text-anchor="middle" class="s">Напольный унитаз</text>
<text x="352" y="260" text-anchor="middle" class="l">Проход / обслуживание</text>
<line x1="220" y1="455" x2="420" y2="455" stroke="#faf9f5" stroke-width="17"/>
<path d="M220 455 V655 M420 455 A200 200 0 0 1 220 655" fill="none" stroke="#b88b64" stroke-width="2" stroke-dasharray="6 5"/>
<line x1="0" y1="-35" x2="705" y2="-35" class="d"/><line x1="0" y1="-43" x2="0" y2="-27" class="d"/><line x1="705" y1="-43" x2="705" y2="-27" class="d"/>
<text x="352" y="-49" text-anchor="middle" class="l">≈ 2 820 мм</text>
<line x1="-38" y1="0" x2="-38" y2="455" class="d"/><text transform="translate(-56 230) rotate(-90)" text-anchor="middle" class="l">≈ 1 820 мм</text>
</g>
<rect x="0" y="880" width="1100" height="120" fill="#faf9f5"/>
<text x="55" y="910" class="s">Открывание двери показано наружу условно: проверить прихожую, коллизии и проектную допустимость.</text>
<text x="55" y="938" class="s">Проверить стояк/выпуск, реальные корпуса, люк, зоны электрики и открывание стиральной машины.</text>
<text x="55" y="966" class="s">Это новая идея с ванной. В ранее опубликованном HTML показан душ; он пока не заменён.</text>
</svg>'''
(P/'Bathroom_proposal.svg').write_text(svg)

# Portable brief with embedded Cyrillic fonts and cached figures matching the workbook.
fontdir=Path('/usr/share/fonts/truetype/dejavu')
pdfmetrics.registerFont(TTFont('D',str(fontdir/'DejaVuSans.ttf')))
pdfmetrics.registerFont(TTFont('DB',str(fontdir/'DejaVuSans-Bold.ttf')))
pdfmetrics.registerFontFamily('D',normal='D',bold='DB',italic='D',boldItalic='DB')
styles=getSampleStyleSheet()
for name in ['Normal','BodyText','Title','Heading1','Heading2','Heading3']:styles[name].fontName='D'
styles.add(ParagraphStyle('SmallRu',fontName='D',fontSize=8.5,leading=12,textColor=colors.HexColor('#4c5146'),spaceAfter=7))
styles.add(ParagraphStyle('BodyRu',fontName='D',fontSize=10,leading=14,spaceAfter=8))
styles.add(ParagraphStyle('TitleRu',fontName='DB',fontSize=21,leading=27,textColor=colors.HexColor('#4a553d'),spaceAfter=12))
styles.add(ParagraphStyle('HeadRu',fontName='DB',fontSize=13,leading=18,spaceBefore=7,spaceAfter=9))
money=lambda n:f'{n:,.0f}'.replace(',',' ')+' ₸'
flow=[]
def p(t,style='BodyRu'):return Paragraph(t,styles[style])
def table(data,widths):
 t=Table([[p(str(c),'SmallRu') for c in row] for row in data],colWidths=widths,repeatRows=1,hAlign='LEFT')
 t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e6ebdc')),('GRID',(0,0),(-1,-1),.4,colors.HexColor('#d7d9cf')),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),3)]));return t
flow += [p('Санды-кала 2<br/>Материалы и смета','TitleRu'),p('Астана · 13.10.2026 · 41 м² по данным владельца · кухня на месте','SmallRu'),Image(str(P/'Interior_concept.png'),width=505,height=505*1024/1536),Spacer(1,10),p('Рекомендация: матовая краска, качественный ламинат в жилой комнате, плитка в кухне/прихожей/ванной, ванна 170×70, обычная кровать 140×200. Потолок гостиной и одна условная соседская стена — акустическая система.'),p('<b>Расчёт: '+money(subtotal+reserve)+'</b>, включая резерв '+money(reserve)+'. Это на '+money(subtotal+reserve-data['budget'])+' выше лимита. Крупная техника исключена.'),p('Цены товаров — допущения, не найденные прайсы: каталоги и поиск недоступны (CONNECT 403). Ставка бригады 50 000 ₸/м² сообщена владельцем. Изображение иллюстрирует стиль, а не точную геометрию.','SmallRu'),PageBreak()]
flow += [p('Сводная смета','TitleRu'),p('Работы считаются отдельно от материалов. Подготовка стен, сантехника и электрика входят в ставку; повторно их монтаж не начисляется. Специальную акустическую систему и высокие двери подтвердить в перечне работ.','SmallRu')]
tab=[['Статья','Расчёт','Лимит для 7 млн']]
tab += [[g,money(v),money(t)] for (g,v),t in zip(groups.items(),targets)]
tab += [['Без резерва',money(subtotal),money(sum(targets))],['Резерв',money(reserve),money(reserve)],['Всего',money(subtotal+reserve),money(data['budget'])]]
flow += [table(tab,[265,120,120]),Spacer(1,10),p('Правый столбец — целевые лимиты для запроса предложений, не подтверждённая корзина. Если при одинаковой комплектации цены выше, потребуется другой состав или этапность ремонта.','SmallRu'),p('Объёмы для проверки','HeadRu'),p('Полезная площадь — 38,8 м². Ламинат: 20,5 × 1,10 ≈ 23 м² к закупке. Напольная плитка: 18,3 × 1,10 ≈ 21 м². Стены ванной: условно 23 м², закупка 26 м². Стены под краску: условно 96 м², закупка краски 30 л. Акустический потолок: 20,5 м², двухслойная обшивка — 46 м² листов с запасом.','SmallRu'),p('Работы по договору: 41 м² = 2 050 000 ₸; 40,7 м² = 2 035 000 ₸; 38,8 м² = 1 940 000 ₸. Лоджия не отделывается в этой модели. Уточнить договорную площадь.','SmallRu'),PageBreak()]
flow += [p('Инженерия до закупки','TitleRu'),p('Электроплита и духовка','HeadRu'),table([['Линия','Предварительный кандидат','Проверка'],['Свет','3×1,5 мм²; часто до 10 А','Нагрузка и прокладка'],['Розетки / отдельные машины','3×2,5 мм²; обычно 16 А','Разделить кухонные группы'],['Духовка 3,5 кВт','3×2,5 / 16 А при подтверждении','≈15,2 А при 230 В'],['Панель 7,2 кВт','3×6 / 32 А как кандидат','≈31,3 А; ввод и условия прокладки']],[145,185,175]),Spacer(1,10),p('Панель и духовка вместе — 10,7 кВт, около 46,5 А. При вводе 32 А полной мощности не хватает: проверить выделенную мощность и ограничение нагрузки, не увеличивать автомат самовольно. Сечения, защиту и подключение окончательно выбирает электрик по расчёту и паспортам.','SmallRu'),p('Шумоизоляция и двери','HeadRu'),p('Потолок: виброподвесы, каркас, минвата, два слоя листов, герметичные примыкания по испытанной системе. Результат при ударах шагов не гарантирован: важен пол соседа и передача по стенам. Декоративные панели не заменяют эту конструкцию. Наружные стены угловой квартиры не утеплять изнутри автоматически.','SmallRu'),p('Пример высоты: 3000 − 60 пола − 140 потолка = 2800 мм. Полотно 2600 оставляет 200 мм сверху, но коробку, перемычку и чистые отметки проверяют отдельно. Если 3 м уже от чистого пола, слои повторно не вычитать.','SmallRu'),p('Вода и ванна','HeadRu'),p('Ванна 170×70 предварительно вдоль длинной стороны санузла, с заводской опорой и ревизией сифона. Гидроизоляция пола/мокрых стен с лентами и манжетами. Совместимая система труб и фитингов, доступная арматура, проверка до зашивки, защита от протечек. Кухню не переносим, длинная канализация через гостиную не требуется.','SmallRu'),p('Счётчики сначала проверить у застройщика: паспорта, поверка и ввод. В базе новые приборы не покупаются. Кухонную вытяжку не вести в канал туалета; штатный кухонный канал и режим подключения обследовать.','SmallRu'),PageBreak()]
flow += [p('Выбор и закупка','TitleRu'),p('Краска, ламинат, мебель','HeadRu'),p('Краска не защищает от усадки. Важны сухое прочное основание, совместимая подготовка, обработка стыков и сохранение деформационных швов. Стеклохолст может уменьшать проявление мелких поверхностных трещин, но не удерживает активную конструкционную трещину.','SmallRu'),p('При лимите 7 млн предпочтителен ламинат 8–10 мм с документами, качественным замком и подходящей подложкой. Инженерная доска в условном сценарии добавляет около 436 500 ₸. Деревянные покрытия требуют микроклимата по паспорту.','SmallRu'),p('Обычная кровать дешевле и проще для ежедневного сна. Шкаф-кровать имеет смысл ради свободного места днём; в условном сценарии доплата 550 000 ₸. Проверить основание крепления, механизм, массу матраса и пространство раскрытия.','SmallRu'),p('Стекло и стол','HeadRu'),p('Частичное раскрытие стены кухни — только после проверки конструкции. Стол 90×50 на надёжной опоре/механизме, не на стекле; проверить проход и выход на лоджию. В базе — стол на существующей стене и штора у кровати, без демонтажа.','SmallRu'),p('Цены и качество','HeadRu'),p('Проверены попытки доступа к 12.kz, akvilon.kz, leroymerlin.kz и satu.kz: CONNECT HTTP 403. Цены, наличие и доставка в Астану не получены. Наиболее дешёвый рынок/продавец не установлен. Примеры марок в полной записке — кандидаты для сравнения конкретных продуктов, не проверенные предложения.','SmallRu'),p('Запросить три КП на одинаковые артикулы/комплекты с доставкой, подъёмом, налогами, сроком и гарантией. Особенно проверить происхождение кабеля, автоматики, труб и гидроизоляции; у плитки тон/калибр, у мешков дату/сухое хранение. На этих узлах не заменять спецификацию случайным самым дешёвым товаром.','SmallRu'),p('Ограничения','HeadRu'),p('Не учтены техника, кондиционер, тёплый пол, лоджия, перенос кухни/стояков, согласование/проёмы, входная дверь/окна/радиаторы, стеклянная перепланировка. База предполагает пригодную стяжку; новая толстая стяжка — доплата после расчёта. Удалить повторную доставку/сборку, если она включена в договор.','SmallRu'),p('Полная записка, редактируемая смета и эскиз санузла: <link href="https://github.com/Monste11osz/dis/tree/main/design/materials" color="#176AB1">github.com/Monste11osz/dis/design/materials</link>','SmallRu')]
def footer(canvas,doc):
 canvas.setFont('D',8);canvas.setFillColor(colors.HexColor('#6c7164'));canvas.drawString(42,24,'Санды-кала 2 · Предпроектная смета · 13.10.2026');canvas.drawRightString(553,24,str(doc.page))
doc=SimpleDocTemplate(str(P/'Renovation_brief.pdf'),pagesize=(595.28,841.89),rightMargin=45,leftMargin=45,topMargin=36,bottomMargin=40)
doc.build(flow,onFirstPage=footer,onLaterPages=footer)
print(json.dumps({'rows':len(rows),'subtotal':subtotal,'reserve':reserve,'total':subtotal+reserve,'scenarios_with_reserve':{k:v+reserve for k,v in totals.items()}},ensure_ascii=False))
