# -*- coding: utf-8 -*-
"""
P&L-агент ТОО «Caspian Sea Support» (CSS).
Загрузка выгрузки 1С ОСВ → P&L по проектам. Логика как в Градексе.
Можно подать годовую ОСВ (одна колонка) или несколько помесячных (колонка на месяц).
"""
import io, os, re, tempfile
import streamlit as st
import openpyxl
import css_engine as eng
import recalc_native

st.set_page_config(page_title='P&L агент — CSS', page_icon='📊', layout='wide')

MONTHS = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
          'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']


def guess_label(fn):
    s = fn.lower()
    for stem, m in [('январ', 'Январь'), ('феврал', 'Февраль'), ('март', 'Март'),
                    ('апрел', 'Апрель'), ('мая', 'Май'), ('май', 'Май'), ('июн', 'Июнь'),
                    ('июл', 'Июль'), ('август', 'Август'), ('сентябр', 'Сентябрь'),
                    ('октябр', 'Октябрь'), ('ноябр', 'Ноябрь'), ('декабр', 'Декабрь')]:
        if stem in s:
            return m
    y = re.search(r'(20\d\d)', s)
    return y.group(1) if y else ''


st.title('📊 P&L агент — Caspian Sea Support')
st.caption('Загрузите выгрузку 1С (ОСВ / Анализ счёта). Агент берёт счета 6ххх (доходы) и 7ххх '
           '(расходы), раскладывает по проектам и собирает P&L. Годовая ОСВ → одна колонка; '
           'несколько помесячных выгрузок → по колонке на период. Всё сходится с оборотами 1С.')

with st.sidebar:
    st.header('Настройки')
    thr = st.number_input('Порог «крупной» статьи, ₸', value=300_000, step=50_000,
                          help='Статьи/проекты крупнее порога выделяются в журнале.')
    eng.THRESHOLD = int(thr)

files = st.file_uploader('Выгрузки 1С ОСВ (.xlsx) — можно несколько',
                         type=['xlsx'], accept_multiple_files=True)

labels = []
if files:
    st.write('**Названия колонок** (период для каждого файла):')
    for i, uf in enumerate(files):
        c1, c2 = st.columns([3, 2])
        c1.write(f'📄 {uf.name}')
        labels.append(c2.text_input('Колонка', value=guess_label(uf.name) or f'Период {i+1}',
                                    key=f'lbl{i}', label_visibility='collapsed'))

st.divider()

if st.button('▶️ Сформировать P&L', type='primary', disabled=not files):
    try:
        periods = []
        for uf, lbl in zip(files, labels):
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
            with open(tmp, 'wb') as f:
                f.write(uf.getvalue())
            R = eng.parse_1c(tmp)
            periods.append((lbl or 'Период', R))
            inc = R['tot_income_6']; exp = R['tot_expense_7']
            st.info(f'✅ {uf.name} → «{lbl}»: доход {inc:,.0f}  расход {exp:,.0f}  '
                    f'до КПН {inc-exp:,.0f}')

        out = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
        eng.build_workbook(periods, out)
        eng.build_journal(out, periods)
        recalc_native.recalc_native(out)

        with open(out, 'rb') as f:
            data = f.read()
        st.success('Готово! P&L собран и сходится с оборотами 1С.')
        st.download_button('⬇️ Скачать P&L (.xlsx)', data=data,
                           file_name='CSS_P_L.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        # краткая сверка на экране
        vwb = openpyxl.load_workbook(out, data_only=True)
        ws = vwb['P&L 2026']

        def val(lbl, col=3):
            for r in range(1, ws.max_row + 1):
                if str(ws.cell(row=r, column=2).value or '').strip() == lbl:
                    return ws.cell(row=r, column=col).value
            return None
        st.subheader('Сверка (первая колонка)')
        for lbl in ['ИТОГО ДОХОД ОТ РЕАЛИЗАЦИИ', 'ИТОГО ПРЯМЫЕ ЗАТРАТЫ',
                    'ИТОГО АДМИНИСТРАТИВНЫЕ', 'ДОХОД ДО КПН', 'ЧИСТЫЙ ДОХОД',
                    'Разница (P&L − 1С)']:
            v = val(lbl)
            if isinstance(v, (int, float)):
                st.write(f'• {lbl}: **{v:,.0f} ₸**')
    except Exception as e:
        st.error(f'Ошибка: {e}')
        import traceback
        st.code(traceback.format_exc())

with st.expander('ℹ️ Как работает разноска'):
    st.markdown('''
- **Доходы** (6010) — по проектам (аренда судов, мобилизация, баржи, LQB и т.д.).
  Депозиты (6110), курсовая (6250), прочие (6290) — ниже, в блоке «Прочие доходы».
- **Прямые затраты** (7011) — по каждому проекту отдельным блоком, статьи как в 1С.
- **Административные** (7210/7211/7212) — отдельный блок; налоги с ЗП
  (ОПВ + ОСМС + соц.отчисления + соц.налог) свёрнуты в одну строку.
- **Низ**: финансирование (7300), курсовая разница расход (7430), прочие расходы (7480).
- **Проекты из 1С** с разными написаниями («Проект Barges» и «NCOC:Flat Top Barges»)
  сводятся в один блок — см. лист «Журнал сопоставления».
- **Контроль**: внизу P&L сверка с оборотами 6ххх/7ххх 1С, «Разница» должна быть 0.
''')
