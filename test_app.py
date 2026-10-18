"""UI tests use a temporary XLSX. Never put demonstration clients in the delivered seed."""
from pathlib import Path
from datetime import date
import pytest
from streamlit.testing.v1 import AppTest
import core

def widget(elements,label):return next(w for w in elements if w.label==label)

def test_create_client_from_ui_persists_to_excel(tmp_path,monkeypatch):
    path=tmp_path/'system.xlsx';path.write_bytes(core.dump(core.initialize()))
    monkeypatch.setenv('INSTALLMENT_WORKBOOK',str(path))
    at=AppTest.from_file(str(Path(__file__).with_name('app.py')),default_timeout=30).run()
    at.radio(key='nav').set_value('👤 عملاء التقسيط').run()
    widget(at.text_input,'اسم العميل *').set_value('عميل تجربة الواجهة')
    widget(at.text_input,'رقم الجوال *').set_value('0501234567')
    widget(at.button,'حفظ بيانات العميل').click().run()
    assert not at.exception and not at.error
    assert core.rows(core.load(path.read_bytes()),'العملاء')[0]['اسم العميل']=='عميل تجربة الواجهة'
    # Every page must also render with real records, not only empty placeholders.
    wb=core.load(path.read_bytes());cid=core.rows(wb,'العملاء')[0]['رقم العميل']
    ct=core.perform(wb,'add_contract',{'رقم العميل':cid,'مبلغ التمويل':5000,'إجمالي السداد':8000,'مدة العقد':16,'تاريخ العقد':date(2026,7,1),'أول قسط':date(2026,8,1),'رسوم العقد':200})
    core.perform(wb,'transaction',dict(kind='صرف تمويل',amount=5000,when=date(2026,7,2),contract=ct))
    path.write_bytes(core.dump(wb))
    at=AppTest.from_file(str(Path(__file__).with_name('app.py')),default_timeout=30).run()
    for page in at.radio(key='nav').options:
        at.radio(key='nav').set_value(page).run()
        assert not at.exception,(page,[e.value for e in at.exception])
        assert not at.error,(page,[e.value for e in at.error])

def test_receipt_from_ui_changes_excel_and_dashboard(tmp_path,monkeypatch):
    wb=core.initialize();cid=core.perform(wb,'add_client',{'اسم العميل':'عميل اختبار','رقم الجوال':'0500000000'})
    ct=core.perform(wb,'add_contract',{'رقم العميل':cid,'مبلغ التمويل':5000,'إجمالي السداد':8000,'مدة العقد':16,'تاريخ العقد':date(2026,7,1),'أول قسط':date(2026,8,1)})
    core.perform(wb,'transaction',dict(kind='صرف تمويل',amount=5000,when=date(2026,7,2),contract=ct))
    path=tmp_path/'system.xlsx';path.write_bytes(core.dump(wb));monkeypatch.setenv('INSTALLMENT_WORKBOOK',str(path))
    at=AppTest.from_file(str(Path(__file__).with_name('app.py')),default_timeout=30).run()
    at.radio(key='nav').set_value('📄 العقود').run()
    # The ordinary payment form has the first matching amount input.
    widget(at.number_input,'المبلغ المدفوع فعليًا · ر.س').set_value(750.)
    widget(at.button,'تأكيد سداد قسط وحفظه').click().run()
    assert not at.exception and not at.error
    snap=core.snapshot(core.load(path.read_bytes()))
    assert snap['cash_h']==9575000 and snap['contracts'][0]['paid_h']==75000
    at.radio(key='nav').set_value('📊 ملخص المحفظة').run()
    assert at.metric[0].value=='95,750.00'
