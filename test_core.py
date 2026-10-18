from datetime import date,timedelta
from pathlib import Path
from unittest.mock import Mock,patch
import pytest
import core
from storage import LocalStore,OneDriveStore,digest

@pytest.fixture
def book():return core.initialize(100000,date(2026,1,1))
def customer(wb,**kwargs):
    return core.perform(wb,'add_client',{'اسم العميل':'عميل اختبار','رقم الجوال':'0500000000','رقم الهوية':'1000000001','حالة العميل':'نشط','التصنيف الائتماني':'A','الحد الائتماني':0,**kwargs})
def contract(wb,cid=None,p=5000,total=8000,n=16,first=date(2026,2,28),**kwargs):
    cid=cid or customer(wb)
    return core.perform(wb,'add_contract',{'رقم العميل':cid,'مبلغ التمويل':p,'إجمالي السداد':total,'مدة العقد':n,'تاريخ العقد':date(2026,1,1),'أول قسط':first,**kwargs})
def trans(wb,kind,amount,cid='',when=date(2026,1,2),**kwargs):
    return core.perform(wb,'transaction',dict(kind=kind,amount=amount,when=when,contract=cid,**kwargs))
def funded(wb,**kwargs):
    c=contract(wb,**kwargs);trans(wb,'صرف تمويل',kwargs.get('p',5000),c);return c

def test_seed_and_excel_roundtrip(book):
    wb=core.load(core.dump(book));assert core.snapshot(wb)['cash_h']==10000000
    assert len(wb.sheetnames)==9

def test_draft_does_not_move_cash(book):
    c=contract(book);s=core.snapshot(book)
    assert s['cash_h']==10000000 and not s['contracts'][0]['disbursed']

def test_disburse_and_partial_multi_installment(book):
    c=funded(book)
    trans(book,'سداد قسط',750,c,date(2026,2,28))
    s=core.snapshot(book);p=s['installments']
    assert s['cash_h']==9575000
    assert p[0]['paid_h']==50000 and p[1]['paid_h']==25000 and p[1]['remaining_h']==25000
    assert s['contracts'][0]['paid_count']==1
    assert sum(x['paid_h'] for x in p)==75000

def test_month_end_and_exact_cents(book):
    c=contract(book,p=900,total=1000,n=3,first=date(2026,1,31))
    p=core.snapshot(book)['installments']
    assert [x['due'] for x in p]==[date(2026,1,31),date(2026,2,28),date(2026,3,31)]
    assert [x['total_h'] for x in p]==[33333,33333,33334]

def test_settlement_discount_not_cash(book):
    c=funded(book);trans(book,'سداد قسط',1000,c,date(2026,2,28))
    trans(book,'سداد مبكر',6000,c,date(2026,3,1),discount=1000)
    s=core.snapshot(book);ct=s['contracts'][0]
    assert ct['status']=='مسدد مبكرًا' and ct['remaining_h']==0 and ct['discount_h']==100000
    assert ct['paid_h']==700000 and s['cash_h']==10200000
    assert sum(x['waived_h'] for x in s['installments'])==100000

def test_void_and_restore_payment(book):
    c=funded(book);t=trans(book,'سداد قسط',500,c,date(2026,2,28))
    core.perform(book,'void_transaction',{'id':t});assert core.snapshot(book)['contracts'][0]['remaining_h']==800000
    core.perform(book,'restore_transaction',{'id':t});assert core.snapshot(book)['contracts'][0]['remaining_h']==750000

def test_overpayment_and_double_disbursement(book):
    c=funded(book)
    with pytest.raises(core.ValidationError): trans(core.load(core.dump(book)),'سداد قسط',8000.01,c)
    with pytest.raises(core.ValidationError): trans(core.load(core.dump(book)),'صرف تمويل',5000,c)

def test_payment_before_disbursement(book):
    c=contract(book)
    with pytest.raises(core.ValidationError):trans(book,'سداد قسط',500,c)

def test_dependency_prevents_void_disbursement(book):
    c=funded(book);trans(book,'سداد قسط',500,c)
    d=next(r['رقم العملية'] for r in core.active_tx(book) if r['نوع العملية']=='صرف تمويل')
    with pytest.raises(core.ValidationError):core.perform(core.load(core.dump(book)),'void_transaction',{'id':d})

def test_credit_limit(book):
    cl=customer(book,**{'الحد الائتماني':4000});c=contract(book,cid=cl)
    with pytest.raises(core.ValidationError):trans(book,'صرف تمويل',5000,c)

def test_reference_and_request_idempotency(book):
    req='same-request';data=dict(kind='إضافة رأس مال',amount=1000,when=date(2026,1,2),reference='bank-ref')
    core.perform(book,'transaction',data,req);core.perform(book,'transaction',data,req)
    assert core.cash(book)==10100000
    with pytest.raises(core.ValidationError):core.perform(book,'transaction',data,'different-request')

def test_formula_injection_and_identity_leading_zero(book):
    c=customer(book,**{'اسم العميل':'=HYPERLINK("bad")','رقم الهوية':'0012345678'})
    raw=core.dump(book);wb=core.load(raw)
    assert wb['العملاء']['B2'].data_type=='s' and wb['العملاء']['C2'].value=='0012345678'

def test_cash_insufficient(book):
    with pytest.raises(core.ValidationError):trans(book,'سحب رأس مال',100000.01)

def test_fees_separate_from_installments(book):
    c=funded(book,**{'رسوم العقد':200});trans(book,'رسوم عقد',200,c)
    s=core.snapshot(book)
    assert s['contracts'][0]['remaining_h']==800000 and s['cash_h']==9520000
    with pytest.raises(core.ValidationError):trans(core.load(core.dump(book)),'رسوم عقد',.01,c)

def test_late_history_retained(book):
    c=funded(book);trans(book,'سداد قسط',500,c,date(2026,3,2))
    p=core.snapshot(book)['installments'][0]
    assert p['late_days']==2 and p['status']=='مسدد'

def test_local_atomic_failure_and_stale_version(book,tmp_path):
    path=tmp_path/'system.xlsx';path.write_bytes(core.dump(book));store=LocalStore(path)
    raw,tag=store.read()
    with pytest.raises(core.ValidationError):store.commit(tag,'transaction',dict(kind='سحب رأس مال',amount=200000,when=date(2026,1,3)),'fail')
    assert path.read_bytes()==raw
    store.commit(tag,'transaction',dict(kind='إضافة رأس مال',amount=100,when=date(2026,1,3)),'ok')
    with pytest.raises(core.ConflictError):store.commit(tag,'settings',{'default_months':10,'default_profit_rate':.5},'stale')
    assert list((tmp_path/'backups').glob('*.xlsx'))

def test_edit_payment_replaces_without_duplication(book):
    c=funded(book);t=trans(book,'سداد قسط',500,c,date(2026,2,28),reference='ref')
    core.perform(book,'edit_transaction',{'id':t,'replacement':dict(kind='سداد قسط',amount=750,when=date(2026,2,28),contract=c,reference='ref')})
    s=core.snapshot(book)
    assert s['contracts'][0]['paid_h']==75000
    assert len([x for x in s['transactions'] if x['حالة العملية']=='ملغاة'])==1

def test_draft_edit_and_followup(book):
    cid=customer(book);c=contract(book,cid=cid)
    core.perform(book,'edit_contract',{'id':c,'رقم العميل':cid,'مبلغ التمويل':1000,'إجمالي السداد':1600,'مدة العقد':2,'تاريخ العقد':date(2026,1,1),'أول قسط':date(2026,2,1)})
    assert len(core.snapshot(book)['installments'])==2
    f=core.perform(book,'add_followup',{'رقم العميل':cid,'رقم العقد':c,'تاريخ التواصل':date(2026,1,1),'المتابعة القادمة':date(2026,2,1),'المبلغ الموعود':200})
    core.perform(book,'close_followup',{'id':f});assert core.rows(book,'متابعة العملاء')[0]['الحالة']=='مكتملة'

def test_graph_conflict_does_not_upload(book):
    store=OneDriveStore('TEST-NOT-A-TOKEN');store.metadata=Mock(return_value={'id':'a','eTag':'new'})
    store.session.post=Mock()
    with pytest.raises(core.ConflictError):store._publish('old',core.dump(book))
    store.session.post.assert_not_called()

def test_graph_upload_no_bearer_and_etag(book):
    store=OneDriveStore('TEST-NOT-A-TOKEN');store.item_id='item'
    store.metadata=Mock(return_value={'id':'item','eTag':'old'})
    response=Mock(status_code=200);response.json.return_value={'uploadUrl':'https://example.invalid/upload'}
    store.session.post=Mock(return_value=response)
    finish=Mock(status_code=200);finish.json.return_value={'eTag':'new'}
    with patch('storage.requests.put',return_value=finish) as upload:
        raw,tag=store._publish('old',core.dump(book))
        assert tag=='new'
        assert store.session.post.call_args.kwargs['headers']['If-Match']=='old'
        assert 'Authorization' not in upload.call_args.kwargs['headers']

def test_uploaded_workbook_rejects_duplicate_identity(book):
    customer(book)
    core.append(book,'العملاء',{'رقم العميل':'OTHER','اسم العميل':'آخر','رقم الهوية':'1000000001','رقم الجوال':'0555555555','حالة العميل':'نشط','التصنيف الائتماني':'A','الحد الائتماني':0,'الدخل الشهري':0})
    with pytest.raises(core.ValidationError):core.load(core.dump(book))

def test_graph_change_during_staging_prevents_final_upload(book):
    store=OneDriveStore('TEST');store.item_id='item'
    store.metadata=Mock(side_effect=[{'id':'item','eTag':'old'},{'id':'item','eTag':'changed'}])
    response=Mock(status_code=200);response.json.return_value={'uploadUrl':'https://example.invalid/upload'}
    store.session.post=Mock(return_value=response)
    with patch('storage.requests.put') as upload:
        with pytest.raises(core.ConflictError):store._publish('old',core.dump(book))
        upload.assert_not_called()

def test_void_clears_excel_completion_date(book):
    c=funded(book);p=trans(book,'سداد قسط',500,c,date(2026,2,28))
    assert core.day(book['الأقساط']['J2'].value)==date(2026,2,28)
    core.perform(book,'void_transaction',{'id':p})
    assert book['الأقساط']['J2'].value is None

def test_settlement_clears_next_due_and_followup_edit(book):
    c=funded(book);assert book['العقود']['R2'].value is not None
    trans(book,'سداد مبكر',7000,c,date(2026,2,28),discount=1000)
    assert book['العقود']['R2'].value is None
    cid=core.rows(book,'العملاء')[0]['رقم العميل']
    f=core.perform(book,'add_followup',{'رقم العميل':cid,'رقم العقد':c,'وعد السداد':date(2026,2,28),'المبلغ الموعود':500})
    core.perform(book,'edit_followup',{'id':f,'وعد السداد':None,'المبلغ الموعود':0,'الحالة':'مكتملة','نتيجة التواصل':'تم السداد'})
    record=core.rows(book,'متابعة العملاء')[0]
    assert record['وعد السداد'] is None and record['الحالة']=='مكتملة'
