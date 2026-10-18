from __future__ import annotations
from datetime import date, timedelta
from pathlib import Path
from io import BytesIO
from html import escape
from uuid import uuid4
import hmac, os
import pandas as pd
import streamlit as st
import core
from storage import LocalStore, OneDriveStore, atomic_write

ROOT=Path(__file__).resolve().parent
PAGES=['📊 ملخص المحفظة','👤 عملاء التقسيط','🤝 المستثمرون','📄 العقود','💰 الحركات المالية','📞 متابعة العملاء','📈 التقارير','⚙️ الإعدادات والنسخ']

def cfg(section,key,default=''):
    try: return st.secrets[section].get(key,default)
    except Exception: return default

def theme():
    st.markdown('''<style>
    :root {color-scheme:dark}
    .stApp {background:#111719;color:#e8eeee;direction:rtl;font-family:Tahoma,"Segoe UI",Arial,sans-serif}
    [data-testid="stSidebar"] {background:#171e22;border-left:1px solid #2d393e;direction:rtl}
    h1,h2,h3,p,label {font-family:Tahoma,"Segoe UI",Arial,sans-serif!important}
    .stApp h1 {font-size:2.1rem!important;font-weight:800;letter-spacing:0;margin-bottom:.3rem}
    .stApp h2 {font-size:1.5rem!important} .stApp h3 {font-size:1.2rem!important}
    [data-testid="stCaptionContainer"] {color:#aab6b8}
    [data-testid="stMetric"] {background:#1b2327;border:1px solid #2d393e;border-radius:15px;padding:18px 20px;min-height:115px}
    [data-testid="stMetricValue"] {color:#42c6bd;font-size:1.9rem;direction:ltr;text-align:right}
    [data-testid="stMetricLabel"] {color:#aab6b8;font-size:.95rem}
    [data-testid="stForm"] {border:1px solid #2d393e;border-radius:16px;background:#182126;padding:22px}
    .stButton button[kind="primary"], .stFormSubmitButton button[kind="primary"] {background:#42c6bd;color:#102023;border:0;font-weight:700}
    .stButton button,.stDownloadButton button {border-radius:10px}
    [data-testid="stDataFrame"] {border:1px solid #2d393e;border-radius:12px}
    [data-testid="stSidebar"] [data-testid="stRadio"] label {padding:7px 4px}
    .eyebrow {color:#42c6bd;font-size:.88rem;letter-spacing:.04em}
    .section-note {color:#aab6b8;padding:8px 0 14px}
    .profile {background:#1b2327;border:1px solid #2d393e;border-radius:16px;padding:22px;margin:12px 0 20px}
    .profile strong {font-size:1.5rem}.profile p {color:#aab6b8;margin:10px 0 0}
    .block-container {max-width:1550px;padding-top:2.2rem;padding-bottom:3rem}
    @media(max-width:700px){.block-container{padding:1rem}.stApp h1{font-size:1.6rem!important}[data-testid="stMetricValue"]{font-size:1.45rem}}
    </style>''',unsafe_allow_html=True)

def fmt(h): return f'{core.sar(h):,.2f}'
def notify(message): st.session_state['notice']=message

def submit(action,data):
    request_id=st.session_state.setdefault('pending_request_id',core.uid('REQ'))
    try:
        ident,raw,tag=st.session_state['_store'].commit(st.session_state['_tag'],action,data,request_id)
        st.session_state['_raw']=raw;st.session_state['_tag']=tag
        st.session_state.pop('pending_request_id',None)
        notify(f'تم الحفظ في Excel بنجاح · {ident}')
        st.rerun()
    except Exception as exc:
        st.error(str(exc))
        st.caption(f'رقم الطلب: {request_id}. إذا انقطع الاتصال، حدّث المصدر وتحقق من سجل التغييرات قبل إعادة الإدخال.')

def client_names(snap): return {c['رقم العميل']:c['اسم العميل'] for c in snap['clients']}
def investor_names(snap): return {c['رقم المستثمر']:c['اسم المستثمر'] for c in snap['investors']}
def choose_client(snap,key,active_only=False):
    clients=[c for c in snap['clients'] if not active_only or c['حالة العميل']=='نشط']
    if not clients: st.info('أضف عميلًا نشطًا من صفحة العملاء أولًا.');return None
    byid={c['رقم العميل']:c for c in clients}
    return st.selectbox('العميل',list(byid),format_func=lambda x:f"{byid[x]['اسم العميل']} · {x}",key=key)
def table(data,empty='لا توجد سجلات حتى الآن.'):
    if not data: st.info(empty);return
    df=pd.DataFrame(data)
    for c in df.columns:
        if any(isinstance(x,(core.date,core.datetime)) for x in df[c] if x is not None):
            df[c]=df[c].map(lambda x:x.isoformat() if isinstance(x,(core.date,core.datetime)) else x)
    st.dataframe(df,hide_index=True,width='stretch')
def metrics(items):
    cols=st.columns(min(4,len(items)))
    for i,(label,value) in enumerate(items): cols[i%len(cols)].metric(label,value)

def contract_table(contracts,names):
    return [{'رقم العقد':c['رقم العقد'],'العميل':names.get(c['client_id'],c['client_id']),'نوع الصفقة':c['deal_type'],
       'تكلفة الشراء':core.sar(c['purchase_cost_h']),'سعر الكاش':core.sar(c['cash_price_h']),'إجمالي التقسيط':core.sar(c['total_h']),
       'صافي ربحك المتوقع':core.sar(c['office_expected_profit_h']),'مبلغ القسط':core.sar(c['total_h']//int(c['مدة العقد'])),
       'المحصل':core.sar(c['paid_h']),'المتبقي':core.sar(c['remaining_h']),'حالة العقد':c['status'],
       'الأقساط المسددة':c['paid_count'],'الأقساط المتبقية':int(c['مدة العقد'])-c['paid_count'],
       'الاستحقاق القادم':c['next_due']} for c in contracts]
def schedule_table(parts,names=None):
    return [{'رقم العقد':p['رقم العقد'],**({'العميل':names.get(p['client_id'],'')} if names else {}),'رقم القسط':p['تسلسل القسط'],
       'الاستحقاق':p['due'],'المبلغ':core.sar(p['total_h']),'المحصل':core.sar(p['paid_h']),
       'خصم التسوية':core.sar(p['waived_h']),'المتبقي':core.sar(p['remaining_h']),'الحالة':p['status'],'أيام التأخير':p['late_days'],
       'اكتمال السداد':p['paid_date']} for p in parts]
def tx_table(rr):
    return [{k:r.get(k) for k in ['رقم العملية','التاريخ','نوع العملية','رقم العقد','رقم المستثمر','ملكية الحركة','المبلغ','المحصل','المرجع','حالة العملية','ملاحظات']} for r in rr]

def dashboard(wb,snap):
    names=client_names(snap);funded=[c for c in snap['contracts'] if c['disbursed']]
    owned=[c for c in funded if c['deal_type']!=core.DEAL_MANAGED];managed=[c for c in funded if c['deal_type']==core.DEAL_MANAGED]
    metrics([('السيولة الحرة للمكتب · ر.س',fmt(snap['cash_h'])),('رصيد الحساب البنكي · ر.س',fmt(snap['bank_cash_h'])),
      ('أموال المستثمرين المحتجزة · ر.س',fmt(snap['managed_cash_h'])),('ربح المكتب المتوقع · ر.س',fmt(sum(c['office_expected_profit_h'] for c in funded))),
      ('متبقي عقودك · ر.س',fmt(sum(c['remaining_h'] for c in owned))),
      ('متبقي عقود المستثمرين · ر.س',fmt(sum(c['remaining_h'] for c in managed))),('عقود قائمة',str(sum(c['remaining_h']>0 for c in funded))),
      ('مستثمرون',str(len(snap['investors'])))])
    st.caption('السيولة الحرة = رصيد الحساب البنكي ناقص الأموال المحتجزة للمستثمرين. أقساط المستثمر لا تُحسب دخلًا للمكتب.')
    st.subheader('العقود')
    table(contract_table(snap['contracts'],names),'لا توجد عقود. ابدأ بإضافة عميل، ثم إنشاء عقد وصرف تمويله.')
    a,b=st.columns(2)
    with a:
        st.subheader('⏳ مستحق خلال 7 أيام')
        parts=[p for p in snap['installments'] if p['disbursed'] and p['remaining_h']>0 and core.today()<=p['due']<=core.today()+timedelta(days=7)]
        table(schedule_table(parts,names),'لا توجد أقساط مستحقة خلال الأيام السبعة القادمة.')
    with b:
        st.subheader('⚠️ العقود المتأخرة')
        table(contract_table([c for c in funded if c['overdue_h']],names),'لا توجد عقود متأخرة.')
    st.subheader('آخر التحصيلات')
    table(tx_table(list(reversed([r for r in snap['transactions'] if r['نوع العملية'] in core.RECEIPTS and r['حالة العملية']=='نشطة']))[:10]),'لم تسجل تحصيلات بعد.')
    a,b=st.columns(2)
    with a:
        st.subheader('حركة السيولة')
        table(tx_table(list(reversed([r for r in snap['transactions'] if r['حالة العملية']=='نشطة']))[:8]))
    with b:
        st.subheader('المتابعات المطلوبة')
        pending=[r for r in snap['followups'] if r['الحالة']=='مفتوحة' and r['المتابعة القادمة'] and core.day(r['المتابعة القادمة'])<=core.today()]
        table([{'العميل':names.get(r['رقم العميل'],''),'المتابعة القادمة':r['المتابعة القادمة'],'النتيجة السابقة':r['نتيجة التواصل']} for r in pending],'لا توجد متابعات مستحقة.')

def client_form(snap,existing=None):
    old=existing or {};prefix=f"client-{old.get('رقم العميل','new')}"
    with st.form(prefix):
        a,b=st.columns(2)
        name=a.text_input('اسم العميل *',value=old.get('اسم العميل',''))
        identity=b.text_input('رقم الهوية',value=old.get('رقم الهوية',''),max_chars=30)
        phone=a.text_input('رقم الجوال *',value=old.get('رقم الجوال',''),max_chars=30)
        joined=b.date_input('تاريخ إضافة العميل',value=core.day(old.get('تاريخ إضافة العميل')) or core.today(),disabled=bool(existing))
        state=a.selectbox('حالة العميل',core.CLIENT_STATES,index=core.CLIENT_STATES.index(old.get('حالة العميل','نشط')))
        grade=b.selectbox('التصنيف الائتماني الداخلي',core.GRADES,index=core.GRADES.index(old.get('التصنيف الائتماني','غير مصنف')))
        limit=a.number_input('الحد الائتماني · ر.س',min_value=0.,value=float(old.get('الحد الائتماني') or 0),step=1000.,help='صفر يعني عدم وضع حد. يقارن الحد بأصل التمويل المتبقي تقديريًا عند الصرف.')
        income=b.number_input('الدخل الشهري · ر.س',min_value=0.,value=float(old.get('الدخل الشهري') or 0),step=100.)
        employer=a.text_input('جهة العمل',value=old.get('جهة العمل',''))
        address=b.text_input('العنوان',value=old.get('العنوان',''))
        guarantor=a.text_input('اسم الكفيل',value=old.get('اسم الكفيل',''))
        guarantor_phone=b.text_input('جوال الكفيل',value=old.get('جوال الكفيل',''))
        notes=st.text_area('ملاحظات',value=old.get('ملاحظات',''))
        st.caption('التصنيف الداخلي يدوي. عدد مرات التأخير يحسب من الأقساط المصروفة، ويشمل الأقساط المسددة بعد موعدها.')
        if st.form_submit_button('حفظ بيانات العميل',type='primary'):
            data={'اسم العميل':name,'رقم الهوية':identity,'رقم الجوال':phone,'تاريخ إضافة العميل':joined,'حالة العميل':state,
              'التصنيف الائتماني':grade,'الحد الائتماني':limit,'الدخل الشهري':income,'جهة العمل':employer,'العنوان':address,
              'اسم الكفيل':guarantor,'جوال الكفيل':guarantor_phone,'ملاحظات':notes}
            if existing: data['id']=old['رقم العميل']
            submit('edit_client' if existing else 'add_client',data)

def statement_html(client,cs,parts,transactions):
    def grid(data):
        if not data:return '<p>لا توجد سجلات</p>'
        heads=list(data[0]);return '<table><thead><tr>'+''.join(f'<th>{escape(str(k))}</th>' for k in heads)+'</tr></thead><tbody>'+''.join('<tr>'+''.join(f'<td>{escape(str(r.get(k) or ""))}</td>' for k in heads)+'</tr>' for r in data)+'</tbody></table>'
    return f'''<!doctype html><html lang="ar" dir="rtl"><meta charset="utf-8"><title>كشف حساب العميل</title><style>body{{font-family:Tahoma,Arial;padding:35px;color:#16323b}}table{{border-collapse:collapse;width:100%;margin:20px 0;font-size:12px}}th,td{{border:1px solid #ddd;padding:9px;text-align:right}}th{{background:#e1f1ef}}h1{{font-size:24px}}@media print{{body{{padding:0}}}}</style><h1>كشف حساب: {escape(client['اسم العميل'])}</h1><p>رقم العميل: {escape(client['رقم العميل'])} · تاريخ الكشف: {core.today()}</p><h2>العقود</h2>{grid(contract_table(cs,{}))}<h2>الأقساط</h2>{grid(schedule_table(parts))}<h2>الحركات</h2>{grid(tx_table(transactions))}</html>'''.encode('utf-8')

def clients_page(wb,snap):
    view,create=st.tabs(['سجل العملاء وملف العميل','إضافة عميل'])
    with create: client_form(snap)
    with view:
        query=st.text_input('ابحث بالاسم أو الجوال أو الهوية',key='client_search')
        state=st.selectbox('تصفية الحالة',['الكل']+core.CLIENT_STATES,key='client_state_filter')
        rr=[c for c in snap['clients'] if (state=='الكل' or c['حالة العميل']==state) and query.strip().casefold() in ' '.join(str(c.get(k) or '') for k in ['اسم العميل','رقم الهوية','رقم الجوال']).casefold()]
        summaries=[]
        for c in rr:
            cc=[x for x in snap['contracts'] if x['client_id']==c['رقم العميل'] and x['disbursed']]
            summaries.append({'رقم العميل':c['رقم العميل'],'اسم العميل':c['اسم العميل'],'الجوال':c['رقم الجوال'],'الحالة':c['حالة العميل'],
             'التصنيف':c['التصنيف الائتماني'],'إجمالي التمويل':core.sar(sum(x['principal_h'] for x in cc)),
             'المتبقي':core.sar(sum(x['remaining_h'] for x in cc)),'الحد الائتماني':c['الحد الائتماني'],
             'مرات التأخير':sum(1 for p in snap['installments'] if p['client_id']==c['رقم العميل'] and p['disbursed'] and p['late_days']>0)})
        table(summaries,'لا يوجد عملاء مطابقون. أضف أول عميل من التبويب الآخر.')
        if not rr:return
        ids={c['رقم العميل']:c for c in rr};cid=st.selectbox('فتح ملف العميل',list(ids),format_func=lambda x:ids[x]['اسم العميل'],key='open_client')
        c=ids[cid];cs=[x for x in snap['contracts'] if x['client_id']==cid];parts=[p for p in snap['installments'] if p['client_id']==cid]
        tids={x['رقم العقد'] for x in cs};trx=[t for t in snap['transactions'] if t['رقم العقد'] in tids]
        st.markdown(f'<div class="profile"><strong>{escape(c["اسم العميل"])}</strong><p>{escape(str(c["رقم الجوال"]))} · {escape(str(c["حالة العميل"]))} · التصنيف {escape(str(c["التصنيف الائتماني"]))}</p></div>',unsafe_allow_html=True)
        metrics([('إجمالي المتبقي',fmt(sum(x['remaining_h'] for x in cs if x['disbursed']))),('المتأخر',fmt(sum(x['overdue_h'] for x in cs))),('العقود',str(len(cs)))])
        tabs=st.tabs(['بيانات العميل','عقود العميل','أقساطه وحركاته','المتابعات'])
        with tabs[0]:client_form(snap,c)
        with tabs[1]:table(contract_table(cs,client_names(snap)))
        with tabs[2]:
            table(schedule_table(parts));table(tx_table(trx))
        with tabs[3]:table([r for r in snap['followups'] if r['رقم العميل']==cid])
        st.download_button('تنزيل كشف حساب للطباعة',statement_html(c,cs,parts,trx),file_name=f'{cid}_statement.html',mime='text/html')

def investor_form(existing=None):
    old=existing or {}
    with st.form('investor-'+old.get('رقم المستثمر','new')):
        a,b=st.columns(2)
        name=a.text_input('اسم المستثمر *',value=old.get('اسم المستثمر',''))
        phone=b.text_input('جوال المستثمر *',value=old.get('رقم الجوال',''))
        identity=a.text_input('هوية المستثمر',value=old.get('رقم الهوية',''))
        state=b.selectbox('حالة المستثمر',core.INVESTOR_STATES,index=core.INVESTOR_STATES.index(old.get('الحالة','نشط')))
        notes=st.text_area('ملاحظات المستثمر',value=old.get('ملاحظات',''))
        if st.form_submit_button('حفظ المستثمر',type='primary'):
            data={'اسم المستثمر':name,'رقم الجوال':phone,'رقم الهوية':identity,'الحالة':state,'ملاحظات':notes}
            if existing:data['id']=old['رقم المستثمر']
            submit('edit_investor' if existing else 'add_investor',data)

def investors_page(wb,snap):
    view,new=st.tabs(['حسابات المستثمرين والتحويلات','إضافة مستثمر'])
    with new:investor_form()
    with view:
        rr=snap['investors']
        table([{'رقم المستثمر':r['رقم المستثمر'],'المستثمر':r['اسم المستثمر'],'الجوال':r['رقم الجوال'],'الحالة':r['الحالة'],
          'إجمالي التمويل':r['إجمالي التمويل'] or 0,'متبقي العقود':r['متبقي العقود'] or 0,
          'أموال محتجزة للتحويل':r['تحصيلات محتجزة'] or 0,'رسوم المكتب المحصلة':r['رسوم المكتب'] or 0} for r in rr],
          'لا يوجد مستثمرون. أضف المستثمر قبل إنشاء عقد مُدار له.')
        if not rr:return
        byid={r['رقم المستثمر']:r for r in rr};iid=st.selectbox('فتح حساب المستثمر',list(byid),format_func=lambda x:f"{byid[x]['اسم المستثمر']} · {x}")
        inv=byid[iid];managed=[c for c in snap['contracts'] if c['investor_id']==iid and c['deal_type']==core.DEAL_MANAGED]
        held=core.money(inv.get('تحصيلات محتجزة'))
        metrics([('المبلغ المحتجز للتحويل',fmt(held)),('متبقي عقوده',fmt(sum(c['remaining_h'] for c in managed if c['disbursed']))),
          ('ربحه المتوقع',fmt(sum(c['investor_expected_profit_h'] for c in managed if c['disbursed']))),
          ('رسوم المكتب المحصلة',f"{float(inv.get('رسوم المكتب') or 0):,.2f}"),('عدد العقود',str(len(managed)))])
        tabs=st.tabs(['عقود المستثمر','تحويل أمواله','تعديل بياناته'])
        with tabs[0]:table(contract_table(managed,client_names(snap)))
        with tabs[1]:
            if held>0:
                with st.form('investor-transfer-'+iid):
                    options=['']+[c['رقم العقد'] for c in managed if c['disbursed']]
                    contract=st.selectbox('العقد المرتبط (اختياري)',options,format_func=lambda x:x or 'تحويل عام للمستثمر')
                    amount=st.number_input('المبلغ المحول للمستثمر · ر.س',min_value=.01,max_value=core.sar(held),value=core.sar(held),step=100.)
                    when=st.date_input('تاريخ التحويل',value=core.today(),max_value=core.today())
                    reference=st.text_input('مرجع التحويل');notes=st.text_area('ملاحظات التحويل')
                    st.caption('يخصم التحويل من الحساب البنكي ومن الأموال المحتجزة للمستثمر، ولا يخصم من سيولة المكتب الحرة.')
                    if st.form_submit_button('تأكيد التحويل للمستثمر',type='primary'):
                        submit('transaction',dict(kind='تحويل للمستثمر',amount=amount,when=when,contract=contract,investor=iid,reference=reference,notes=notes))
            else:st.info('لا توجد أموال محتجزة مستحقة للتحويل حاليًا.')
        with tabs[2]:investor_form(inv)
        st.subheader('حركات المستثمر')
        table(tx_table([t for t in snap['transactions'] if t.get('رقم المستثمر')==iid]))

def contract_form(wb,snap,existing=None):
    old=existing or {};cid=choose_client(snap,f"contract-client-{old.get('رقم العقد','new')}",True)
    if not cid:return
    if existing: cid=old['client_id'];st.caption(f"عميل العقد: {client_names(snap)[cid]}")
    params=core.settings(wb)
    current=old.get('deal_type') or old.get('نوع الصفقة') or core.DEAL_SELF
    mode=st.selectbox('دورك في الصفقة',core.DEAL_TYPES,index=core.DEAL_TYPES.index(current),key='deal-mode-'+old.get('رقم العقد','new'))
    investor=''
    if mode==core.DEAL_MANAGED:
        active=[r for r in snap['investors'] if r['الحالة']=='نشط']
        if not active:st.warning('أضف مستثمرًا نشطًا من صفحة المستثمرين أولًا.');return
        byid={r['رقم المستثمر']:r for r in active};default=old.get('investor_id') or old.get('رقم المستثمر')
        investor=st.selectbox('المستثمر مالك كامل العقد والتحصيلات',list(byid),index=list(byid).index(default) if default in byid else 0,format_func=lambda x:f"{byid[x]['اسم المستثمر']} · {x}")
    with st.form(f"contract-{old.get('رقم العقد','new')}"):
        contract_number=st.text_input('رقم العقد (اختياري؛ يولّد تلقائيًا عند تركه فارغًا)',value=old.get('رقم العقد',''),disabled=bool(existing))
        a,b=st.columns(2)
        purchase=a.number_input('تكلفة شراء السلعة الفعلية · ر.س',min_value=.01,value=core.sar(old['purchase_cost_h']) if existing else 5000.,step=10.)
        cash_price=b.number_input('سعر الكاش / أصل التمويل · ر.س',min_value=.01,value=core.sar(old['cash_price_h']) if existing else 5000.,step=100.)
        funding=a.number_input('المبلغ المحول للشراء · ر.س',min_value=.01,value=core.sar(old['funding_h']) if existing else 5000.,step=100.,disabled=mode==core.DEAL_SELF,
          help='في التمويل الذاتي يعتمد التطبيق تكلفة الشراء نفسها. في عقد المستثمر يظهر الفائض غير المستخدم كمبلغ مستحق له.')
        total=b.number_input('إجمالي البيع بالتقسيط · ر.س',min_value=.01,value=core.sar(old['total_h']) if existing else round(5000*(1+float(params['default_profit_rate'])),2),step=100.)
        term=b.number_input('مدة العقد · شهر',min_value=1,max_value=120,value=int(old.get('مدة العقد') or params['default_months']))
        signed=b.date_input('تاريخ العقد',value=core.day(old.get('تاريخ العقد')) or core.today(),max_value=core.today())
        first=a.date_input('أول قسط',value=core.day(old.get('أول قسط')) or core.add_month(core.today(),1))
        office_name=st.text_input('المكتب المنفذ',value=old.get('المكتب المنفذ') or '',disabled=mode!=core.DEAL_EXTERNAL)
        paid_fee=a.number_input('رسوم مكتب مدفوعة · ر.س',min_value=0.,value=float(old.get('رسوم مكتب مدفوعة') or 0),step=50.,disabled=mode==core.DEAL_MANAGED)
        fixed_fee=b.number_input('رسوم المكتب الثابتة من المستثمر عند التنفيذ · ر.س',min_value=0.,value=float(old.get('رسوم مكتب ثابتة') or 0),step=50.,disabled=mode!=core.DEAL_MANAGED)
        other=a.number_input('مصاريف أخرى على حسابك · ر.س',min_value=0.,value=float(old.get('مصاريف أخرى') or 0),step=10.)
        fee=b.number_input('رسوم إضافية على العميل خارج الأقساط · ر.س',min_value=0.,value=float(old.get('رسوم العقد') or 0),step=50.)
        notes=st.text_area('ملاحظات العقد',value=old.get('ملاحظات') or '')
        if mode==core.DEAL_MANAGED:
            preview=fixed_fee+fee-other;investor_profit=total-purchase-fixed_fee
            st.info(f'ربح المكتب المتوقع: {preview:,.2f} ريال فقط · ربح المستثمر المتوقع: {investor_profit:,.2f} ريال · كامل أقساط {total:,.2f} ريال ملك للمستثمر.')
        else:
            preview=total+fee-purchase-paid_fee-other;surplus=max(0.,funding-purchase) if mode==core.DEAL_EXTERNAL else 0.
            st.info(f'صافي ربحك المتوقع: {preview:,.2f} ريال · فرق الشراء: {cash_price-purchase:,.2f} ريال'+(f' · فائض متوقع من المكتب: {surplus:,.2f} ريال' if surplus else ''))
        st.caption('إنشاء العقد يحفظ مسودة فقط. التنفيذ المالي يتم من زر تنفيذ الصفقة بعد مراجعة جميع القيم.')
        if st.form_submit_button('حفظ مسودة العقد',type='primary'):
            data={'رقم العقد':contract_number,'رقم العميل':cid,'مبلغ التمويل':cash_price,'سعر الكاش':cash_price,'تكلفة الشراء الفعلية':purchase,
              'مبلغ التمويل المحول':purchase if mode==core.DEAL_SELF else funding,'إجمالي السداد':total,'مدة العقد':int(term),'تاريخ العقد':signed,
              'أول قسط':first,'رسوم العقد':fee,'نوع الصفقة':mode,'رقم المستثمر':investor,'المكتب المنفذ':office_name,
              'رسوم مكتب مدفوعة':paid_fee if mode!=core.DEAL_MANAGED else 0,'رسوم مكتب ثابتة':fixed_fee if mode==core.DEAL_MANAGED else 0,
              'مصاريف أخرى':other,'ملاحظات':notes}
            if existing:data['id']=old['رقم العقد']
            submit('edit_contract' if existing else 'add_contract',data)

def activation_form(c,key):
    if c['deal_type']==core.DEAL_MANAGED:
        st.info(f"سيُسجَّل إيداع المستثمر {core.sar(c['funding_h']):,.2f} ريال، وشراء السلعة {core.sar(c['purchase_cost_h']):,.2f} ريال، ورسوم المكتب {core.sar(c['fixed_office_fee_h']):,.2f} ريال، ومصاريف المكتب {core.sar(c['other_cost_h']):,.2f} ريال. فائض التمويل يبقى مستحقًا للمستثمر.")
    else:
        st.info(f"سيُخصم {core.sar(c['cash_outflow_h']):,.2f} ريال من سيولتك، وهو كامل المبلغ الخارج لتنفيذ الصفقة.")
    with st.form(key):
        when=st.date_input('تاريخ تنفيذ الصفقة',value=core.today(),max_value=core.today())
        collector=st.text_input('المسؤول',value='إبراهيم')
        reference=st.text_input('مرجع التحويل أو فاتورة الشراء')
        notes=st.text_input('ملاحظات التنفيذ')
        if st.form_submit_button('تأكيد تنفيذ الصفقة وحفظ الحركات',type='primary'):
            submit('activate_contract',dict(ident=c['رقم العقد'],when=when,collector=collector,reference=reference,notes=notes))

def payment_form(c,kind,key):
    principal=core.sar(c['principal_h']);remaining=core.sar(c['remaining_h'])
    if kind=='صرف تمويل':st.info(f'سيُخصم {principal:,.2f} ريال من السيولة عند الصرف.')
    if kind in core.RECEIPTS:
        owner=' هذه الدفعة ملك للمستثمر وستضاف إلى أمواله المحتجزة، لا إلى سيولة المكتب.' if c['deal_type']==core.DEAL_MANAGED else ''
        st.info(f'المتبقي على العقد: {remaining:,.2f} ريال. يوزع السداد على أقدم الأقساط أولًا.{owner}')
    with st.form(key):
        when=st.date_input('تاريخ العملية',value=core.today(),max_value=core.today())
        if kind=='صرف تمويل': amount=principal
        else:
            initial=(remaining if kind=='سداد مبكر' else min(remaining,core.sar(c['total_h']//int(c['مدة العقد'])))) if kind in core.RECEIPTS else float(c.get('رسوم العقد') or 0)
            amount=st.number_input('المبلغ المدفوع فعليًا · ر.س',min_value=.01,value=max(.01,initial),step=100.)
        discount=0.
        if kind=='سداد مبكر':
            discount=st.number_input('الخصم المتفق عليه · ر.س',min_value=0.,value=0.,step=100.)
            st.caption('النقد المدفوع + الخصم يجب أن يساويا المتبقي. لا يحدد التطبيق خصمًا تلقائيًا.')
        collector=st.text_input('المحصل / المسؤول',value='إبراهيم')
        reference=st.text_input('مرجع التحويل أو الإيصال',help='يمنع تكرار المرجع في الحركات النشطة.')
        notes=st.text_input('ملاحظات العملية')
        if st.form_submit_button(f'تأكيد {kind} وحفظه',type='primary'):
            submit('transaction',dict(kind=kind,amount=amount,when=when,contract=c['رقم العقد'],collector=collector,reference=reference,notes=notes,discount=discount))

def surplus_refund_form(c):
    due=core.sar(c['surplus_due_h'])
    st.info(f"المبلغ المحول للمكتب المنفذ: {core.sar(c['funding_h']):,.2f} ريال · تكلفة الشراء: {core.sar(c['purchase_cost_h']):,.2f} ريال · الفائض المتبقي لك: {due:,.2f} ريال.")
    if due<=0:
        st.success('تم استرداد كامل فائض التمويل المتوقع.')
        return
    with st.form('surplus-refund-'+c['رقم العقد']):
        amount=st.number_input('الفائض المسترد فعليًا · ر.س',min_value=.01,max_value=due,value=due,step=10.)
        when=st.date_input('تاريخ استرداد الفائض',value=core.today(),max_value=core.today())
        reference=st.text_input('مرجع الاسترداد');notes=st.text_area('ملاحظات الاسترداد')
        if st.form_submit_button('تسجيل استرداد الفائض',type='primary'):
            submit('transaction',dict(kind='استرداد فائض تمويل',amount=amount,when=when,contract=c['رقم العقد'],reference=reference,notes=notes))

def contracts_page(wb,snap):
    names=client_names(snap)
    view,new=st.tabs(['سجل العقود وإدارة السداد','إنشاء عقد'])
    with new:contract_form(wb,snap)
    with view:
        q=st.text_input('بحث برقم العقد أو اسم العميل',key='contract_search')
        status=st.selectbox('حالة العقد',['الكل','مسودة','منتظم','متأخر','مكتمل','مسدد مبكرًا'],key='contract_filter')
        cc=[c for c in snap['contracts'] if (status=='الكل' or c['status']==status) and q.casefold() in f"{c['رقم العقد']} {names.get(c['client_id'],'')}".casefold()]
        table(contract_table(cc,names))
        if not cc:return
        byid={c['رقم العقد']:c for c in cc};cid=st.selectbox('العقد',list(byid),format_func=lambda x:f"{x} · {names[byid[x]['client_id']]}",key='select_contract')
        c=byid[cid]
        owner='المكتب'
        if c['deal_type']==core.DEAL_MANAGED:
            inv=investor_names(snap).get(c['investor_id'],c['investor_id']);owner=f'المستثمر: {inv}'
        st.caption(f"نوع الصفقة: {c['deal_type']} · مالك كامل الأقساط: {owner}")
        contract_metrics=[('تكلفة الشراء',fmt(c['purchase_cost_h'])),('سعر الكاش',fmt(c['cash_price_h'])),('إجمالي التقسيط',fmt(c['total_h'])),
          ('المحصل',fmt(c['paid_h'])),('المتبقي',fmt(c['remaining_h'])),('المتأخر',fmt(c['overdue_h'])),
          ('ربح المكتب المتوقع',fmt(c['office_expected_profit_h']))]
        if c['deal_type']==core.DEAL_MANAGED:
            contract_metrics.extend([('ربح المستثمر المتوقع',fmt(c['investor_expected_profit_h'])),('عائد المستثمر المتوقع',f"{float(c['roi'])*100:,.2f}%")])
        else:
            contract_metrics.append(('عائدك المتوقع',f"{float(c['roi'])*100:,.2f}%"))
        metrics(contract_metrics)
        st.subheader('جدول الأقساط')
        table(schedule_table([p for p in snap['installments'] if p['رقم العقد']==cid]))
        if not c['disbursed']:
            a,b=st.tabs(['تنفيذ الصفقة','تعديل المسودة'])
            with a:activation_form(c,'activate-'+cid)
            with b:contract_form(wb,snap,c)
        else:
            labels=['سداد قسط / سداد جزئي','سداد مبكر','تحصيل رسوم العميل']
            if c['deal_type']==core.DEAL_EXTERNAL:labels.append('استرداد فائض التمويل')
            tabs=st.tabs(labels);t1,t2,t3=tabs[:3]
            with t1:
                if c['remaining_h']>0:payment_form(c,'سداد قسط','pay-'+cid)
                else:st.success('العقد مكتمل السداد.')
            with t2:
                if c['remaining_h']>0:payment_form(c,'سداد مبكر','settle-'+cid)
                else:st.success('لا يوجد مبلغ متبقٍ للتسوية.')
            with t3:
                paid=sum(core.money(t['المبلغ']) for t in snap['transactions'] if t['رقم العقد']==cid and t['نوع العملية']=='رسوم عقد' and t['حالة العملية']=='نشطة')
                st.caption(f"الرسوم المتفق عليها {float(c.get('رسوم العقد') or 0):,.2f} ريال · المحصل {core.sar(paid):,.2f} ريال")
                if core.money(c.get('رسوم العقد'))>paid:payment_form({**c,'رسوم العقد':core.sar(core.money(c.get('رسوم العقد'))-paid)},'رسوم عقد','fee-'+cid)
                else:st.info('لا توجد رسوم عقد متبقية.')
            if c['deal_type']==core.DEAL_EXTERNAL:
                with tabs[3]:surplus_refund_form(c)
        st.subheader('سجل حركات العقد')
        table(tx_table([r for r in snap['transactions'] if r['رقم العقد']==cid]))

def transactions_page(wb,snap):
    add,history=st.tabs(['حركة سيولة جديدة','سجل الحركات والتعديل'])
    with add:
        st.caption('تنفيذ الصفقات والأقساط ورسوم المستثمر وتحويلاته تسجل من صفحات العقود والمستثمرين. رأس المال والسحب والمصروف العام من هنا.')
        with st.form('cash_transaction'):
            kind=st.selectbox('نوع العملية',['إضافة رأس مال','سحب رأس مال','مصروف'])
            when=st.date_input('تاريخ الحركة',value=core.today(),max_value=core.today())
            amount=st.number_input('قيمة الحركة · ر.س',min_value=.01,value=1000.,step=100.)
            reference=st.text_input('المرجع',key='cash_ref');notes=st.text_area('ملاحظات',key='cash_notes')
            if st.form_submit_button('حفظ حركة السيولة',type='primary'):submit('transaction',dict(kind=kind,amount=amount,when=when,reference=reference,notes=notes))
    with history:
        show_void=st.checkbox('إظهار الحركات الملغاة')
        q=st.text_input('بحث في الحركات',key='tx_search')
        rr=[r for r in snap['transactions'] if (show_void or r['حالة العملية']=='نشطة') and q.casefold() in ' '.join(str(r.get(k) or '') for k in ['رقم العملية','رقم العقد','المرجع','ملاحظات','نوع العملية']).casefold()]
        table(tx_table(list(reversed(rr))))
        if not rr:return
        byid={r['رقم العملية']:r for r in rr};ident=st.selectbox('اختر حركة لإدارتها',list(reversed(byid)),format_func=lambda x:f"{byid[x]['نوع العملية']} · {float(byid[x]['المبلغ']):,.2f} · {x}")
        r=byid[ident]
        st.caption('التعديل يحفظ نسخة بديلة ويلغي القديمة مع بقاء السجل. الإلغاء والاستعادة يعيدان حساب الأقساط والسيولة، وتُرفض العملية إذا كسرت حركة لاحقة.')
        a,b=st.tabs(['تعديل الحركة','إلغاء / استعادة'])
        with a:
            if r['حالة العملية']=='نشطة':
                with st.form('edit-tx-'+ident):
                    amount=st.number_input('المبلغ المصحح',min_value=.01,value=float(r['المبلغ']))
                    when=st.date_input('التاريخ المصحح',value=core.day(r['التاريخ']),max_value=core.today())
                    collector=st.text_input('المحصل',value=r['المحصل'] or '')
                    ref=st.text_input('مرجع الحركة',value=r['المرجع'] or '')
                    note=st.text_area('ملاحظات التصحيح',value=r['ملاحظات'] or '')
                    disc=st.number_input('خصم التسوية',min_value=0.,value=float(r['خصم التسوية'] or 0),disabled=r['نوع العملية']!='سداد مبكر')
                    if st.form_submit_button('حفظ التصحيح',type='primary'):
                        submit('edit_transaction',{'id':ident,'replacement':dict(kind=r['نوع العملية'],amount=amount,when=when,contract=r['رقم العقد'] or '',investor=r.get('رقم المستثمر') or '',collector=collector,reference=ref,notes=note,discount=disc)})
            else:st.info('استعد الحركة أولًا لتعديلها.')
        with b:
            verb='إلغاء' if r['حالة العملية']=='نشطة' else 'استعادة'
            ok=st.checkbox(f'تأكيد {verb} هذه الحركة',key='confirm-'+ident)
            if st.button(f'{verb} الحركة وإعادة الحساب',disabled=not ok):submit('void_transaction' if verb=='إلغاء' else 'restore_transaction',{'id':ident})

def followups_page(wb,snap):
    names=client_names(snap)
    view,new=st.tabs(['قائمة المتابعة','تسجيل تواصل / وعد سداد'])
    with new:
        cid=choose_client(snap,'followup-client')
        if cid:
            cs=[c for c in snap['contracts'] if c['client_id']==cid];opts=['']+[c['رقم العقد'] for c in cs]
            with st.form('followup'):
                contract=st.selectbox('العقد المرتبط (اختياري)',opts,format_func=lambda x:x or 'متابعة عامة للعميل')
                a,b=st.columns(2)
                when=a.date_input('تاريخ التواصل',value=core.today(),max_value=core.today())
                method=b.selectbox('وسيلة التواصل',['اتصال','واتساب','رسالة نصية','زيارة','أخرى'])
                result=st.text_input('نتيجة التواصل')
                has_promise=st.checkbox('يوجد وعد سداد',value=False)
                promise=a.date_input('تاريخ وعد السداد',value=core.today()+timedelta(days=3))
                amount=b.number_input('المبلغ الموعود',min_value=0.,value=0.,step=100.)
                next_date=st.date_input('المتابعة القادمة',value=core.today()+timedelta(days=3))
                notes=st.text_area('ملاحظات المتابعة')
                if st.form_submit_button('حفظ المتابعة',type='primary'):
                    submit('add_followup',{'رقم العميل':cid,'رقم العقد':contract,'تاريخ التواصل':when,'وسيلة التواصل':method,
                       'نتيجة التواصل':result,'وعد السداد':promise if has_promise else None,'المبلغ الموعود':amount if has_promise else 0,
                       'المتابعة القادمة':next_date,'ملاحظات':notes})
    with view:
        mode=st.selectbox('عرض المتابعات',['المفتوحة','المستحقة اليوم وما قبله','المكتملة','الكل'])
        rr=[r for r in snap['followups'] if mode=='الكل' or (mode=='المكتملة' and r['الحالة']=='مكتملة') or (mode=='المفتوحة' and r['الحالة']=='مفتوحة') or (mode=='المستحقة اليوم وما قبله' and r['الحالة']=='مفتوحة' and r['المتابعة القادمة'] and core.day(r['المتابعة القادمة'])<=core.today())]
        table([{'العميل':names.get(r['رقم العميل'],''),**r} for r in rr])
        if rr:
            records={r['رقم المتابعة']:r for r in rr}
            with st.expander('تعديل نتيجة التواصل أو وعد السداد'):
                chosen=st.selectbox('المتابعة المراد تعديلها',list(records),format_func=lambda x:f"{names.get(records[x]['رقم العميل'],'')} · {records[x]['نتيجة التواصل']}",key='edit_followup_id')
                old=records[chosen]
                with st.form('edit-followup-'+chosen):
                    result=st.text_input('نتيجة التواصل المحدثة',value=old['نتيجة التواصل'] or '')
                    has_promise=st.checkbox('الاحتفاظ بوعد سداد',value=bool(old['وعد السداد']))
                    promised=st.date_input('موعد الوعد الجديد',value=core.day(old['وعد السداد']) or core.today())
                    amount=st.number_input('المبلغ الموعود الجديد',min_value=0.,value=float(old['المبلغ الموعود'] or 0))
                    next_date=st.date_input('موعد المتابعة الجديد',value=core.day(old['المتابعة القادمة']) or core.today())
                    state=st.selectbox('حالة المتابعة',['مفتوحة','مكتملة'],index=0 if old['الحالة']=='مفتوحة' else 1)
                    notes=st.text_area('ملاحظات محدثة',value=old['ملاحظات'] or '')
                    if st.form_submit_button('حفظ تعديل المتابعة'):
                        submit('edit_followup',{'id':chosen,'نتيجة التواصل':result,'وعد السداد':promised if has_promise else None,
                            'المبلغ الموعود':amount if has_promise else 0,'المتابعة القادمة':next_date,'الحالة':state,'ملاحظات':notes})
        pending=[r for r in rr if r['الحالة']=='مفتوحة']
        if pending:
            byid={r['رقم المتابعة']:r for r in pending}
            fid=st.selectbox('المتابعة',list(byid),format_func=lambda x:f"{names.get(byid[x]['رقم العميل'],'')} · {byid[x]['نتيجة التواصل']}")
            if st.button('تحديد المتابعة كمكتملة'):submit('close_followup',{'id':fid})

def reports(wb,snap):
    a,b=st.columns(2)
    start=a.date_input('من تاريخ',value=core.today().replace(day=1));end=b.date_input('إلى تاريخ',value=core.today())
    if start>end:st.error('بداية الفترة بعد نهايتها.');return
    txs=[r for r in snap['transactions'] if r['حالة العملية']=='نشطة' and start<=core.day(r['التاريخ'])<=end]
    sums={kind:sum(core.money(r['المبلغ']) for r in txs if r['نوع العملية']==kind) for kind in core.CASH_TYPES}
    contracts={c['رقم العقد']:c for c in snap['contracts']}
    managed_receipts=sum(core.money(r['المبلغ']) for r in txs if r['نوع العملية'] in core.RECEIPTS and contracts.get(r['رقم العقد'],{}).get('deal_type')==core.DEAL_MANAGED)
    office_receipts=sum(core.money(r['المبلغ']) for r in txs if r['نوع العملية'] in core.RECEIPTS and contracts.get(r['رقم العقد'],{}).get('deal_type')!=core.DEAL_MANAGED)
    office_fees=sums['رسوم مكتب']+sums['رسوم عقد']
    metrics([('تحصيل عقود المكتب',fmt(office_receipts)),('تحصيلات مملوكة للمستثمرين',fmt(managed_receipts)),
      ('دخل الرسوم للمكتب',fmt(office_fees)),('تحويلات للمستثمرين',fmt(sums['تحويل للمستثمر'])),
      ('تنفيذ صفقات من سيولتك',fmt(sums['صرف تمويل'])),('السيولة الحرة الحالية',fmt(snap['cash_h']))])
    st.caption('تحصيل عقد المستثمر يظهر منفصلًا لأنه ملك للمستثمر بالكامل. الرسوم الثابتة ورسوم العميل الإضافية فقط تدخل للمكتب.')
    if txs:
        frame=pd.DataFrame([{'الشهر':core.day(t['التاريخ']).strftime('%Y-%m'),'أثر الحركة':core.sar(core.money(t['المبلغ'])*core.CASH_TYPES[t['نوع العملية']])} for t in txs]).groupby('الشهر').sum()
        st.bar_chart(frame,color='#42c6bd')
    st.subheader('تفصيل التدفقات')
    table([{'نوع العملية':k,'إجمالي المبلغ':core.sar(v),'الأثر على الحساب البنكي':core.sar(v*core.CASH_TYPES[k])} for k,v in sums.items() if v])
    st.subheader('ربحية العقود المتوقعة')
    names=client_names(snap);inames=investor_names(snap)
    table([{'العقد':c['رقم العقد'],'العميل':names.get(c['client_id'],''),'نوع الصفقة':c['deal_type'],
      'المستثمر':inames.get(c['investor_id'],'') if c['investor_id'] else '',
      'تكلفة الشراء':core.sar(c['purchase_cost_h']),'إجمالي التقسيط':core.sar(c['total_h']),
      'ربح المكتب المتوقع':core.sar(c['office_expected_profit_h']),'ربح المستثمر المتوقع':core.sar(c['investor_expected_profit_h']) if c['deal_type']==core.DEAL_MANAGED else 0,
      'العائد المتوقع %':round(float(c['roi'])*100,2),'حالة العقد':c['status']} for c in snap['contracts']])
    st.subheader('المستحقات الحالية حسب العميل')
    table([{'العميل':names.get(c['client_id'],''),'العقد':c['رقم العقد'],'نوع الصفقة':c['deal_type'],'المتبقي':core.sar(c['remaining_h']),'المتأخر':core.sar(c['overdue_h']),'الحالة':c['status']} for c in snap['contracts'] if c['disbursed'] and c['remaining_h']])
    st.subheader('الأقساط القادمة خلال 30 يومًا')
    table(schedule_table([p for p in snap['installments'] if p['disbursed'] and p['remaining_h'] and core.today()<=p['due']<=core.today()+timedelta(days=30)],names))

def settings_page(wb,snap):
    st.subheader('إعدادات العقود الجديدة')
    params=core.settings(wb)
    with st.form('settings'):
        term=st.number_input('المدة الافتراضية · شهر',min_value=1,max_value=120,value=int(params['default_months']))
        profit=st.number_input('نسبة الربح الإجمالية الافتراضية %',min_value=0.,max_value=1000.,value=float(params['default_profit_rate'])*100,step=5.)
        st.caption('هذه قيم مساعدة للمسودات الجديدة. إجمالي السداد المدخل في العقد هو المعتمد. التعديل لا يغير العقود المحفوظة.')
        if st.form_submit_button('حفظ الإعدادات'):submit('settings',{'default_months':term,'default_profit_rate':profit/100})
    st.subheader('ملف Excel والنسخ الاحتياطية')
    st.download_button('تنزيل ملف Excel الحالي',st.session_state['_raw'],file_name='Installment_System.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.caption('الإكسل هو مصدر البيانات. التطبيق يعيد الحساب من السجلات الأصلية دون الاعتماد على نتيجة صيغ مخزنة قديمة.')
    st.info('رصيد البداية المنقول من نوشن: ١٠٠٬٠٠٠ ريال بتاريخ ١ يوليو ٢٠٢٦. يمكن تصحيحه من سجل الحركات المالية. لا توجد سجلات عملاء أو عقود منقولة.')
    with st.expander('استيراد نسخة Excel للنظام / استعادة نسخة احتياطية'):
        uploaded=st.file_uploader('ملف نظام التقسيط فقط',type=['xlsx'],key='restore_book')
        checked=False
        if uploaded:
            try:
                candidate=core.load(uploaded.getvalue());ss=core.snapshot(candidate)
                st.write(f"النسخة تحتوي {len(ss['clients'])} عميل و{len(ss['contracts'])} عقد. السيولة {fmt(ss['cash_h'])} ريال.")
                checked=True
            except Exception as exc:st.error(f'النسخة غير متوافقة: {exc}')
        ok=st.checkbox('أريد استبدال المصدر بهذه النسخة بعد المراجعة')
        if st.button('استعادة النسخة',disabled=not(checked and ok)):
            try:
                raw,tag=st.session_state['_store'].replace(st.session_state['_tag'],uploaded.getvalue())
                st.session_state['_raw']=raw;st.session_state['_tag']=tag;notify('تمت الاستعادة.');st.rerun()
            except Exception as exc:st.error(str(exc))
    if isinstance(st.session_state['_store'],LocalStore):
        folder=st.session_state['_store'].path.parent/'backups';backups=sorted(folder.glob('*.xlsx'),reverse=True) if folder.exists() else []
        if backups:
            chosen=st.selectbox('نسخ الحفظ التلقائي',backups,format_func=lambda p:p.name)
            st.download_button('تنزيل النسخة المختارة',chosen.read_bytes(),file_name=chosen.name)
    st.subheader('قواعد الحساب')
    st.write('• تكلفة الشراء هي ما دفع فعليًا، وسعر الكاش هو أصل التمويل، وإجمالي التقسيط هو كامل الأقساط.\n\n• في «إدارة تمويل مستثمر» يملك المستثمر كامل تحصيلات الأقساط، ويملك المكتب الرسوم الثابتة ورسوم العميل الإضافية فقط.\n\n• أموال المستثمر المحتجزة لا تدخل في السيولة الحرة للمكتب، ولا يسمح النظام باستخدامها في مصروفات المكتب.\n\n• إجمالي السداد والمدة يحددان مبلغ الأقساط، وفرق التقريب يضاف للقسط الأخير.\n\n• السداد يوزع على الأقدم، ويقبل جزءًا من القسط أو أكثر من قسط.\n\n• التسوية المبكرة تتطلب إدخال النقد والخصم المتفق عليهما.\n\n• التأخير يحسب حتى اكتمال السداد، ويظل ظاهرًا بعد السداد المتأخر.\n\n• التصنيف الائتماني يدوي، والمتبقي من التمويل يحسب بالتناسب مع المتبقي من إجمالي السداد.')
    st.caption('مطابقة التقسيم والحقول الظاهرة في نوشن. صيغ التصنيف والأزرار الداخلية لم تكن متاحة من الرابط العام.')
    with st.expander('سجل التغييرات وأرقام الطلبات'):table(list(reversed(core.rows(wb,'سجل التغييرات'))))

def configure_store():
    mode=os.getenv('INSTALLMENT_STORAGE_MODE',str(cfg('storage','mode','local')))
    if mode not in ['local','onedrive']:st.error('وضع الحفظ يجب أن يكون local أو onedrive.');st.stop()
    if mode=='onedrive':
        import microsoft_auth as auth
        if not auth.cloud_mode_enabled():st.error('أضف إعدادات Microsoft في Secrets لتفعيل OneDrive. راجع دليل التشغيل.');st.stop()
        token=auth.acquire_microsoft_token()
        if not token:st.stop()
        config=auth.microsoft_config()
        store=OneDriveStore(token,config['onedrive_folder'],config['workbook_name'])
        st.sidebar.success('الحفظ: OneDrive')
        if st.sidebar.button('تسجيل الخروج'):
            for key in list(st.session_state):del st.session_state[key]
            st.rerun()
    else:
        password=str(cfg('app','access_password',''))
        if password and not st.session_state.get('authenticated'):
            entered=st.text_input('كلمة مرور التطبيق',type='password')
            if st.button('دخول'):
                if hmac.compare_digest(entered.encode(),password.encode()):st.session_state['authenticated']=True;st.rerun()
                else:st.error('كلمة المرور غير صحيحة.')
            st.stop()
        store=LocalStore(Path(os.getenv('INSTALLMENT_WORKBOOK',str(cfg('storage','workbook',ROOT/'Installment_System.xlsx')))))
        st.sidebar.info('الحفظ: ملف Excel محلي')
        st.sidebar.caption('استخدم OneDrive للعمل من الجوال والكمبيوتر على ملف واحد.')
    identity=f'{mode}:{getattr(store,"path",getattr(store,"folder",""))}'
    if '_raw' not in st.session_state or st.session_state.get('_identity')!=identity:
        try:raw,tag=store.read()
        except Exception as exc:st.error(str(exc));st.stop()
        st.session_state['_raw']=raw;st.session_state['_tag']=tag;st.session_state['_identity']=identity
    st.session_state['_store']=store
    if st.sidebar.button('↻ تحديث من Excel',width='stretch'):
        try:
            raw,tag=store.read();st.session_state['_raw']=raw;st.session_state['_tag']=tag
            # An uncertain request may already exist remotely; core.perform recognizes its id.
            notify('تم تحميل أحدث نسخة من المصدر.');st.rerun()
        except Exception as exc:st.error(str(exc))

def main():
    st.set_page_config(page_title='إدارة التقسيط | إبراهيم',page_icon='📊',layout='wide')
    theme()
    st.sidebar.markdown('### 📊 إدارة التقسيط')
    st.sidebar.caption('IBRAHIM · INSTALLMENT SYSTEM')
    configure_store()
    page=st.sidebar.radio('الصفحات',PAGES,key='nav',label_visibility='collapsed')
    st.sidebar.divider();st.sidebar.caption(f'الإصدار {core.VERSION} · جميع المبالغ بالريال')
    st.markdown('<div class="eyebrow">إدارة التقسيط</div>',unsafe_allow_html=True)
    st.title(page)
    st.caption('العملاء والعقود والتحصيلات في مكان واحد')
    if 'notice' in st.session_state:st.success(st.session_state.pop('notice'))
    try:
        wb=core.load(st.session_state['_raw']);snap=core.snapshot(wb)
        {'📊 ملخص المحفظة':dashboard,'👤 عملاء التقسيط':clients_page,'🤝 المستثمرون':investors_page,'📄 العقود':contracts_page,'💰 الحركات المالية':transactions_page,
          '📞 متابعة العملاء':followups_page,'📈 التقارير':reports,'⚙️ الإعدادات والنسخ':settings_page}[page](wb,snap)
    except Exception as exc:st.error(f'تعذر قراءة بيانات النظام: {exc}')

if __name__=='__main__':main()
