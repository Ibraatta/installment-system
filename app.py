"""Installment System 2.0. One engine and UI entrypoint; XLSX is authoritative.

Money is calculated in integer halalas. Transactions are replayed, not trusted
from cached formula values. Executed terms are immutable; corrections void a
whole posting group and replay every subsequent group before publication.
"""
from __future__ import annotations
from calendar import monthrange
from copy import deepcopy
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path
from uuid import uuid4
from zoneinfo import ZoneInfo
import hashlib
import json
import os
import tempfile
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st
from filelock import FileLock
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.properties import CalcProperties
from openpyxl.utils import get_column_letter

VERSION = '2.0.0'
FORMAT = 'IBRAHIM_INSTALLMENTS_V2'
ROOT = Path(__file__).resolve().parent
TZ = ZoneInfo('Asia/Riyadh')
SELF, EXTERNAL, MANAGED = 'تمويل ذاتي', 'استثمار عبر مكتب', 'إدارة تمويل مستثمر'
MODES = [SELF, EXTERNAL, MANAGED]
HEADERS = {
 'العملاء': ['المعرف','الاسم','الجوال','الهوية','المستوى','الحالة','الحد الائتماني','العنوان','جهة العمل','الدخل','الكفيل','جوال الكفيل','ملاحظات'],
 'المستثمرون': ['المعرف','الاسم','الجوال','الهوية','الحالة','ملاحظات'],
 'المكاتب': ['المعرف','الاسم','الجوال','الحالة','ملاحظات'],
 'العقود': ['المعرف','العميل','النوع','المستثمر','المكتب','تاريخ العقد','أول استحقاق','الأشهر','تكلفة الشراء','سعر الكاش','إجمالي التقسيط','التمويل المحول','رسوم مكتب مدفوعة','رسوم مكتب ثابتة','رسوم العميل','مصاريف أخرى','ملاحظات','فرق الشراء','تكلفة المستثمر','ربح المستثمر المتوقع','ربح المكتب المتوقع','عائد المستثمر','المحصل','الخصم','المتبقي','الحالة','فائض لدى المكتب','تاريخ التنفيذ'],
 'الأقساط': ['المعرف','العقد','التسلسل','الاستحقاق','المبلغ','المحصل','الخصم','المتبقي','الحالة','أيام التأخير','تاريخ السداد'],
 'الحركات': ['المعرف','المجموعة','التاريخ','الوقت','النوع','العقد','المستثمر','المبلغ','أثر البنك','أثر أموال المستثمر','خصم التسوية','المرجع','ملاحظات','الحالة','رقم الطلب','الرصيد البنكي','أموال المستثمرين','السيولة الحرة'],
 'المتابعات': ['المعرف','العميل','العقد','تاريخ التواصل','الوسيلة','النتيجة','وعد السداد','المبلغ الموعود','المتابعة القادمة','الحالة','ملاحظات'],
 'النظام': ['المفتاح','القيمة'],
}
INPUT_COLUMNS = {'العقود':17,'الحركات':15}
DEFAULTS = {'format':FORMAT,'version':VERSION,'default_months':16,'default_profit_rate':0.60}

class Error(ValueError): pass
class Conflict(Error): pass
def today(): return datetime.now(TZ).date()
def uid(prefix): return prefix+'-'+uuid4().hex[:12].upper()
def money(v):
    try:
        x=Decimal(str(v if v not in [None,''] else 0))
        if not x.is_finite(): raise ValueError()
        return int((x*100).quantize(Decimal('1'),rounding=ROUND_HALF_UP))
    except Exception: raise Error('أدخل مبلغًا رقميًا صالحًا.')
def sar(h): return float(Decimal(h)/100)
def dt(v):
    if isinstance(v,datetime): return v.date()
    if isinstance(v,date): return v
    if not v: return None
    try: return date.fromisoformat(str(v)[:10])
    except ValueError: raise Error('التاريخ غير صالح.')
def month(d,n):
    y,m=divmod(d.year*12+d.month-1+n,12);m+=1
    return date(y,m,min(d.day,monthrange(y,m)[1]))
def txt(v):
    v=str(v or '').strip()
    if len(v)>4000 or any(ord(c)<32 and c not in '\n\t\r' for c in v):raise Error('النص غير صالح أو طويل جدًا.')
    return v
def empty(): return {**{k:[] for k in HEADERS},'النظام':[{'المفتاح':k,'القيمة':v} for k,v in DEFAULTS.items()]}
def lookup(data,sheet,ident):
    r=next((x for x in data[sheet] if x['المعرف']==ident),None)
    if r is None:raise Error('السجل المرتبط غير موجود: '+str(ident))
    return r
def setting(data,k): return next((r['القيمة'] for r in data['النظام'] if r['المفتاح']==k),None)

def pricing(c):
    a={k:money(c.get(k)) for k in HEADERS['العقود'][8:16]}
    purchase=a['تكلفة الشراء'];cash=a['سعر الكاش'];total=a['إجمالي التقسيط'];fund=a['التمويل المحول']
    managed=c['النوع']==MANAGED
    cost=purchase+(a['رسوم مكتب ثابتة'] if managed else a['رسوم مكتب مدفوعة']+a['مصاريف أخرى'])
    investor_profit=total-cost
    office_profit=(a['رسوم مكتب ثابتة']+a['رسوم العميل']-a['مصاريف أخرى']) if managed else investor_profit+a['رسوم العميل']
    return {'purchase':purchase,'cash':cash,'total':total,'fund':fund,'paid_fee':a['رسوم مكتب مدفوعة'],
      'fee':a['رسوم مكتب ثابتة'],'client_fee':a['رسوم العميل'],'expenses':a['مصاريف أخرى'],
      'saving':cash-purchase,'surplus':fund-purchase if c['النوع']!=SELF else 0,
      'cost':cost,'investor_profit':investor_profit,'office_profit':office_profit,
      'roi':float(Decimal(investor_profit)/cost) if cost else 0}

def valid_contract(data,c):
    client=lookup(data,'العملاء',c['العميل'])
    if c['النوع'] not in MODES:raise Error('اختر نوع الصفقة.')
    p=pricing(c)
    if min(p['purchase'],p['cash'],p['fund'])<=0 or p['total']<p['cash']:raise Error('تحقق من تكلفة الشراء وسعر الكاش وإجمالي التقسيط.')
    if min(p['paid_fee'],p['fee'],p['client_fee'],p['expenses'])<0:raise Error('الرسوم والمصاريف لا تقبل السالب.')
    if c['النوع']==SELF and p['fund']!=p['purchase']:raise Error('التمويل الذاتي يساوي تكلفة الشراء الفعلية.')
    if c['النوع']!=SELF and p['fund']<p['purchase']:raise Error('المبلغ المحول أقل من تكلفة الشراء.')
    if c['النوع']==MANAGED:
        lookup(data,'المستثمرون',c['المستثمر'])
        if p['paid_fee']:raise Error('عقد المستثمر يستخدم رسوم المكتب الثابتة.')
    elif c.get('المستثمر') or p['fee']:raise Error('المستثمر والرسوم الثابتة يخصان العقود المدارة فقط.')
    if c['النوع']==EXTERNAL:lookup(data,'المكاتب',c['المكتب'])
    if c['النوع']!=EXTERNAL and c.get('المكتب'):raise Error('المكتب الخارجي يخص الاستثمار عبر مكتب.')
    n=c['الأشهر']
    if not isinstance(n,(int,float)) or int(n)!=n or not 1<=n<=120 or p['total']<n:raise Error('مدة العقد من شهر إلى 120 شهرًا.')
    if not dt(c['تاريخ العقد']) or dt(c['تاريخ العقد'])>today() or not dt(c['أول استحقاق']) or dt(c['أول استحقاق'])<dt(c['تاريخ العقد']):raise Error('تواريخ العقد غير صالحة.')
    return client,p

def snapshot(data,as_of=None):
    """Replay all groups in date order. Void is all-or-nothing for each group."""
    now=as_of or today();ids={};contracts={};parts=[];states={};ledger=[]
    for sheet in ['العملاء','المستثمرون','المكاتب','العقود','الحركات','المتابعات']:
        vals=[r.get('المعرف') for r in data[sheet]]
        if any(not x for x in vals) or len(set(vals))!=len(vals):raise Error('معرف مفقود أو مكرر في '+sheet)
    for sheet in ['العملاء','المستثمرون','المكاتب']:
        for r in data[sheet]:
            if not txt(r.get('الاسم')) or not txt(r.get('الجوال')) or r.get('الحالة') not in ['نشط','موقوف','مؤرشف']:raise Error('بيانات غير صالحة في '+sheet)
        identities=[txt(r.get('الهوية')) for r in data[sheet] if txt(r.get('الهوية'))]
        if len(set(identities))!=len(identities):raise Error('الهوية مكررة في '+sheet)
    for r in data['العملاء']:
        if r.get('المستوى') not in ['غير مصنف','A','B','C','D'] or min(money(r.get('الحد الائتماني')),money(r.get('الدخل')))<0:raise Error('تصنيف أو بيانات مالية للعميل غير صالحة.')
    for c in data['العقود']:
        _,p=valid_contract(data,c);cid=c['المعرف'];contracts[cid]=c
        n=int(c['الأشهر']);base=p['total']//n
        schedule=[{'id':cid+'-'+str(i+1).zfill(3),'contract':cid,'sequence':i+1,'due':month(dt(c['أول استحقاق']),i),
          'amount':base if i<n-1 else p['total']-base*(n-1),'paid':0,'waived':0,'paid_date':None} for i in range(n)]
        states[cid]={'c':c,'p':p,'parts':schedule,'active':False,'activated':None,'settled':False,'refund':0}
        parts.extend(schedule)
    bank=0;holdings={r['المعرف']:0 for r in data['المستثمرون']};refs=set();requests_seen={};groups={}
    for tr in data['الحركات']:
        if tr.get('الحالة') not in ['نشطة','ملغاة']:raise Error('حالة حركة غير صالحة.')
        if not dt(tr['التاريخ']) or dt(tr['التاريخ'])>today():raise Error('تاريخ الحركة غير صالح.')
        group=txt(tr.get('المجموعة'));req=txt(tr.get('رقم الطلب'))
        if not group or not req:raise Error('مجموعة الحركة ورقم الطلب مطلوبان.')
        if req in requests_seen and requests_seen[req]!=group:raise Error('رقم طلب مكرر.')
        requests_seen[req]=group
        groups.setdefault(group,[]).append(tr)
    ordered=sorted(groups.values(),key=lambda ts:(dt(ts[0]['التاريخ']),str(ts[0]['الوقت']),ts[0]['المجموعة']))
    for ts in ordered:
        if len({(str(t['التاريخ']),t['الوقت'],t['الحالة'],t['رقم الطلب'],t.get('العقد') or '') for t in ts})!=1:raise Error('مجموعة حركة غير مكتملة.')
        if ts[0]['الحالة']=='ملغاة':continue
        kinds=[t['النوع'] for t in ts];cid=ts[0].get('العقد') or '';s=states.get(cid);c=contracts.get(cid)
        if cid and not s:raise Error('حركة مرتبطة بعقد مفقود.')
        if c and dt(ts[0]['التاريخ'])<dt(c['تاريخ العقد']):raise Error('الحركة تسبق العقد.')
        if 'تنفيذ' in kinds:
            if s['active']:raise Error('تنفيذ مكرر للعقد.')
            expected=activation_rows(c)
            actual=[(t['النوع'],money(t['المبلغ']),money(t['أثر البنك']),money(t['أثر أموال المستثمر']),t.get('المستثمر') or '') for t in ts]
            exp=[(t[0],t[1],t[2],t[3],c.get('المستثمر') or '') for t in expected]
            if actual!=exp:raise Error('قيود التنفيذ لا تطابق شروط العقد.')
            client=lookup(data,'العملاء',c['العميل']);limit=money(client.get('الحد الائتماني'))
            outstanding=sum(x['p']['cash']*sum(z['amount']-z['paid']-z['waived'] for z in x['parts'])//x['p']['total'] for x in states.values() if x['active'] and x['c']['العميل']==c['العميل'])
            if limit and outstanding+s['p']['cash']>limit:raise Error('التنفيذ يتجاوز الحد الائتماني للعميل.')
            s['active']=True;s['activated']=dt(ts[0]['التاريخ'])
        elif len(ts)!=1:raise Error('مجموعة مالية غير معروفة.')
        for tr in ts:
            kind=tr['النوع'];amount=money(tr['المبلغ']);bd=money(tr['أثر البنك']);hd=money(tr['أثر أموال المستثمر']);inv=tr.get('المستثمر') or ''
            if amount<0:raise Error('مبلغ الحركة لا يقبل السالب.')
            if kind in ['سداد','تسوية']:
                if not s['active'] or s['settled'] or amount<=0:raise Error('السداد يتطلب عقدًا منفذًا وغير مسوى.')
                rem=sum(x['amount']-x['paid']-x['waived'] for x in s['parts']);waive=money(tr.get('خصم التسوية'))
                if waive<0 or amount>rem or (kind=='سداد' and waive) or (kind=='تسوية' and amount+waive!=rem):raise Error('مبلغ السداد أو الخصم لا يطابق المتبقي.')
                left=amount
                for z in s['parts']:
                    take=min(left,z['amount']-z['paid']-z['waived']);z['paid']+=take;left-=take
                    if z['paid']==z['amount'] and z['paid_date'] is None:z['paid_date']=dt(tr['التاريخ'])
                if kind=='تسوية':
                    for z in s['parts']:
                        z['waived']=z['amount']-z['paid']
                        if z['paid_date'] is None:z['paid_date']=dt(tr['التاريخ'])
                    s['settled']=True
                expected_hd=amount if c['النوع']==MANAGED else 0
                if bd!=amount or hd!=expected_hd or inv!=(c.get('المستثمر') or ''):raise Error('ملكية السداد غير مطابقة للعقد.')
            elif kind=='استرداد فائض':
                if not s or not s['active'] or c['النوع']!=EXTERNAL or amount<=0:raise Error('استرداد الفائض يتطلب استثمارًا منفذًا عبر مكتب.')
                s['refund']+=amount
                if s['refund']>s['p']['surplus'] or bd!=amount or hd or inv:raise Error('استرداد الفائض يتجاوز المستحق أو قيده غير مطابق.')
            elif kind=='تحويل للمستثمر':
                if cid or inv not in holdings or bd!=-amount or hd!=-amount or amount<=0:raise Error('تحويل المستثمر غير صالح.')
            elif kind in ['إضافة رأس مال','سحب رأس مال','مصروف مكتب']:
                if cid or inv or hd or bd!=(amount if kind=='إضافة رأس مال' else -amount) or amount<=0:raise Error('قيد السيولة غير مطابق.')
            elif 'تنفيذ' not in kinds:raise Error('نوع حركة غير معروف.')
            if kind!='تسوية' and money(tr.get('خصم التسوية')):raise Error('الخصم يخص التسوية فقط.')
            ref=txt(tr.get('المرجع'))
            if ref:
                if ref in refs:raise Error('مرجع مالي مكرر.')
                refs.add(ref)
            bank+=bd
            if inv:
                if inv not in holdings:raise Error('مستثمر الحركة غير موجود.')
                holdings[inv]+=hd
            elif hd:raise Error('أثر المستثمر دون معرف مستثمر.')
            if bank<0 or any(v<0 for v in holdings.values()) or bank-sum(holdings.values())<0:raise Error('رصيد غير كافٍ أو العملية تستخدم أموال المستثمر لصالح المكتب.')
            ledger.append({**tr,'bank':bank,'held':sum(holdings.values()),'free':bank-sum(holdings.values())})
    summaries=[]
    for cid,s in states.items():
        pp=s['parts'];p=s['p'];paid=sum(z['paid'] for z in pp);waived=sum(z['waived'] for z in pp);rem=p['total']-paid-waived
        overdue=sum(z['amount']-z['paid']-z['waived'] for z in pp if z['due']<now) if s['active'] else 0
        state='مسودة' if not s['active'] else ('مسدد مبكرًا' if s['settled'] else ('مكتمل' if rem==0 else ('متأخر' if overdue else 'منتظم')))
        summaries.append({**s,'id':cid,'paid':paid,'waived':waived,'remaining':rem,'overdue':overdue,'status':state,
          'surplus_due':p['surplus']-s['refund'] if s['active'] and s['c']['النوع']==EXTERNAL else 0,
          'next_due':min((z['due'] for z in pp if z['amount']>z['paid']+z['waived']),default=None) if s['active'] else None})
        for z in pp:
            z['remaining']=z['amount']-z['paid']-z['waived'];z['active']=s['active'];z['client']=s['c']['العميل']
            z['late_days']=max(0,((z['paid_date'] or now)-z['due']).days) if s['active'] else 0
            z['status']='لم ينفذ' if not s['active'] else ('مسدد' if z['remaining']==0 else ('متأخر' if z['due']<now else ('مستحق اليوم' if z['due']==now else ('سداد جزئي' if z['paid'] else 'قادم'))))
    for r in data['المتابعات']:
        lookup(data,'العملاء',r['العميل'])
        if r.get('العقد') and lookup(data,'العقود',r['العقد'])['العميل']!=r['العميل']:raise Error('متابعة مرتبطة بعميل مختلف.')
        if r.get('الحالة') not in ['مفتوحة','مكتملة'] or money(r.get('المبلغ الموعود'))<0:raise Error('متابعة غير صالحة.')
    return {'contracts':summaries,'parts':parts,'ledger':ledger,'bank':bank,'holdings':holdings,'held':sum(holdings.values()),'free':bank-sum(holdings.values())}

def activation_rows(c):
    p=pricing(c);mode=c['النوع']
    # Investor pays the fixed fee in addition to funding. Transfer to office is
    # an internal ownership transfer: bank 0, investor-held cash -fee.
    if mode==MANAGED:
        a=[('تنفيذ',p['fund']+p['fee'],p['fund']+p['fee'],p['fund']+p['fee']),
           ('شراء للمستثمر',p['purchase'],-p['purchase'],-p['purchase'])]
        if p['fee']:a.append(('رسوم ثابتة',p['fee'],0,-p['fee']))
    else:
        out=(p['fund'] if mode==EXTERNAL else p['purchase'])+p['paid_fee']+p['expenses']
        a=[('تنفيذ',out,-out,0)]
    if p['client_fee']:a.append(('رسوم عميل',p['client_fee'],p['client_fee'],0))
    if mode==MANAGED and p['expenses']:a.append(('مصروف صفقة',p['expenses'],-p['expenses'],0))
    return a

def post(data,rows,when,cid='',investor='',ref='',notes='',request=None,discount=0):
    when=dt(when)
    if not when or when>today():raise Error('اختر تاريخًا لا يتجاوز اليوم.')
    group=uid('G');request=request or uid('REQ');stamp=datetime.now(TZ).isoformat(timespec='microseconds')
    for i,(kind,amount,bd,hd) in enumerate(rows):
        data['الحركات'].append({'المعرف':uid('TX'),'المجموعة':group,'التاريخ':when,'الوقت':stamp,'النوع':kind,'العقد':cid,'المستثمر':investor,
          'المبلغ':sar(amount),'أثر البنك':sar(bd),'أثر أموال المستثمر':sar(hd),'خصم التسوية':sar(discount if i==0 else 0),
          'المرجع':txt(ref) if i==0 else '', 'ملاحظات':txt(notes),'الحالة':'نشطة','رقم الطلب':request})
    return group

def perform(original,action,payload,request=None):
    data=deepcopy(original);p=deepcopy(payload);request=request or uid('REQ')
    if any(t['رقم الطلب']==request for t in data['الحركات']):return data
    if action in ['save_client','save_investor','save_office']:
        sheet={'save_client':'العملاء','save_investor':'المستثمرون','save_office':'المكاتب'}[action]
        ident=p.get('المعرف') or uid({'العملاء':'CL','المستثمرون':'INV','المكاتب':'OF'}[sheet]);p['المعرف']=ident
        old=next((r for r in data[sheet] if r['المعرف']==ident),None)
        if p.get('الحالة')=='مؤرشف':
            key={'العملاء':'العميل','المستثمرون':'المستثمر','المكاتب':'المكتب'}[sheet]
            if any(s['c'].get(key)==ident and s['active'] and s['remaining']>0 for s in snapshot(data)['contracts']):raise Error('السجل لديه عقد قائم. استخدم موقوف بدل الأرشفة.')
        if old:old.update(p)
        else:data[sheet].append(p)
    elif action=='save_contract':
        ident=p.get('المعرف') or uid('CT');p['المعرف']=ident
        if any(t.get('العقد')==ident and t['الحالة']=='نشطة' for t in data['الحركات']):raise Error('شروط العقد المنفذ مقفلة. ألغِ التنفيذ والحركات التابعة أولًا.')
        if lookup(data,'العملاء',p['العميل'])['الحالة']!='نشط':raise Error('العميل غير نشط.')
        if p['النوع']==MANAGED and lookup(data,'المستثمرون',p['المستثمر'])['الحالة']!='نشط':raise Error('المستثمر غير نشط.')
        if p['النوع']==EXTERNAL and lookup(data,'المكاتب',p['المكتب'])['الحالة']!='نشط':raise Error('المكتب غير نشط.')
        old=next((r for r in data['العقود'] if r['المعرف']==ident),None)
        if old:old.update(p)
        else:data['العقود'].append(p)
    elif action=='activate':
        c=lookup(data,'العقود',p['contract'])
        if lookup(data,'العملاء',c['العميل'])['الحالة']!='نشط':raise Error('العميل غير نشط.')
        if c['النوع']==MANAGED and lookup(data,'المستثمرون',c['المستثمر'])['الحالة']!='نشط':raise Error('المستثمر غير نشط.')
        if c['النوع']==EXTERNAL and lookup(data,'المكاتب',c['المكتب'])['الحالة']!='نشط':raise Error('المكتب غير نشط.')
        post(data,activation_rows(c),p['date'],c['المعرف'],c.get('المستثمر') or '',p.get('ref'),p.get('notes'),request)
    elif action in ['receipt','settle','refund']:
        c=lookup(data,'العقود',p['contract']);a=money(p['amount']);hd=a if c['النوع']==MANAGED else 0
        kind={'receipt':'سداد','settle':'تسوية','refund':'استرداد فائض'}[action]
        post(data,[(kind,a,a,hd if action!='refund' else 0)],p['date'],c['المعرف'],c.get('المستثمر') or '',p.get('ref'),p.get('notes'),request,money(p.get('discount')))
    elif action=='cash':
        a=money(p['amount']);kind=p['kind']
        post(data,[(kind,a,a if kind=='إضافة رأس مال' else -a,-a if kind=='تحويل للمستثمر' else 0)],p['date'],investor=p.get('investor') or '',ref=p.get('ref'),notes=p.get('notes'),request=request)
    elif action in ['void','restore']:
        matches=[t for t in data['الحركات'] if t['المجموعة']==p['group']]
        if not matches:raise Error('مجموعة الحركة غير موجودة.')
        for t in matches:t['الحالة']='ملغاة' if action=='void' else 'نشطة'
    elif action=='save_followup':
        p['المعرف']=p.get('المعرف') or uid('FU');old=next((r for r in data['المتابعات'] if r['المعرف']==p['المعرف']),None)
        if old:old.update(p)
        else:data['المتابعات'].append(p)
    elif action=='settings':
        if not 1<=int(p['default_months'])<=120 or not 0<=float(p['default_profit_rate'])<=10:raise Error('الإعدادات خارج المدى.')
        for r in data['النظام']:
            if r['المفتاح'] in p:r['القيمة']=p[r['المفتاح']]
    else:raise Error('عملية غير معروفة.')
    snapshot(data)
    return data

def read_excel(raw):
    try:wb=load_workbook(BytesIO(raw),data_only=False)
    except Exception:raise Error('الملف ليس ملف Excel صالحًا.')
    data={}
    for name,headers in HEADERS.items():
        if name not in wb.sheetnames:raise Error('هذا ليس ملف النظام الجديد V2. استخدم Installment_System_V2.xlsx.')
        ws=wb[name];actual=[x.value for x in ws[1]]
        if actual[:len(headers)]!=headers:raise Error('أعمدة ورقة '+name+' غير متوافقة مع V2.')
        if name=='الأقساط':
            data[name]=[]
            continue
        count=INPUT_COLUMNS.get(name,len(headers));data[name]=[]
        for row in ws.iter_rows(min_row=2):
            if row[0].value in [None,'']:continue
            record={headers[i]:row[i].value for i in range(count)}
            if any(row[i].data_type=='f' for i in range(count)):raise Error('حقول البيانات الأصلية لا تقبل معادلات. استخدم التطبيق لتعديلها.')
            data[name].append(record)
    if setting(data,'format')!=FORMAT:raise Error('بنية الملف ليست V2.')
    snapshot(data);return data

def put(ws,row,col,value):
    cell=ws.cell(row,col);cell.value=value
    if isinstance(value,str):cell.data_type='s'

def write_excel(data,template=None):
    """Application serializer. Rebuild calculated sheets from the raw records."""
    snap=snapshot(data)
    wb=load_workbook(BytesIO(template)) if template else load_workbook(ROOT/'Installment_System_V2.xlsx')
    for name,headers in HEADERS.items():
        ws=wb[name]
        for key in list(ws.tables):del ws.tables[key]
        ws.delete_rows(2,ws.max_row)
        for i,record in enumerate(data[name],2):
            for j,key in enumerate(headers,1):
                if key in record:put(ws,i,j,record[key])
    schedules=[]
    for z in snap['parts']:
        schedules.append([z['id'],z['contract'],z['sequence'],z['due'],sar(z['amount']),sar(z['paid']),sar(z['waived']),sar(z['remaining']),z['status'],z['late_days'],z['paid_date']])
    ws=wb['الأقساط'];ws.delete_rows(2,ws.max_row)
    for i,record in enumerate(schedules,2):
        for j,v in enumerate(record,1):put(ws,i,j,v)
        ws.cell(i,8,f'=MAX(0,E{i}-F{i}-G{i})')
    active={s['id']:s for s in snap['contracts']};sn=max(2,len(schedules)+1)
    ws=wb['العقود']
    for i,c in enumerate(data['العقود'],2):
        s=active[c['المعرف']];p=s['p'];managed=c['النوع']==MANAGED
        ws.cell(i,18,f'=J{i}-I{i}')
        ws.cell(i,19,f'=I{i}+N{i}' if managed else f'=I{i}+M{i}+P{i}')
        ws.cell(i,20,f'=K{i}-S{i}')
        ws.cell(i,21,f'=N{i}+O{i}-P{i}' if managed else f'=T{i}+O{i}')
        ws.cell(i,22,f'=IF(S{i}=0,0,T{i}/S{i})')
        ws.cell(i,23,f'=SUMIFS(\'الأقساط\'!$F$2:$F${sn},\'الأقساط\'!$B$2:$B${sn},A{i})')
        ws.cell(i,24,f'=SUMIFS(\'الأقساط\'!$G$2:$G${sn},\'الأقساط\'!$B$2:$B${sn},A{i})')
        ws.cell(i,25,f'=K{i}-W{i}-X{i}')
        put(ws,i,26,s['status']);put(ws,i,27,sar(s['surplus_due']));put(ws,i,28,s['activated'])
        ws.cell(i,22).number_format='0.0%'
    bytx={t['المعرف']:t for t in snap['ledger']}
    ws=wb['الحركات']
    for i,tr in enumerate(data['الحركات'],2):
        t=bytx.get(tr['المعرف'])
        if t:
            for j,k in [(16,'bank'),(17,'held'),(18,'free')]:put(ws,i,j,sar(t[k]))
    dash=wb['الملخص'];dash.delete_rows(1,dash.max_row)
    dash.append(['إدارة التقسيط 2.0','القيمة بالريال / العدد']);dash.append(['رصيد البنك',sar(snap['bank'])]);dash.append(['أموال المستثمرين المحتجزة',sar(snap['held'])]);dash.append(['السيولة الحرة للمكتب','=B2-B3'])
    funded=[s for s in snap['contracts'] if s['active']]
    for row in [('متبقي عقودك',sar(sum(s['remaining'] for s in funded if s['c']['النوع']!=MANAGED))),('متبقي عقود المستثمرين',sar(sum(s['remaining'] for s in funded if s['c']['النوع']==MANAGED))),('الربح المتوقع للمكتب',sar(sum(s['p']['office_profit'] for s in funded))),('تحصيل الأقساط',sar(sum(s['paid'] for s in funded))),('الخصومات',sar(sum(s['waived'] for s in funded))),('فائض مستحق لدى المكاتب',sar(sum(s['surplus_due'] for s in funded))),('عدد العملاء',len(data['العملاء'])),('العقود المنفذة',len(funded)),('آخر حفظ',datetime.now(TZ).isoformat(timespec='seconds'))]:dash.append(row)
    for ws in wb:
        ws.sheet_view.rightToLeft=True;ws.sheet_view.showGridLines=False;ws.freeze_panes='C2' if ws.title!='الملخص' else None
        ws.row_dimensions[1].height=38
        for cell in ws[1]:cell.font=Font(name='Arial',size=11,bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='1B3744');cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        for row in ws.iter_rows(min_row=2):
            ws.row_dimensions[row[0].row].height=27
            for cell in row:
                cell.font=Font(name='Arial',size=11,color='156D5A' if cell.data_type=='f' else '18323D')
                cell.alignment=Alignment(horizontal='right',vertical='center')
                if isinstance(cell.value,(int,float)) and cell.number_format=='General':cell.number_format='#,##0.00;(#,##0.00);"—"'
                if isinstance(cell.value,(date,datetime)):cell.number_format='yyyy-mm-dd'
        for j in range(1,ws.max_column+1):ws.column_dimensions[get_column_letter(j)].width=25 if ws.title!='الملخص' else (45 if j==1 else 35)
        if ws.title in HEADERS and ws.max_row>1:
            table=Table(displayName='Data'+str(wb.sheetnames.index(ws.title)),ref=f'A1:{get_column_letter(ws.max_column)}{ws.max_row}')
            table.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True);ws.add_table(table)
    wb.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True)
    b=BytesIO();wb.save(b);return b.getvalue()

def digest(raw):return hashlib.sha256(raw).hexdigest()
def atomic(path,raw):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    fd,tmp=tempfile.mkstemp(prefix='.save-',suffix='.xlsx',dir=path.parent)
    try:
        with os.fdopen(fd,'wb') as f:f.write(raw);f.flush();os.fsync(f.fileno())
        os.replace(tmp,path)
    finally:
        if os.path.exists(tmp):os.unlink(tmp)

class LocalStore:
    def __init__(self,path):self.path=Path(path)
    def read(self):
        with FileLock(str(self.path)+'.lock'):
            raw=self.path.read_bytes();read_excel(raw);return raw,digest(raw)
    def commit(self,tag,action,payload,request):
        with FileLock(str(self.path)+'.lock'):
            raw=self.path.read_bytes()
            if digest(raw)!=tag:raise Conflict('تغير المصدر. اضغط تحديث ثم راجع العملية.')
            data=perform(read_excel(raw),action,payload,request);new=write_excel(data,raw)
            back=self.path.parent/'backups'/(datetime.now(TZ).strftime('%Y%m%d_%H%M%S_%f')+'.xlsx');atomic(back,raw)
            atomic(self.path,new);return new,digest(new)

class CloudStore:
    def __init__(self,token,folder,name):
        self.session=requests.Session();self.session.headers['Authorization']='Bearer '+token
        self.graph='https://graph.microsoft.com/v1.0';self.folder=folder.strip('/');self.name=name
        if name!='Installment_System_V2.xlsx':raise Error('اسم ملف النسخة الجديدة يجب أن يكون Installment_System_V2.xlsx.')
        self.path=quote(self.folder+'/'+name,safe='/');self.item=''
    def check(self,r):
        if r.status_code in [409,412]:raise Conflict('تغير المصدر أثناء الحفظ. حدّث وتحقق من سجل الحركات قبل الإعادة.')
        if r.status_code==404:raise Error('ارفع Installment_System_V2.xlsx داخل مجلد '+self.folder+' في OneDrive أولًا.')
        if r.status_code in [401,403]:raise Error('انتهت جلسة Microsoft أو لا تملك صلاحية الملف. أعد تسجيل الدخول.')
        if r.status_code==423:raise Error('الملف مقفل. أغلقه في Excel وحاول مجددًا.')
        if not 200<=r.status_code<300:raise Error('تعذر الاتصال بـOneDrive. HTTP '+str(r.status_code))
    def meta(self):
        r=self.session.get(self.graph+'/me/drive/root:/'+self.path,params={'$select':'id,eTag'},timeout=35);self.check(r)
        m=r.json();self.item=m['id']
        if not m.get('eTag'):raise Conflict('تعذر تحديد نسخة الملف.')
        return m
    def read(self):
        a=self.meta();r=self.session.get(self.graph+'/me/drive/items/'+quote(self.item,safe='')+'/content',timeout=90);self.check(r)
        b=self.meta()
        if a['eTag']!=b['eTag']:raise Conflict('تغير الملف أثناء القراءة. حدّث مرة أخرى.')
        read_excel(r.content);return r.content,b['eTag']
    def commit(self,tag,action,payload,request):
        raw,current=self.read();data=read_excel(raw)
        if any(t['رقم الطلب']==request for t in data['الحركات']):return raw,current
        if current!=tag:raise Conflict('تغير Excel بعد فتح الشاشة. اضغط تحديث وراجع البيانات ثم احفظ.')
        new=write_excel(perform(data,action,payload,request),raw)
        # Keep recoverable bytes before the conditional remote write.
        backup=Path(tempfile.gettempdir())/'installments_v2_recovery'/digest(self.item.encode())/(datetime.now(TZ).strftime('%Y%m%d_%H%M%S_%f')+'.xlsx');atomic(backup,raw)
        r=self.session.post(self.graph+'/me/drive/items/'+quote(self.item,safe='')+'/createUploadSession',headers={'If-Match':current},json={'item':{'@microsoft.graph.conflictBehavior':'replace'}},timeout=45);self.check(r)
        url=r.json()['uploadUrl']
        if self.meta()['eTag']!=current:raise Conflict('تغير الملف أثناء تجهيز الحفظ. حدّث المصدر.')
        try:r=requests.put(url,data=new,headers={'Content-Length':str(len(new)),'Content-Range':f'bytes 0-{len(new)-1}/{len(new)}'},timeout=120)
        except requests.RequestException:raise Conflict('لم يصل تأكيد الحفظ. حدّث المصدر وتحقق من رقم الطلب قبل إعادة العملية.')
        self.check(r)
        if r.status_code not in [200,201] or not r.json().get('eTag'):raise Conflict('لم يصل تأكيد نهائي. حدّث المصدر وتحقق من رقم الطلب.')
        return new,r.json()['eTag']

def config(section,key,default=''):
    try:return st.secrets[section].get(key,default)
    except Exception:return default
def format_money(h):return f'{sar(h):,.2f}'
def show_table(rows):
    if not rows:st.info('لا توجد سجلات حتى الآن.');return
    clean=[{k:(v.isoformat() if isinstance(v,(date,datetime)) else v) for k,v in r.items()} for r in rows]
    st.dataframe(pd.DataFrame(clean),hide_index=True,width='stretch')
def metrics(items):
    cols=st.columns(min(4,len(items)))
    for i,(label,value) in enumerate(items):cols[i%len(cols)].metric(label,value)
def names(data,sheet):return {r['المعرف']:r['الاسم'] for r in data[sheet]}
def choose(data,sheet,label,key,allow_empty=False,active=False,default=None):
    rows=[r for r in data[sheet] if not active or r['الحالة']=='نشط'];mapping={r['المعرف']:r['الاسم'] for r in rows}
    options=([''] if allow_empty else [])+list(mapping)
    if not options:st.info('أضف سجلًا نشطًا في '+sheet+' أولًا.');return None
    return st.selectbox(label,options,index=options.index(default) if default in options else 0,format_func=lambda x:mapping.get(x,'بدون'),key=key)
def save(action,payload):
    req=st.session_state.setdefault('request',uid('REQ'))
    try:
        raw,tag=st.session_state['store'].commit(st.session_state['tag'],action,payload,req)
        st.session_state['raw']=raw;st.session_state['tag']=tag;st.session_state.pop('request',None)
        st.session_state['notice']='تم الحفظ في Excel بنجاح.';st.rerun()
    except Exception as exc:
        st.error(str(exc));st.caption('رقم الطلب: '+req+'. عند انقطاع الاتصال، حدّث المصدر قبل إعادة الحفظ.')

def contract_rows(data,ss):
    cl=names(data,'العملاء');iv=names(data,'المستثمرون');office=names(data,'المكاتب')
    return [{'العقد':s['id'],'العميل':cl.get(s['c']['العميل'],''),'النوع':s['c']['النوع'],'المستثمر':iv.get(s['c'].get('المستثمر'),''),'المكتب':office.get(s['c'].get('المكتب'),''),
      'تكلفة الشراء':sar(s['p']['purchase']),'سعر الكاش':sar(s['p']['cash']),'إجمالي التقسيط':sar(s['p']['total']),'ربح المكتب المتوقع':sar(s['p']['office_profit']),
      'المحصل':sar(s['paid']),'الخصم':sar(s['waived']),'المتبقي':sar(s['remaining']),'المتأخر':sar(s['overdue']),'الاستحقاق القادم':s['next_due'],'الحالة':s['status']} for s in ss]
def part_rows(pp):
    return [{'القسط':z['sequence'],'العقد':z['contract'],'الاستحقاق':z['due'],'المبلغ':sar(z['amount']),'المحصل':sar(z['paid']),'الخصم':sar(z['waived']),'المتبقي':sar(z['remaining']),'الحالة':z['status'],'أيام التأخير':z['late_days']} for z in pp]

def home(data,snap):
    funded=[s for s in snap['contracts'] if s['active']]
    metrics([('السيولة الحرة',format_money(snap['free'])),('رصيد البنك',format_money(snap['bank'])),('أمانات المستثمرين',format_money(snap['held'])),('ربح المكتب المتوقع',format_money(sum(s['p']['office_profit'] for s in funded))),
      ('متبقي عقودك',format_money(sum(s['remaining'] for s in funded if s['c']['النوع']!=MANAGED))),('متبقي عقود المستثمرين',format_money(sum(s['remaining'] for s in funded if s['c']['النوع']==MANAGED))),('المتأخر',format_money(sum(s['overdue'] for s in funded))),('العملاء',str(len(data['العملاء'])))])
    st.caption('كل القيم بالريال. الربح المتوقع يفترض اكتمال السداد قبل الخصومات والتعثر. أمانات المستثمرين مستبعدة من السيولة الحرة.')
    st.subheader('العقود');show_table(contract_rows(data,snap['contracts']))
    a,b=st.columns(2)
    with a:
        st.subheader('المستحق والمتأخر');show_table(part_rows([z for z in snap['parts'] if z['active'] and z['remaining']>0 and (z['due']-today()).days<=7]))
    with b:
        st.subheader('متابعات مطلوبة');cl=names(data,'العملاء')
        show_table([{'العميل':cl.get(r['العميل'],''),'الموعد':r['المتابعة القادمة'],'النتيجة':r.get('النتيجة')} for r in data['المتابعات'] if r['الحالة']=='مفتوحة' and r.get('المتابعة القادمة') and dt(r['المتابعة القادمة'])<=today()])

def people(data,snap,sheet):
    cl=sheet=='العملاء';inv=sheet=='المستثمرون'
    action='save_client' if cl else ('save_investor' if inv else 'save_office')
    mapping=names(data,sheet);ident=st.selectbox('إضافة أو تعديل',['']+list(mapping),format_func=lambda x:mapping.get(x,'إضافة سجل جديد'),key='person-'+sheet)
    old=lookup(data,sheet,ident) if ident else {};p={}
    with st.form('person-form-'+sheet+'-'+ident):
        a,b=st.columns(2)
        with a:
            p['الاسم']=st.text_input('الاسم',old.get('الاسم') or '')
            p['الجوال']=st.text_input('الجوال',old.get('الجوال') or '')
            if cl or inv:p['الهوية']=st.text_input('الهوية',old.get('الهوية') or '')
        with b:
            states=['نشط','موقوف','مؤرشف'];p['الحالة']=st.selectbox('الحالة',states,index=states.index(old.get('الحالة') or 'نشط'))
            if cl:
                grades=['غير مصنف','A','B','C','D'];p['المستوى']=st.selectbox('المستوى',grades,index=grades.index(old.get('المستوى') or 'غير مصنف'))
                p['الحد الائتماني']=st.number_input('الحد الائتماني (صفر = غير محدد)',min_value=0.0,value=float(old.get('الحد الائتماني') or 0))
        if cl:
            a,b=st.columns(2)
            with a:
                for k in ['العنوان','جهة العمل','الكفيل']:p[k]=st.text_input(k,old.get(k) or '')
            with b:
                p['الدخل']=st.number_input('الدخل الشهري',min_value=0.0,value=float(old.get('الدخل') or 0));p['جوال الكفيل']=st.text_input('جوال الكفيل',old.get('جوال الكفيل') or '')
        p['ملاحظات']=st.text_area('ملاحظات',old.get('ملاحظات') or '')
        if st.form_submit_button('حفظ في Excel',type='primary'):save(action,{**p,**({'المعرف':ident} if ident else {})})
    if inv:
        show_table([{'الاسم':r['الاسم'],'الجوال':r['الجوال'],'أموال محتجزة':sar(snap['holdings'].get(r['المعرف'],0)),'متبقي العقود':sar(sum(s['remaining'] for s in snap['contracts'] if s['active'] and s['c'].get('المستثمر')==r['المعرف'])),'الحالة':r['الحالة']} for r in data[sheet]])
    else:show_table(data[sheet])
    if ident:
        key='العميل' if cl else ('المستثمر' if inv else 'المكتب');show_table(contract_rows(data,[s for s in snap['contracts'] if s['c'].get(key)==ident]))

def contract_form(data,old=None,key='new'):
    old=old or {};p={};st.caption('الحسابات تتحدث عند تعديل الحقول. حفظ العقد ينشئ مسودة؛ التنفيذ يسجل النقد والرسوم الفعلية.')
    modes=MODES;p['النوع']=st.selectbox('نوع الصفقة',modes,index=modes.index(old.get('النوع',SELF)),key=key+'-mode')
    p['العميل']=choose(data,'العملاء','العميل',key+'-client',active=not bool(old),default=old.get('العميل'))
    p['المستثمر']='';p['المكتب']=''
    if p['النوع']==MANAGED:
        p['المستثمر']=choose(data,'المستثمرون','المستثمر',key+'-investor',active=not bool(old),default=old.get('المستثمر'))
        st.info('كامل تحصيلات الأقساط للمستثمر. رسومك الثابتة يدفعها المستثمر إضافة إلى مبلغ التمويل عند التنفيذ. فرق الشراء يبقى أمانة للمستثمر.')
    if p['النوع']==EXTERNAL:p['المكتب']=choose(data,'المكاتب','المكتب المنفذ',key+'-office',active=not bool(old),default=old.get('المكتب'))
    a,b=st.columns(2)
    with a:
        p['تكلفة الشراء']=st.number_input('تكلفة الشراء الفعلية',min_value=0.0,value=float(old.get('تكلفة الشراء') or 4970),step=10.0,key=key+'-purchase')
        p['سعر الكاش']=st.number_input('سعر الكاش / أصل التمويل للعميل',min_value=0.0,value=float(old.get('سعر الكاش') or 5000),step=100.0,key=key+'-cash')
        suggested=float(Decimal(str(p['سعر الكاش']))*(1+Decimal(str(setting(data,'default_profit_rate')))))
        p['إجمالي التقسيط']=st.number_input('إجمالي الأقساط شامل الربح',min_value=0.0,value=float(old.get('إجمالي التقسيط') or suggested),step=100.0,key=key+'-total')
        st.caption('القيمة المقترحة وفق الإعدادات: '+f'{suggested:,.2f}'+' ريال. عدّل إجمالي التقسيط حسب الاتفاق.')
        p['التمويل المحول']=p['تكلفة الشراء'] if p['النوع']==SELF else st.number_input('التمويل المحول لشراء السلعة (دون الرسوم)',min_value=0.0,value=float(old.get('التمويل المحول') or 5000),step=100.0,key=key+'-fund')
    with b:
        p['رسوم مكتب مدفوعة']=st.number_input('رسوم مدفوعة للمكتب المنفذ',min_value=0.0,value=float(old.get('رسوم مكتب مدفوعة') or 0),key=key+'-paidfee') if p['النوع']==EXTERNAL else 0
        p['رسوم مكتب ثابتة']=st.number_input('رسومك الثابتة على المستثمر',min_value=0.0,value=float(old.get('رسوم مكتب ثابتة') or 0),key=key+'-fee') if p['النوع']==MANAGED else 0
        p['رسوم العميل']=st.number_input('رسوم إضافية يدفعها العميل خارج الأقساط',min_value=0.0,value=float(old.get('رسوم العميل') or 0),key=key+'-clientfee')
        p['مصاريف أخرى']=st.number_input('مصاريف أخرى تتحملها أنت',min_value=0.0,value=float(old.get('مصاريف أخرى') or 0),key=key+'-expenses')
        p['الأشهر']=st.number_input('عدد الأقساط الشهرية',min_value=1,max_value=120,value=int(old.get('الأشهر') or setting(data,'default_months')),key=key+'-months')
        p['تاريخ العقد']=st.date_input('تاريخ العقد',value=dt(old.get('تاريخ العقد')) or today(),max_value=today(),key=key+'-signed')
        p['أول استحقاق']=st.date_input('أول استحقاق',value=dt(old.get('أول استحقاق')) or month(today(),1),key=key+'-first')
    p['ملاحظات']=st.text_area('ملاحظات العقد',old.get('ملاحظات') or '',key=key+'-notes')
    q=pricing(p);metrics([('فرق الشراء',format_money(q['saving'])),('ربح المكتب المتوقع',format_money(q['office_profit'])),('ربح المستثمر المتوقع',format_money(q['investor_profit'])),('القسط التقريبي',format_money(q['total']//int(p['الأشهر'])))])
    if p['النوع']==EXTERNAL and q['surplus']>0:st.caption('فائض لدى المكتب يجب استرداده وتسجيله: '+format_money(q['surplus'])+' ريال. الربح المتوقع يفترض استرداده.')
    if p['النوع']==MANAGED:st.caption('المطلوب استلامه من المستثمر عند التنفيذ: '+format_money(q['fund']+q['fee'])+' ريال. أمانته المتبقية بعد الشراء: '+format_money(q['surplus'])+' ريال.')
    if st.button('حفظ مسودة العقد',type='primary',key=key+'-save'):
        if not p['العميل'] or (p['النوع']==MANAGED and not p['المستثمر']) or (p['النوع']==EXTERNAL and not p['المكتب']):st.error('أضف واختر العميل والجهة الممولة أولًا.')
        else:save('save_contract',{**p,**({'المعرف':old['المعرف']} if old else {})})

def contracts_page(data,snap):
    new,existing=st.tabs(['عقد جديد','العقود والتحصيل'])
    with new:contract_form(data)
    with existing:
        cl=names(data,'العملاء');ss={s['id']:s for s in snap['contracts']}
        show_table(contract_rows(data,snap['contracts']))
        if not ss:return
        cid=st.selectbox('اختر العقد',list(ss),format_func=lambda x:cl.get(ss[x]['c']['العميل'],'')+' · '+x)
        s=ss[cid];c=s['c'];p=s['p']
        metrics([('إجمالي التقسيط',format_money(p['total'])),('المحصل',format_money(s['paid'])),('المتبقي',format_money(s['remaining'])),('الحالة',s['status'])])
        if not s['active']:
            st.write('قيود التنفيذ التي ستسجل:')
            show_table([{'البيان':r[0],'المبلغ':sar(r[1]),'أثر البنك':sar(r[2]),'أثر أمانات المستثمر':sar(r[3])} for r in activation_rows(c)])
            with st.form('activate-'+cid):
                d=st.date_input('تاريخ التنفيذ',value=today(),max_value=today());ref=st.text_input('مرجع التحويل / العملية');notes=st.text_input('ملاحظات التنفيذ')
                ok=st.checkbox('تمت جميع التحويلات والشراء وتحصيل الرسوم الموضحة أعلاه فعليًا')
                if st.form_submit_button('تنفيذ الصفقة',type='primary'):
                    if not ok:st.error('أكد إتمام العمليات الفعلية أولًا.')
                    else:save('activate',{'contract':cid,'date':d,'ref':ref,'notes':notes})
            with st.expander('تعديل المسودة'):contract_form(data,c,key='edit-'+cid)
        elif s['remaining']>0:
            if c['النوع']==MANAGED:st.info('كامل مبلغ السداد يضاف إلى أموال المستثمر المحتجزة.')
            with st.form('receipt-'+cid):
                action=st.selectbox('العملية',['سداد قسط أو مبلغ جزئي','تسوية مبكرة'])
                amount=st.number_input('المبلغ المستلم فعليًا',min_value=0.0,max_value=sar(s['remaining']),value=min(sar(s['remaining']),sar(p['total']//int(c['الأشهر']))))
                discount=st.number_input('خصم التسوية المتفق عليه (للتسوية المبكرة فقط)',min_value=0.0,value=0.0)
                d=st.date_input('تاريخ السداد',max_value=today());ref=st.text_input('مرجع السداد');notes=st.text_input('المحصل / ملاحظات')
                if st.form_submit_button('تسجيل السداد',type='primary'):save('settle' if action=='تسوية مبكرة' else 'receipt',{'contract':cid,'amount':amount,'discount':discount,'date':d,'ref':ref,'notes':notes})
        if s['surplus_due']>0:
            with st.form('refund-'+cid):
                st.write('فائض مستحق لك لدى المكتب: '+format_money(s['surplus_due']))
                amount=st.number_input('المبلغ المسترد',min_value=0.0,max_value=sar(s['surplus_due']),value=sar(s['surplus_due']));d=st.date_input('تاريخ استرداد الفائض',max_value=today());ref=st.text_input('مرجع الاسترداد')
                if st.form_submit_button('تسجيل استرداد الفائض'):save('refund',{'contract':cid,'amount':amount,'date':d,'ref':ref})
        show_table(part_rows(s['parts']))

def cash_page(data,snap):
    metrics([('البنك',format_money(snap['bank'])),('أمانات المستثمرين',format_money(snap['held'])),('المتاح للمكتب',format_money(snap['free']))])
    kind=st.selectbox('نوع الحركة',['إضافة رأس مال','سحب رأس مال','مصروف مكتب','تحويل للمستثمر'])
    inv=choose(data,'المستثمرون','المستثمر','transfer-investor') if kind=='تحويل للمستثمر' else ''
    if inv:st.caption('المتاح تحويله للمستثمر: '+format_money(snap['holdings'].get(inv,0)))
    with st.form('cash'):
        amount=st.number_input('المبلغ',min_value=0.0,value=0.0);d=st.date_input('تاريخ الحركة',max_value=today());ref=st.text_input('المرجع');notes=st.text_input('الوصف')
        if st.form_submit_button('حفظ الحركة',type='primary'):save('cash',{'kind':kind,'amount':amount,'date':d,'investor':inv,'ref':ref,'notes':notes})
    st.subheader('سجل الحركات');show_table(list(reversed(data['الحركات'])))
    with st.expander('تصحيح حركة بإلغاء المجموعة أو استعادتها'):
        groups={t['المجموعة']:t for t in data['الحركات']}
        if groups:
            group=st.selectbox('المجموعة',list(groups),format_func=lambda x:x+' · '+str(groups[x]['التاريخ'])+' · '+groups[x]['الحالة'])
            ok=st.checkbox('راجعت المجموعة كاملة وأريد تغيير حالتها')
            if st.button('إلغاء المجموعة' if groups[group]['الحالة']=='نشطة' else 'استعادة المجموعة',disabled=not ok):save('void' if groups[group]['الحالة']=='نشطة' else 'restore',{'group':group})
        st.caption('إلغاء التنفيذ يلغي قيوده معًا. إذا ترتب عليه سداد أو تحويل لاحق غير صالح فسيرفض النظام الإلغاء حتى تصحح الحركات التابعة.')

def followups(data,snap):
    rows={r['المعرف']:r for r in data['المتابعات']};ident=st.selectbox('إضافة أو تعديل متابعة',['']+list(rows));old=rows.get(ident,{})
    client=choose(data,'العملاء','العميل','follow-client')
    if old:client=old['العميل'];st.caption('عميل المتابعة: '+names(data,'العملاء').get(client,''))
    if not client:return
    options=['']+[c['المعرف'] for c in data['العقود'] if c['العميل']==client]
    with st.form('follow-'+ident):
        p={'العميل':client}
        p['العقد']=st.selectbox('العقد (اختياري)',options,index=options.index(old.get('العقد') or '') if (old.get('العقد') or '') in options else 0)
        p['تاريخ التواصل']=st.date_input('تاريخ التواصل',value=dt(old.get('تاريخ التواصل')) or today())
        p['الوسيلة']=st.selectbox('وسيلة التواصل',['اتصال','واتساب','زيارة','رسالة','أخرى']);p['النتيجة']=st.text_input('نتيجة التواصل',old.get('النتيجة') or '')
        promise=st.checkbox('يوجد وعد سداد',value=bool(old.get('وعد السداد')))
        pdue=st.date_input('موعد الوعد',value=dt(old.get('وعد السداد')) or today());p['وعد السداد']=pdue if promise else None
        p['المبلغ الموعود']=st.number_input('المبلغ الموعود',min_value=0.0,value=float(old.get('المبلغ الموعود') or 0))
        p['المتابعة القادمة']=st.date_input('موعد المتابعة القادمة',value=dt(old.get('المتابعة القادمة')) or today());p['الحالة']=st.selectbox('الحالة',['مفتوحة','مكتملة'],index=1 if old.get('الحالة')=='مكتملة' else 0)
        p['ملاحظات']=st.text_area('الملاحظات',old.get('ملاحظات') or '')
        if st.form_submit_button('حفظ المتابعة',type='primary'):save('save_followup',{**p,**({'المعرف':ident} if ident else {})})
    show_table(data['المتابعات'])

def reports(data,snap):
    start=st.date_input('من تاريخ',value=today().replace(day=1));end=st.date_input('إلى تاريخ',value=today())
    if start>end:st.error('بداية الفترة بعد نهايتها.');return
    txs=[t for t in snap['ledger'] if start<=dt(t['التاريخ'])<=end]
    bycontract={s['id']:s for s in snap['contracts']}
    managed=sum(money(t['المبلغ']) for t in txs if t['النوع'] in ['سداد','تسوية'] and bycontract[t['العقد']]['c']['النوع']==MANAGED)
    own=sum(money(t['المبلغ']) for t in txs if t['النوع'] in ['سداد','تسوية'] and bycontract[t['العقد']]['c']['النوع']!=MANAGED)
    fees=sum(money(t['المبلغ']) for t in txs if t['النوع'] in ['رسوم ثابتة','رسوم عميل'])
    metrics([('تحصيل أقساطك خلال الفترة',format_money(own)),('تحصيل للمستثمرين',format_money(managed)),('رسوم مكتسبة',format_money(fees)),('صافي أثر البنك',format_money(sum(money(t['أثر البنك']) for t in txs)))])
    st.caption('تحصيل أقساطك يشمل رأس المال والربح. الرسوم المكتسبة منفصلة. صافي أثر البنك يشمل التمويل ورأس المال والمصروفات والتحويلات.')
    show_table([{k:t.get(k) for k in HEADERS['الحركات'][:15]} for t in txs])
    st.subheader('قائمة التعثر الحالية');show_table(contract_rows(data,[s for s in snap['contracts'] if s['overdue']>0]))
    st.subheader('حقوق المستثمرين الحالية')
    show_table([{'المستثمر':r['الاسم'],'أمانات نقدية':sar(snap['holdings'][r['المعرف']]),'أقساط غير محصلة':sar(sum(s['remaining'] for s in snap['contracts'] if s['active'] and s['c'].get('المستثمر')==r['المعرف']))} for r in data['المستثمرون']])

def settings_page(data,snap):
    st.download_button('تنزيل نسخة Excel الحالية',st.session_state['raw'],file_name='Installment_System_V2.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
    with st.form('settings'):
        months=st.number_input('عدد الأقساط الافتراضي',min_value=1,max_value=120,value=int(setting(data,'default_months')))
        rate=st.number_input('الربح الإجمالي المقترح %',min_value=0.0,max_value=1000.0,value=float(setting(data,'default_profit_rate'))*100)
        if st.form_submit_button('حفظ الإعدادات'):save('settings',{'default_months':months,'default_profit_rate':rate/100})
    st.write('التصنيف A–D يدوي. فرق الشراء = سعر الكاش − التكلفة الفعلية. الربح المتوقع يحسب مرة واحدة من إجمالي التقسيط ناقص التكلفة والرسوم والمصاريف. الربح لا يضاف إلى سعر التقسيط مرة أخرى.')
    st.write('في العقد المدار: المستثمر يملك كامل الأقساط. رسومه الثابتة تدفع إضافة إلى التمويل وتتحول لملكية المكتب عند التنفيذ. فائض التمويل بعد الشراء يبقى أمانة للمستثمر حتى تحويله.')
    st.write('تعديل Excel خارج التطبيق متاح في حقول البيانات الأصلية وفق العناوين. الأقساط والملخص والحقول المحسوبة يعاد توليدها عند الحفظ. لتغيير الشروط المالية لعقد منفذ، صحح الحركات التابعة ثم ألغِ مجموعة التنفيذ أولًا.')
    st.caption('التطبيق يبدأ برصيد صفر. أضف رأس المال الحقيقي من صفحة الخزينة. ملف V1 مستقل ولا يُستبدل بهذا الملف.')

def main():
    st.set_page_config(page_title='إدارة التقسيط | إبراهيم',page_icon='📊',layout='wide')
    st.markdown('''<style>.stApp{direction:rtl;font-family:Tahoma,Arial,sans-serif}h1,h2,h3,p,label{font-family:Tahoma,Arial,sans-serif!important}[data-testid="stSidebar"]{direction:rtl;border-left:1px solid #29424e}[data-testid="stMetric"]{background:#1a282f;border:1px solid #2a4451;border-radius:14px;padding:16px;margin:5px 0}[data-testid="stMetricValue"]{color:#57d2c1;font-size:1.7rem;direction:ltr;text-align:right}.block-container{max-width:1550px;padding-top:2rem}div.stButton>button{border-radius:9px}div[data-testid="stForm"]{border:1px solid #29424e;border-radius:14px}h1{font-size:2rem!important}</style>''',unsafe_allow_html=True)
    st.sidebar.title('إدارة التقسيط');st.sidebar.caption('إبراهيم · الإصدار '+VERSION)
    mode=os.getenv('INSTALLMENT_STORAGE_MODE',str(config('storage','mode','local')))
    try:
        if mode=='onedrive':
            import microsoft_auth as auth
            if not auth.cloud_mode_enabled():st.info('أكمل إعدادات Microsoft في Secrets حسب دليل التثبيت.');st.stop()
            token=auth.acquire_microsoft_token()
            if not token:st.stop()
            cfg=auth.microsoft_config();store=CloudStore(token,cfg['onedrive_folder'],cfg['workbook_name']);identity='cloud:'+store.path
            st.sidebar.success('متصل بملف OneDrive V2')
            if st.sidebar.button('تسجيل الخروج'):
                st.session_state.clear();st.rerun()
        elif mode=='local':
            store=LocalStore(Path(os.getenv('INSTALLMENT_WORKBOOK',str(ROOT/'Installment_System_V2.xlsx'))));identity='local:'+str(store.path);st.sidebar.info('الحفظ المحلي في Excel')
        else:raise Error('mode يجب أن يكون local أو onedrive.')
        st.session_state['store']=store
        if 'raw' not in st.session_state or st.session_state.get('identity')!=identity:
            raw,tag=store.read();st.session_state['raw']=raw;st.session_state['tag']=tag;st.session_state['identity']=identity
        if st.sidebar.button('↻ تحديث من Excel',width='stretch'):
            raw,tag=store.read();st.session_state['raw']=raw;st.session_state['tag']=tag;st.rerun()
        data=read_excel(st.session_state['raw']);snap=snapshot(data)
    except Exception as exc:st.error('تعذر فتح النظام: '+str(exc));st.stop()
    pages=['الملخص','العملاء','المستثمرون','المكاتب','العقود والتحصيل','الخزينة والحركات','متابعة العملاء','التقارير','الإعدادات والنسخ']
    page=st.sidebar.radio('الصفحات',pages);st.title(page)
    if 'notice' in st.session_state:st.success(st.session_state.pop('notice'))
    try:
        if page in ['العملاء','المستثمرون','المكاتب']:people(data,snap,page)
        else:{'الملخص':home,'العقود والتحصيل':contracts_page,'الخزينة والحركات':cash_page,'متابعة العملاء':followups,'التقارير':reports,'الإعدادات والنسخ':settings_page}[page](data,snap)
    except Error as exc:st.error(str(exc))

if __name__=='__main__':main()
