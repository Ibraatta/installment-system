"""Atomic local writes and guarded OneDrive writes. Nothing financial is stored outside XLSX."""
from __future__ import annotations
from pathlib import Path
from io import BytesIO
from uuid import uuid4
from urllib.parse import quote
import hashlib, os, shutil, tempfile
import requests
from filelock import FileLock
import core

MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
GRAPH='https://graph.microsoft.com/v1.0'
def digest(data): return hashlib.sha256(data).hexdigest()
def atomic_write(path,data):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    fd,tmp=tempfile.mkstemp(prefix='.installment-',suffix='.xlsx',dir=path.parent)
    try:
        with os.fdopen(fd,'wb') as f: f.write(data);f.flush();os.fsync(f.fileno())
        os.replace(tmp,path)
    finally:
        if os.path.exists(tmp): os.unlink(tmp)

def backup(path):
    path=Path(path)
    if not path.exists(): return
    folder=path.parent/'backups';folder.mkdir(exist_ok=True)
    target=folder/f'{path.stem}_{core.datetime.now(core.TZ).strftime("%Y%m%d_%H%M%S_%f")}_{uuid4().hex[:6]}.xlsx'
    shutil.copy2(path,target)
    for old in sorted(folder.glob(f'{path.stem}_*.xlsx'),key=lambda p:p.stat().st_mtime,reverse=True)[30:]: old.unlink()

class LocalStore:
    def __init__(self,path): self.path=Path(path)
    def read(self):
        with FileLock(str(self.path)+'.lock',timeout=20):
            data=self.path.read_bytes();core.load(data)
        return data,digest(data)
    def commit(self,expected,action,data,request_id):
        with FileLock(str(self.path)+'.lock',timeout=20):
            old=self.path.read_bytes()
            if digest(old)!=expected: raise core.ConflictError('تغير ملف Excel بعد فتح الشاشة. اضغط تحديث وأعد العملية.')
            wb=core.load(old);result=core.perform(wb,action,data,request_id);new=core.dump(wb)
            backup(self.path);atomic_write(self.path,new)
            return result,new,digest(new)
    def replace(self,expected,data):
        core.load(data)
        with FileLock(str(self.path)+'.lock',timeout=20):
            if digest(self.path.read_bytes())!=expected: raise core.ConflictError('تغير الملف. حدّث قبل الاستعادة.')
            backup(self.path);atomic_write(self.path,data)
        return data,digest(data)

class OneDriveStore:
    def __init__(self,token,folder='Installment System',name='Installment_System.xlsx',item_id=''):
        self.token=token;self.folder=folder.strip('/');self.name=name;self.item_id=item_id
        if Path(name).name!=name or not name.endswith('.xlsx'): raise core.ValidationError('اسم ملف Excel غير صالح.')
        if name=='Personal_Wealth_System_Latest_Integrated.xlsx': raise core.ValidationError('اختر ملف التقسيط المستقل.')
        self.session=requests.Session();self.session.headers.update({'Authorization':f'Bearer {token}'})
        self.path=quote('/'.join(x for x in [self.folder,name] if x),safe='/')
    def _error(self,r):
        if r.status_code in [409,412]: raise core.ConflictError('تغيرت نسخة OneDrive. حدّث من المصدر وأعد العملية.')
        if r.status_code==423: raise core.ValidationError('ملف OneDrive مقفول. أغلق Excel وأعد المحاولة.')
        if r.status_code in [401,403]: raise core.ValidationError('جلسة Microsoft غير صالحة أو الصلاحيات غير كافية. أعد تسجيل الدخول.')
        if not 200<=r.status_code<300:
            raise core.ValidationError(f'تعذر الاتصال بـOneDrive (HTTP {r.status_code}). لم يتم تأكيد الحفظ.')
    def metadata(self):
        url=f'{GRAPH}/me/drive/items/{quote(self.item_id,safe="")}' if self.item_id else f'{GRAPH}/me/drive/root:/{self.path}'
        r=self.session.get(url,params={'$select':'id,name,eTag,lastModifiedDateTime'},timeout=35)
        if r.status_code==404: raise FileNotFoundError(f'لم أجد {self.folder}/{self.name} في OneDrive. ارفع ملف البداية في هذا المجلد أولًا.')
        self._error(r);meta=r.json();self.item_id=meta['id'];return meta
    def read(self):
        before=self.metadata()
        r=self.session.get(f'{GRAPH}/me/drive/items/{quote(self.item_id,safe="")}/content',headers={'Cache-Control':'no-cache'},timeout=90)
        self._error(r);raw=r.content;after=self.metadata()
        if before.get('eTag')!=after.get('eTag'): raise core.ConflictError('تغير الملف أثناء القراءة. اضغط تحديث.')
        core.load(raw)
        if not after.get('eTag'): raise core.ConflictError('OneDrive لم يرجع رقم نسخة؛ تعذر تفعيل الحفظ الآمن.')
        return raw,after['eTag']
    def _publish(self,expected,new):
        # Use a documented conditional upload session to protect the write.
        meta=self.metadata()
        if not expected or meta.get('eTag')!=expected: raise core.ConflictError('تغير ملف OneDrive. حدّث الشاشة قبل الحفظ.')
        r=self.session.post(f'{GRAPH}/me/drive/items/{quote(self.item_id,safe="")}/createUploadSession',
            headers={'If-Match':expected},json={'item':{'@microsoft.graph.conflictBehavior':'fail','name':self.name}},timeout=45)
        self._error(r);upload_url=r.json()['uploadUrl']
        # Upload URLs already contain scoped authorization. Never forward the Graph bearer.
        chunk=320*1024*10;total=len(new);result={}
        for start in range(0,total,chunk):
            part=new[start:start+chunk]
            if start+len(part)==total:
                latest=self.metadata()
                if latest.get('eTag')!=expected:
                    raise core.ConflictError('تغير الملف أثناء تجهيز الحفظ. حدّث المصدر قبل المحاولة مجددًا.')
            try:
                rr=requests.put(upload_url,data=part,headers={'Content-Length':str(len(part)),
                    'Content-Range':f'bytes {start}-{start+len(part)-1}/{total}'},timeout=120)
            except requests.RequestException:
                raise core.ConflictError('لم يصل تأكيد الحفظ. حدّث OneDrive وتحقق من رقم الطلب قبل إعادة الإدخال.')
            self._error(rr)
            if rr.status_code in [200,201]: result=rr.json()
        if not result.get('eTag'):
            # Do not claim success or blindly retry an uncertain commit.
            raise core.ConflictError('لم يصل رقم النسخة بعد الرفع. حدّث الملف وتحقق من سجل الطلبات.')
        return new,result['eTag']
    def commit(self,expected,action,data,request_id):
        old,tag=self.read()
        if tag!=expected: raise core.ConflictError('تغيرت نسخة OneDrive. حدّث وأعد العملية.')
        wb=core.load(old);result=core.perform(wb,action,data,request_id);new=core.dump(wb)
        # Per-session recovery backup; OneDrive additionally keeps its own version history.
        root=Path(tempfile.gettempdir())/'installment_recovery'/hashlib.sha256(self.item_id.encode()).hexdigest()
        root.mkdir(parents=True,exist_ok=True);temp=root/self.name
        with FileLock(str(temp)+'.lock',timeout=20): atomic_write(temp,old);backup(temp)
        raw,newtag=self._publish(expected,new)
        return result,raw,newtag
    def replace(self,expected,data):
        core.load(data)
        old,tag=self.read()
        if tag!=expected: raise core.ConflictError('تغير الملف قبل الاستعادة. حدّث المصدر.')
        root=Path(tempfile.gettempdir())/'installment_recovery'/hashlib.sha256(self.item_id.encode()).hexdigest()
        root.mkdir(parents=True,exist_ok=True)
        target=root/self.name
        with FileLock(str(target)+'.lock',timeout=20): atomic_write(target,old);backup(target)
        return self._publish(expected,data)
