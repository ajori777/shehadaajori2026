from fastapi import FastAPI, HTTPException, Body
from pydantic import BaseModel
import firebase_admin
from firebase_admin import credentials, firestore
import os
from typing import Optional

app = FastAPI()

_db = None

def get_db():
    global _db
    if _db is None:
        try:
            try:
                firebase_admin.get_app()
            except ValueError:
                # 1. Try environment variable (Safer for Render)
                json_creds = os.environ.get("FIREBASE_CREDENTIALS")
                if json_creds:
                    import json
                    creds_dict = json.loads(json_creds)
                    cred = credentials.Certificate(creds_dict)
                    firebase_admin.initialize_app(cred)
                # 2. Try local file
                elif os.path.exists("serviceAccountKey.json"):
                    cred = credentials.Certificate("serviceAccountKey.json")
                    firebase_admin.initialize_app(cred)
                # 3. Default init
                else:
                    firebase_admin.initialize_app()
            _db = firestore.client()
        except Exception as e:
            print(f"Firebase Init Error: {e}")
            return None
    return _db

@app.get("/")
async def health_check():
    database = get_db()
    db_status = "connected" if database is not None else "firebase_error"
    return {
        "status": "online", 
        "db": db_status, 
        "message": "Student Certificate API is running"
    }

class ActivationRequest(BaseModel):
    code: str
    name: str
    hwid: str

@app.post("/activate")
async def activate(req: ActivationRequest):
    database = get_db()
    if not database: raise HTTPException(status_code=500, detail="خطأ في الاتصال بقاعدة البيانات")
    
    code = req.code
    doc_ref = database.collection("certificates_licenses").document(code)
    doc = doc_ref.get()
    
    # محاولة إضافة حرف M إذا لم ينجح البحث الأول وكان الكود لا يبدأ به
    if not doc.exists and not code.startswith('M'):
        code = 'M' + code
        doc_ref = database.collection("certificates_licenses").document(code)
        doc = doc_ref.get()
    
    if not doc.exists:
        raise HTTPException(status_code=404, detail="كود التفعيل غير موجود")
    
    data = doc.to_dict()
    
    if data.get('Status') != 'active':
        raise HTTPException(status_code=403, detail="هذا الكود محظور")
    
    # Check HWID
    stored_hwid = data.get('MachineID')
    if stored_hwid and stored_hwid != req.hwid:
        raise HTTPException(status_code=403, detail="هذا الكود مرتبط بجهاز آخر")
    
    # Bind to this device if not bound
    update_data = {}
    if not stored_hwid:
        update_data['MachineID'] = req.hwid
        update_data['ActivationName'] = req.name
    
    if update_data:
        doc_ref.update(update_data)
        data.update(update_data)
        
    return {
        "status": "active",
        "Name": data.get('Name'),
        "Max": data.get('Max', 0),
        "Used": data.get('Used', 0)
    }

@app.get("/status/{hwid}")
async def get_status(hwid: str):
    database = get_db()
    if not database: raise HTTPException(status_code=500, detail="خطأ في الاتصال بقاعدة البيانات")
    docs = database.collection("certificates_licenses").where("MachineID", "==", hwid).limit(1).get()
    if not docs:
        raise HTTPException(status_code=404, detail="الجهاز غير مفعل")
    
    data = docs[0].to_dict()
    if data.get('Status') != 'active':
        return {"status": "blocked"}
        
    return {
        "status": "active",
        "Name": data.get('Name'),
        "Max": data.get('Max', 0),
        "Used": data.get('Used', 0)
    }

@app.get("/trial_status/{hwid}")
async def get_trial_status(hwid: str):
    database = get_db()
    if not database: raise HTTPException(status_code=500, detail="خطأ في الاتصال بقاعدة البيانات")
    doc_ref = database.collection("trial_users").document(hwid)
    doc = doc_ref.get()
    
    if not doc.exists:
        # First time trial user
        doc_ref.set({"Used": 0, "FirstUsed": firestore.SERVER_TIMESTAMP})
        return {"status": "allowed", "used": 0}
    
    data = doc.to_dict()
    return {"status": "allowed", "used": data.get('Used', 0)}

@app.post("/sync_trial")
async def sync_trial(hwid: str = Body(...), count: int = Body(...)):
    database = get_db()
    if not database: raise HTTPException(status_code=500, detail="خطأ في الاتصال بقاعدة البيانات")
    doc_ref = database.collection("trial_users").document(hwid)
    doc = doc_ref.get()
    
    current_used = 0
    if doc.exists:
        current_used = doc.to_dict().get('Used', 0)
    
    new_used = current_used + count
    doc_ref.set({"Used": new_used}, merge=True)
    
    return {"status": "success", "new_total": new_used}


@app.post("/sync_usage")
async def update_usage(hwid: str = Body(...), count: int = Body(...)):
    database = get_db()
    if not database: raise HTTPException(status_code=500, detail="خطأ في الاتصال بقاعدة البيانات")
    docs = database.collection("certificates_licenses").where("MachineID", "==", hwid).limit(1).get()
    if not docs:
        raise HTTPException(status_code=404, detail="الجهاز غير مفعل")
    
    doc_ref = database.collection("certificates_licenses").document(docs[0].id)
    current_used = docs[0].to_dict().get('Used', 0)
    doc_ref.update({"Used": current_used + count})
    
    return {"status": "success", "new_total": current_used + count}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
