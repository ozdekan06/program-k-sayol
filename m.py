from fastapi import Body, FastAPI, WebSocket, Depends, Query, Request, Form, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pymodbus import FramerType
import serial_asyncio
from sqlalchemy.orm import Session, joinedload
from sqlalchemy import text, or_, func, and_
from sqlalchemy.exc import SQLAlchemyError
from models import UARTLog, Machine, User, Base, SensorLog, Person, CycleLog, TaskLog, Product, Order, TimerLog, CycleEvent
from database import engine, get_db, SessionLocal
from typing import List, Optional
from pydantic import BaseModel
import asyncio
from auth import verify_password
from datetime import datetime
from pymodbus.client import AsyncModbusSerialClient
from backgraound import periodic_sensor_read, read_device_data,cache,client_lock,write_device_data,periodic_sensor_log,periodic_cycle_log,process_cycle_events,run_cycle_log,load_machines,sensor_missing_since,begin_startup_grace,startup_grace_active,temperature_alarm_limits
from connection_api import connection
import connection_api
import io
import pandas as pd
import threading
import os
import shutil
import zipfile
import re
import json
import sqlite3
import tempfile
from xml.sax.saxutils import escape as xml_escape
from pathlib import Path


Base.metadata.create_all(bind=engine)

app = FastAPI()


def _live_machine_status(machine, now=None):
    """Canlı timer bitlerinden ``cooking`` / ``renewal`` / ``idle`` üretir.

    Bu ekranda ``idle`` yalnızca ayarlardan pasif edilen cihaz demektir.
    Aktif cihazlarda discrete[6] start biti 1 ise pişirme, 0 ise yenileme
    fazındadır. Geçmiş CycleEvent kayıtları yalnızca raporlama içindir.
    """
    if not getattr(machine, "is_active", True):
        return "idle"

    # Aktif cihazlar hiçbir zaman "Çalışmayanlar" sütununa düşmemeli.
    # Timer verisi henüz yoksa güvenli varsayım yenileme fazıdır.
    if machine.timer_id is None:
        return "renewal"
    timer_data = cache.get(int(machine.timer_id)) or {}
    discrete = timer_data.get("discrete_inputs") or []
    return "cooking" if len(discrete) > 6 and bool(discrete[6]) else "renewal"

def create_simple_xlsx(rows: list, sheet_name: str = "Sheet1") -> io.BytesIO:
    """Harici Excel motoru yokken temel, geçerli bir XLSX dosyası oluşturur."""
    output = io.BytesIO()
    headers = list(rows[0].keys()) if rows else ["Bilgi"]
    data_rows = [headers] + ([[row.get(header) for header in headers] for row in rows] if rows else [["Kayıt bulunamadı"]])

    def column_name(index):
        name = ""
        while index:
            index, remainder = divmod(index - 1, 26)
            name = chr(65 + remainder) + name
        return name

    sheet_rows = []
    for row_index, values in enumerate(data_rows, start=1):
        cells = []
        for column_index, value in enumerate(values, start=1):
            reference = f"{column_name(column_index)}{row_index}"
            style = ' s="1"' if row_index == 1 else ""
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                cells.append(f'<c r="{reference}"{style}><v>{value}</v></c>')
            else:
                if isinstance(value, datetime):
                    value = value.isoformat(sep=" ", timespec="seconds")
                text_value = "" if value is None else str(value)
                # XML 1.0'da izin verilmeyen kontrol karakterlerini temizle.
                text_value = "".join(char for char in text_value if ord(char) >= 32 or char in "\t\n\r")
                cells.append(f'<c r="{reference}" t="inlineStr"{style}><is><t>{xml_escape(text_value)}</t></is></c>')
        sheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    safe_sheet_name = xml_escape(sheet_name[:31])
    files = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="{safe_sheet_name}" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        "xl/styles.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><b/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0" xfId="0"/><xf fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>',
        "xl/worksheets/sheet1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{"".join(sheet_rows)}</sheetData></worksheet>'
    }
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for path, content in files.items():
            archive.writestr(path, content)
    output.seek(0)
    return output

STATIC_DIR = Path(os.getenv("OZDEKAN_STATIC_DIR", "static")).resolve()
STATIC_DIR.mkdir(parents=True, exist_ok=True)
app.mount("/static", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # veya ["http://127.0.0.1:5500"]
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    
    )

class MachineIn(BaseModel):
    name: str
    sensor_count: int






@app.on_event("startup")
async def startup_event():
    # Ensure DB schema has the is_active column for machines (add if missing)
    try:
        with engine.connect() as conn:
            res = conn.execute(text("PRAGMA table_info(machines);"))
            cols = [r[1] for r in res.fetchall()]
            if 'is_active' not in cols:
                print("is_active sütunu yok — ekleniyor...")
                conn.execute(text("ALTER TABLE machines ADD COLUMN is_active INTEGER DEFAULT 1 NOT NULL;"))
                print("is_active sütunu eklendi.")
            if 'sensor_descriptions' not in cols:
                print("sensor_descriptions sütunu yok — ekleniyor...")
                conn.execute(text("ALTER TABLE machines ADD COLUMN sensor_descriptions JSON DEFAULT '{}' NOT NULL;"))
                print("sensor_descriptions sütunu eklendi.")
            if 'sensor_timeout_minutes' not in cols:
                conn.execute(text("ALTER TABLE machines ADD COLUMN sensor_timeout_minutes INTEGER DEFAULT 5 NOT NULL;"))
            if 'sensor_retry_wait_minutes' not in cols:
                conn.execute(text("ALTER TABLE machines ADD COLUMN sensor_retry_wait_minutes INTEGER DEFAULT 5 NOT NULL;"))
            if 'sensor_retry_window_minutes' not in cols:
                conn.execute(text("ALTER TABLE machines ADD COLUMN sensor_retry_window_minutes INTEGER DEFAULT 1 NOT NULL;"))
            if 'sensor_suspended' not in cols:
                conn.execute(text("ALTER TABLE machines ADD COLUMN sensor_suspended JSON DEFAULT '{}' NOT NULL;"))
            for column, definition in {
                'active_order_number': 'TEXT', 'active_person_id': 'INTEGER',
                'active_mold_count': 'INTEGER', 'active_lot_number': 'TEXT'
            }.items():
                if column not in cols:
                    conn.execute(text(f"ALTER TABLE machines ADD COLUMN {column} {definition};"))
            conn.commit()

        # Eski veritabanlarına yeni sipariş alanlarını ekle.
        with engine.connect() as conn:
            res = conn.execute(text("PRAGMA table_info(orders);"))
            order_cols = [r[1] for r in res.fetchall()]
            if 'unit' not in order_cols:
                print("orders.unit sütunu yok — ekleniyor...")
                conn.execute(text("ALTER TABLE orders ADD COLUMN unit TEXT;"))
                conn.commit()
            additions = {
                'volume_per_piece': 'REAL',
                'is_manually_completed': 'INTEGER DEFAULT 0 NOT NULL',
                'completed_at': 'DATETIME',
                'completion_note': 'TEXT'
            }
            for column, definition in additions.items():
                if column not in order_cols:
                    conn.execute(text(f"ALTER TABLE orders ADD COLUMN {column} {definition};"))
            conn.commit()

        # Lot numarası siparişten bağımsız olarak ürün bazında değiştirilebilir.
        with engine.connect() as conn:
            res = conn.execute(text("PRAGMA table_info(product);"))
            product_cols = [r[1] for r in res.fetchall()]
            if 'lot_number' not in product_cols:
                print("product.lot_number sütunu yok — ekleniyor...")
                conn.execute(text("ALTER TABLE product ADD COLUMN lot_number TEXT;"))
                conn.commit()
                print("product.lot_number sütunu eklendi.")
            product_additions = {
                'status_change': 'TEXT',
                'status_change_reason': 'TEXT',
                'status_changed_at': 'DATETIME',
                'is_reverted': 'INTEGER DEFAULT 0 NOT NULL',
                'reverted_at': 'DATETIME',
                'revert_reason': 'TEXT'
            }
            for column, definition in product_additions.items():
                if column not in product_cols:
                    conn.execute(text(f"ALTER TABLE product ADD COLUMN {column} {definition};"))
            conn.commit()

        # Büyük geçmişte ürün/cycle bağlantılarını ve son görevleri hızlı bul.
        with engine.begin() as conn:
            conn.execute(text("CREATE INDEX IF NOT EXISTS ix_perf_tasks_product_latest ON task_logs (product_id, timestamp, id)"))
            conn.execute(text("CREATE INDEX IF NOT EXISTS ix_perf_tasks_cycle_product ON task_logs (cycle_id, product_id)"))
            conn.execute(text("CREATE INDEX IF NOT EXISTS ix_perf_product_order_reverted ON product (ordernumber, is_reverted)"))

        with engine.connect() as conn:
            cycle_cols = [r[1] for r in conn.execute(text("PRAGMA table_info(cycle_logs);"))]
            for column, definition in {
                'is_valid': 'INTEGER DEFAULT 1 NOT NULL', 'invalid_reason': 'TEXT',
                'invalidated_at': 'DATETIME', 'closure_type': "TEXT DEFAULT 'automatic' NOT NULL",
                'order_number': 'TEXT', 'person_id': 'INTEGER', 'mold_count': 'INTEGER', 'lot_number': 'TEXT'
            }.items():
                if column not in cycle_cols:
                    conn.execute(text(f"ALTER TABLE cycle_logs ADD COLUMN {column} {definition};"))
            conn.commit()
    except Exception as e:
        print(f"is_active sütunu kontrol edilirken hata: {e}")

    # Askilar otomatik timeout sonucudur; yeni program oturumunda cihazlara
    # yeniden cevap verme sansi tanimak icin kalici eski durumu temizle.
    startup_db = SessionLocal()
    try:
        for machine in startup_db.query(Machine).all():
            if getattr(machine, 'sensor_suspended', {}) or {}:
                machine.sensor_suspended = {}
                startup_db.add(machine)
        startup_db.commit()
    finally:
        startup_db.close()
    begin_startup_grace()
    load_machines()  # Makineleri yükle
    await connection()
    print(connection_api.modbus_client.connected)
    #threading.Thread(target=run_cycle_log, daemon=True).start()
    asyncio.create_task(periodic_cycle_log())
    asyncio.create_task(process_cycle_events())
    asyncio.create_task(periodic_sensor_read())
    asyncio.create_task(periodic_sensor_log())
    print("Starting Modbus client...")



@app.on_event("shutdown")
async def shutdown_event():
    await connection_api.close()



@app.get("/read-device")
async def read_device(m: str = Query(...), db: Session = Depends(get_db)):
    if not m.startswith("m"):
        raise HTTPException(status_code=400, detail="Makine ID 'm' ile başlamalıdır")
    try:
        slave_id = int(m[1:])
    except:
        raise HTTPException(status_code=400, detail="Geçersiz makine ID")

    # Askıya alınmış sensörler, istemciden gelen doğrudan okuma taleplerinde de engellenir.
    suspended_machine = None
    for candidate in db.query(Machine).all():
        if str(slave_id) in [str(value) for value in (candidate.sensor_ids or [])]:
            suspended_machine = candidate
            break
    if suspended_machine and (getattr(suspended_machine, 'sensor_suspended', {}) or {}).get(str(slave_id)):
        raise HTTPException(status_code=423, detail="Sensör askıya alınmış. 'Sensörü Aktif Et' ile yeniden başlatın.")
    
    # Cache’den dönmek istiyorsan:
    if slave_id in cache:
        return {
            "slave_id": slave_id,
            "holding": cache[slave_id].get("holding_registers", []),
            "input": cache[slave_id].get("input_registers", []),
            "read_at": cache[slave_id].get("read_at"),
            "holding_read_at": cache[slave_id].get("holding_read_at"),
            "input_read_at": cache[slave_id].get("input_read_at"),
            "discrete_read_at": cache[slave_id].get("discrete_read_at"),
            "coils_read_at": cache[slave_id].get("coils_read_at"),
            "quality": cache[slave_id].get("quality"),
            "latency_ms": cache[slave_id].get("latency_ms"),
            "sequence": cache[slave_id].get("sequence"),
            "last_error": cache[slave_id].get("last_error"),
        }
    
    # Cache yoksa doğrudan oku (opsiyonel)
    data = await read_device_data(slave_id)
    print(data)
    if not data:
        raise HTTPException(status_code=404, detail="Veri alınamadı")
    return {
        "slave_id": slave_id,
        "holding": data.get("holding_registers", []),
        "input": data.get("input_registers", []),
        "read_at": data.get("read_at"),
        "holding_read_at": data.get("holding_read_at"),
        "input_read_at": data.get("input_read_at"),
        "discrete_read_at": data.get("discrete_read_at"),
        "coils_read_at": data.get("coils_read_at")
    }


@app.post("/login")
def login(username: str = Form(...), password: str = Form(...), db: Session = Depends(get_db)):
    print("seleaamm")
    user = db.query(User).filter(User.username == username).first()
    if not user or not verify_password(password, user.hashed_password):
        raise HTTPException(status_code=401, detail="Geçersiz kullanıcı adı veya şifre")
    return {"message": "Giriş başarılı"}


@app.get("/logs/chart")
def get_chart_data(limit: int = 100, db: Session = Depends(get_db)):
    logs = db.query(UARTLog).order_by(UARTLog.timestamp.desc()).limit(limit).all()
    return [
        {
            "timestamp": log.timestamp.isoformat(),
            "direction": log.direction,
            "data": len(log.data)
        }
        for log in reversed(logs)
    ]

@app.get("/read-all")
async def read_all_devices(db: Session = Depends(get_db)):
    data = {}
    active_machines = db.query(Machine).filter(Machine.is_active.is_(True)).all()
    device_ids = []
    for machine in active_machines:
        device_ids.extend(int(sensor_id) for sensor_id in (machine.sensor_ids or []))
        if machine.timer_id is not None:
            device_ids.append(int(machine.timer_id))

    for device_id in dict.fromkeys(device_ids):
        # API trafiği RS-485 polling sırasına ek yük bindirmesin.
        registers = cache.get(device_id)
        data[f"device_{device_id}"] = {
            "holding_registers": registers.get("holding_registers", []) if registers else [],
            "input_registers": registers.get("input_registers", []) if registers else [],
            "holding_read_at": registers.get("holding_read_at") if registers else None,
            "input_read_at": registers.get("input_read_at") if registers else None,
            "quality": registers.get("quality") if registers else "NO_DATA",
            "latency_ms": registers.get("latency_ms") if registers else None,
        }
    return data


@app.get("/modbus-health")
async def modbus_health():
    return {
        "bus": connection_api.get_metrics(),
        "requests": connection_api.get_request_metrics(),
        "devices": {
            str(slave_id): {
                "quality": values.get("quality", "NO_DATA"),
                "read_at": values.get("read_at"),
                "latency_ms": values.get("latency_ms"),
                "sequence": values.get("sequence", 0),
                "consecutive_errors": values.get("consecutive_errors", 0),
                "last_error": values.get("last_error"),
            }
            for slave_id, values in cache.items()
        },
    }

#@app.get("/read-device")
#async def read_device(m: str = Query(..., description="Makine ID (örnek: m1)")):
    if not m.startswith("m"):
        raise HTTPException(status_code=400, detail="Makine ID 'm' ile başlamalıdır. Örn: m1")
    try:
        slave_id = int(m[1:])
    except ValueError:
        raise HTTPException(status_code=400, detail="Makine ID geçerli bir sayı içermiyor.")
    data = await read_device_data(slave_id)
    if not data:
        raise HTTPException(status_code=404, detail=f"Slave ID {slave_id} için veri alınamadı.")
    return {
        "slave_id": slave_id,
        "holding": data.get("holding_registers", []),
        "input": data.get("input_registers", [])
    }


@app.get("/write-device")
async def write_device(
    m: str = Query(..., description="Makine adı. Örnek: m12"),
    address: int = Query(..., description="Modbus register adresi"),
    value: int = Query(..., description="Yazılacak integer değer")
    ):
    try:
        if not m.startswith("m"):
            raise ValueError("Makine adı 'm' ile başlamalı. Örnek: m5")

        slave_id = int(m[1:])
        success = await write_device_data(slave_id=slave_id, address=address, value=value)

        if success:
            return {"status": "ok", "slave_id": slave_id, "address": address, "value": value}
        else:
            raise HTTPException(status_code=500, detail="Yazma başarısız")
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/setup-machines")
def setup_machines(machines: List[MachineIn], db: Session = Depends(get_db)):
    db.query(Machine).delete()
    for m in machines:
        db.add(Machine(name=m.name, sensor_count=m.sensor_count))
    db.commit()
    return {"message": "Ayarlar kaydedildi"}

@app.api_route("/reset-db", methods=["GET", "POST"])
def reset_database(request: Request):
    if request.method != "POST":
        return {"message": "Lütfen POST isteği gönderin."}
    Base.metadata.drop_all(bind=engine)
    Base.metadata.create_all(bind=engine)
    return {"message": "Veritabanı sıfırlandı."}

@app.get("/machines")
def get_machines(db: Session = Depends(get_db)):
    machines = db.query(Machine).order_by(Machine.id.asc()).all()
    return [
        {
            "id": m.id,
            "name": m.name,
            "slave_id": m.slave_id,
            "sensor_ids": m.sensor_ids,
            "sensor_descriptions": m.sensor_descriptions or {},
            "sensor_count": m.sensor_count,
            "image": m.image,
            "timer_id": m.timer_id,
            "is_active": getattr(m, 'is_active', True)
            , "sensor_timeout_minutes": getattr(m, 'sensor_timeout_minutes', 5)
            , "sensor_retry_wait_minutes": getattr(m, 'sensor_retry_wait_minutes', 5)
            , "sensor_retry_window_minutes": getattr(m, 'sensor_retry_window_minutes', 1)
            , "sensor_suspended": getattr(m, 'sensor_suspended', {}) or {}
            , "active_order_number": getattr(m, 'active_order_number', None)
            , "active_person_id": getattr(m, 'active_person_id', None)
            , "active_mold_count": getattr(m, 'active_mold_count', None)
            , "active_lot_number": getattr(m, 'active_lot_number', None)
        }
        for m in machines
    ]

@app.get("/machines/status")
def get_machines_status(db: Session = Depends(get_db)):
    """Her makine için çalışma durumunu belirler: 'cooking', 'renewal', 'idle'"""
    machines = db.query(Machine).order_by(Machine.id.asc()).all()
    result = []
    suspension_changed = False
    for m in machines:
        is_active = getattr(m, 'is_active', True)
        sensor_alarm = False
        sensor_timeout = False
        timeout_seconds = max(1, int(getattr(m, 'sensor_timeout_minutes', 5) or 5)) * 60
        import datetime as _dt
        # Pasif makine okunmaz; eski cache/askiya alma bilgisi de ekranda
        # alarm veya timeout olarak gosterilmemelidir.
        for sensor_id in (m.sensor_ids or []) if is_active else []:
            if (getattr(m, 'sensor_suspended', {}) or {}).get(str(sensor_id)):
                sensor_timeout = True
                continue
            data = cache.get(int(sensor_id)) or {}
            read_at = data.get('input_read_at')
            sensor_stale = True
            if read_at:
                try:
                    sensor_stale = ((_dt.datetime.now(_dt.timezone.utc) - _dt.datetime.fromisoformat(read_at)).total_seconds() > timeout_seconds)
                    sensor_stale = sensor_stale and not startup_grace_active()
                    sensor_timeout = sensor_timeout or sensor_stale
                except Exception: pass
            else:
                missing_key = (m.id, int(sensor_id))
                first_missing = sensor_missing_since.setdefault(missing_key, _dt.datetime.now(_dt.timezone.utc))
                missing_too_long = ((_dt.datetime.now(_dt.timezone.utc) - first_missing).total_seconds() > timeout_seconds)
                sensor_timeout = missing_too_long and not startup_grace_active()
                sensor_stale = False  # İlk okuma henüz gerçekleşmediyse sensörü kilitleme.
            inputs, holding = data.get('input_registers') or [], data.get('holding_registers') or []
            alarm_limits = temperature_alarm_limits(holding)
            if inputs and alarm_limits:
                alarm_lower, alarm_upper = alarm_limits
                sensor_alarm = sensor_alarm or inputs[0] < alarm_lower or inputs[0] > alarm_upper
            if sensor_stale:
                suspended = dict(getattr(m, 'sensor_suspended', {}) or {})
                if not suspended.get(str(sensor_id)):
                    suspended[str(sensor_id)] = True
                    m.sensor_suspended = suspended
                    db.add(m)
                    suspension_changed = True
        # Ekran durumu geçmiş CycleEvent'ten değil canlı timer bitlerinden gelir.
        status = _live_machine_status(m)

        result.append({
            'id': m.id,
            'name': m.name,
            'slave_id': m.slave_id,
            'is_active': is_active,
            'status': status,
            'sensor_alarm': sensor_alarm
            , 'sensor_timeout': sensor_timeout,
            'sensor_suspended': getattr(m, 'sensor_suspended', {}) or {}
        })
    if suspension_changed:
        db.commit()
        load_machines()
    return result


@app.post("/machines")
def create_machine(payload: dict = Body(...), db: Session = Depends(get_db)):
    try:
        m = Machine(
            name=payload.get("name"),
            slave_id=payload.get("slave_id"),
            sensor_ids=payload.get("sensor_ids") or [],
            sensor_descriptions=payload.get("sensor_descriptions") or {},
            sensor_count=payload.get("sensor_count") or 0,
            timer_id=payload.get("timer_id"),
            image=payload.get("image") or None,
            is_active=payload.get("is_active", True)
            , sensor_timeout_minutes=max(1, int(payload.get("sensor_timeout_minutes") or 5))
            , sensor_retry_wait_minutes=max(1, int(payload.get("sensor_retry_wait_minutes") or 5))
            , sensor_retry_window_minutes=max(1, int(payload.get("sensor_retry_window_minutes") or 1))
            , sensor_suspended=payload.get("sensor_suspended") or {}
            , active_order_number=payload.get("active_order_number") or None
            , active_person_id=payload.get("active_person_id") or None
            , active_mold_count=payload.get("active_mold_count") or None
            , active_lot_number=payload.get("active_lot_number") or None
        )
        db.add(m)
        db.commit()
        db.refresh(m)
        # Refresh in-memory machines for background tasks
        try:
            load_machines()
        except Exception:
            pass
        return {"id": m.id}
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=400, detail=str(e))


@app.put("/machines/{machine_id}")
def update_machine(machine_id: int, payload: dict = Body(...), db: Session = Depends(get_db)):
    machine = db.query(Machine).filter(Machine.id == machine_id).first()
    if not machine:
        raise HTTPException(status_code=404, detail="Makine bulunamadı")
    try:
        old_suspended = dict(getattr(machine, 'sensor_suspended', {}) or {})
        was_active = getattr(machine, 'is_active', True)
        reactivating = payload.get('is_active') is True and not was_active
        # If attempting to deactivate, check for ongoing cycles unless forced
        force = bool(payload.get('force', False))
        if 'is_active' in payload and payload.get('is_active') is False and getattr(machine, 'is_active', True):
            # check for open cycle where cooking is still ongoing
            open_cycle = db.query(CycleLog).filter(CycleLog.machine_id == machine_id, CycleLog.cycle_end_time == None).order_by(CycleLog.cycle_start_time.desc()).first()
            if open_cycle and (open_cycle.cooking_duration is None) and not force:
                # return warning so UI can confirm
                return JSONResponse(status_code=409, content={
                    "warning": "Bu makinede halen bir pişirme döngüsü devam ediyor. Pasif yapmak istediğinizden emin misiniz?",
                    "cycle_id": open_cycle.id,
                    "cycle_start": open_cycle.cycle_start_time.isoformat() if open_cycle.cycle_start_time else None
                })

        if "sensor_descriptions" in payload and not isinstance(payload.get("sensor_descriptions"), dict):
            raise ValueError("Sensör açıklamaları geçerli bir nesne olmalıdır")

        for key in ("name", "slave_id", "sensor_ids", "sensor_descriptions", "sensor_count", "timer_id", "image", "is_active", "sensor_timeout_minutes", "sensor_retry_wait_minutes", "sensor_retry_window_minutes", "sensor_suspended", "active_order_number", "active_person_id", "active_mold_count", "active_lot_number"):
            if key in payload:
                value = payload.get(key)
                if key in {"sensor_timeout_minutes", "sensor_retry_wait_minutes", "sensor_retry_window_minutes"}:
                    value = max(1, int(value or 1))
                setattr(machine, key, value)
        if reactivating:
            # Pasif kaldigi sure boyunca biriken eski aski/cache durumu yeni
            # aktif oturuma tasinmasin; ilk polling turu temiz baslasin.
            machine.sensor_suspended = {}
            device_ids = {int(value) for value in (machine.sensor_ids or [])}
            if machine.timer_id is not None:
                device_ids.add(int(machine.timer_id))
            for device_id in device_ids:
                cache.pop(device_id, None)
                sensor_missing_since.pop((machine.id, device_id), None)
        if "sensor_suspended" in payload:
            new_suspended = payload.get("sensor_suspended") or {}
            for sensor_key, was_suspended in old_suspended.items():
                if was_suspended and not new_suspended.get(str(sensor_key)):
                    sensor_id = int(sensor_key)
                    cache.pop(sensor_id, None)
                    sensor_missing_since.pop((machine.id, sensor_id), None)
        db.add(machine)
        db.commit()
        # Refresh in-memory machines for background tasks
        try:
            load_machines()
        except Exception:
            pass
        return {"message": "Güncellendi"}
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=400, detail=str(e))


@app.delete("/machines/{machine_id}")
def delete_machine(machine_id: int, db: Session = Depends(get_db)):
    machine = db.query(Machine).filter(Machine.id == machine_id).first()
    if not machine:
        raise HTTPException(status_code=404, detail="Makine bulunamadı")
    try:
        db.delete(machine)
        db.commit()
        try:
            load_machines()
        except Exception:
            pass
        return {"message": "Silindi"}
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/upload-machine-image/{machine_id}")
async def upload_machine_image(
    machine_id: int,
    file: UploadFile = File(...),
    db: Session = Depends(get_db)
):
    """Makine resmi yükle ve veritabanına kaydet"""
    try:
        # Makineyi kontrol et
        machine = db.query(Machine).filter(Machine.id == machine_id).first()
        if not machine:
            raise HTTPException(status_code=404, detail="Makine bulunamadı")
        
        # Upload klasörünü oluştur
        upload_dir = STATIC_DIR / "uploads" / "machines"
        upload_dir.mkdir(parents=True, exist_ok=True)
        
        # Dosya uzantısını kontrol et
        file_ext = os.path.splitext(file.filename)[1].lower()
        if file_ext not in [".jpg", ".jpeg", ".png", ".gif", ".webp"]:
            raise HTTPException(status_code=400, detail="Sadece resim dosyaları yüklenebilir (.jpg, .png, .gif, .webp)")
        
        # Yeni dosya adı oluştur
        new_filename = f"machine_{machine_id}{file_ext}"
        file_path = upload_dir / new_filename
        
        # Eski resmi sil (varsa)
        if machine.image and os.path.exists(machine.image):
            try:
                os.remove(machine.image)
            except:
                pass
        
        # Yeni resmi kaydet
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Veritabanını güncelle
        machine.image = str(file_path)
        db.commit()
        
        return {
            "message": "Resim başarıyla yüklendi",
            "image_path": str(file_path),
            "machine_id": machine_id
        }
    
    except HTTPException:
        raise
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"Resim yüklenirken hata oluştu: {str(e)}")

@app.get("/machine-image/{machine_id}")
async def get_machine_image(machine_id: int, db: Session = Depends(get_db)):
    """Makine resmini getir"""
    machine = db.query(Machine).filter(Machine.id == machine_id).first()
    if not machine or not machine.image:
        raise HTTPException(status_code=404, detail="Resim bulunamadı")
    
    if not os.path.exists(machine.image):
        raise HTTPException(status_code=404, detail="Resim dosyası bulunamadı")
    
    return FileResponse(machine.image)


@app.delete("/machine-image/{machine_id}")
def delete_machine_image(machine_id: int, db: Session = Depends(get_db)):
    """Makineye ait resmi sil ve veritabanındaki image alanını temizle"""
    machine = db.query(Machine).filter(Machine.id == machine_id).first()
    if not machine:
        raise HTTPException(status_code=404, detail="Makine bulunamadı")
    if not machine.image:
        return {"message": "Resim zaten yok"}
    try:
        image_path = machine.image
        # Clear DB field first
        machine.image = None
        db.add(machine)
        db.commit()
        # Then remove file if exists
        try:
            if image_path and os.path.exists(image_path):
                os.remove(image_path)
        except Exception:
            # If file deletion fails, we don't fail the whole request
            pass
        return {"message": "Resim kaldırıldı"}
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/save-machine-settings")
async def save_machine_settings(request: Request, db: Session = Depends(get_db)):
    try:
        machines = await request.json()
        for m in machines:
            machine = Machine(
                name=m["name"],
                slave_id=m["slave_id"],
                sensor_count=m["sensor_count"],
                sensor_ids=m["sensor_ids"],
                timer_id=m["timer_id"],
            )
            db.add(machine)
        db.commit()
        return {"message": "Ayarlar başarıyla kaydedildi."}
    except SQLAlchemyError as e:
        db.rollback()
        return JSONResponse(status_code=500, content={"message": f"Veritabanı hatası: {str(e)}"})
    except Exception as e:
        return JSONResponse(status_code=400, content={"message": f"Gönderim hatası: {str(e)}"})

class UARTProtocol(asyncio.Protocol):
    def __init__(self, websocket: WebSocket):
        self.websocket = websocket
        self.transport = None

    def connection_made(self, transport):
        self.transport = transport

    def data_received(self, data):
        message = data.decode(errors="ignore").strip()
        asyncio.create_task(self.websocket.send_text(message))
        # ✅ Celery üzerinden log kaydı




# Windows için COM port ayarı - gerekirse değiştir
UART_PORT = "COM1"  # Windows'ta genelde COM1, COM3, COM4 vb.
BAUDRATE = 9600
@app.websocket("/ws")
async def websocket_uart(websocket: WebSocket):
    await websocket.accept()
    loop = asyncio.get_running_loop()
    try:
        transport, protocol = await serial_asyncio.create_serial_connection(
            loop, lambda: UARTProtocol(websocket), UART_PORT, BAUDRATE
        )
    except Exception as e:
        await websocket.send_text(f"HATA: UART bağlantısı sağlanamadı: {e}")
        await websocket.close()
        return

    try:
        while True:
            msg = await websocket.receive_text()
            # Celery ile loglama
            transport.write(msg.encode())
    except asyncio.CancelledError:
        # Websocket bağlantısı iptal edildiğinde temiz kapatma
        pass
    except Exception as e:
        # Diğer hatalar için loglama yapabiliriz
        print(f"Websocket hata: {e}")
    finally:
        if transport:
            transport.close()
        await websocket.close()

@app.get("/sensor-logs/export")
def export_sensor_logs(
    machine_id: int = Query(None, description="Makine ID"),
    sensor_id: int = Query(None, description="Sensör ID"),
    start: Optional[str] = Query(None, description="Başlangıç tarihi (YYYY-MM-DD)"),
    end: Optional[str] = Query(None, description="Bitiş tarihi (YYYY-MM-DD)"),
    db: Session = Depends(get_db)
):
    """
    SensorLog tablosundan verileri Excel olarak export eder.
    """
    query = db.query(SensorLog)
    if machine_id is not None:
        query = query.filter(SensorLog.machine_id == machine_id)
    if sensor_id is not None:
        query = query.filter(SensorLog.sensor_id == sensor_id)
    if start:
        query = query.filter(SensorLog.timestamp >= start)
    if end:
        query = query.filter(SensorLog.timestamp <= end)
    logs = query.order_by(SensorLog.timestamp.asc()).all()
    data = [
        {
            "ID": log.id,
            "Makine ID": log.machine_id,
            "Sensör ID": log.sensor_id,
            "Tarih": log.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
            "Değer": log.value
        }
        for log in logs
    ]
    output = create_simple_xlsx(data, "SensorLogs")
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=sensor_logs.xlsx"}
    )

@app.post("/sensor-logs/export")
def export_sensor_logs_by_ids(payload: dict = Body(...), db: Session = Depends(get_db)):
    filtered = bool(payload.get("filtered"))
    log_ids = payload.get("log_ids") or []
    query = db.query(SensorLog).order_by(SensorLog.timestamp.asc())
    if filtered:
        query = query.filter(SensorLog.id.in_(log_ids)) if log_ids else query.filter(text("1=0"))
    rows = [{
        "ID": log.id, "Makine ID": log.machine_id, "Sensör ID": log.sensor_id,
        "Tarih": log.timestamp, "Ölçülen": log.value, "Set": log.value_set,
        "Histerezis": log.value_histerezis, "Sapma": log.value - log.value_set
    } for log in query.all()]
    output = create_simple_xlsx(rows, "Sensor Loglari")
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=sensor_loglari.xlsx"})

@app.post("/sensor-logs/delete")
def delete_sensor_logs(payload: dict = Body(...), db: Session = Depends(get_db)):
    log_ids = payload.get("log_ids") or []
    if not log_ids:
        raise HTTPException(status_code=400, detail="Silinecek sensör logu seçilmedi")
    deleted = db.query(SensorLog).filter(SensorLog.id.in_(log_ids)).delete(synchronize_session=False)
    db.commit()
    return {"message": f"{deleted} sensör logu silindi", "deleted": deleted}

@app.get("/persons")
def get_persons(db: Session = Depends(get_db)):
    """
    Person tablosundaki tüm kişileri getirir.
    """
    persons = db.query(Person).all()
    return [
        {
            "id": person.id,
            "name": person.name,
            "surname": person.surname,
            "age": person.age,
            "full_name": f"{person.name} {person.surname}"
        }
        for person in persons
    ]


@app.post("/persons")
def create_person(person: dict = Body(...), db: Session = Depends(get_db)):
    try:
        name = person.get("name")
        surname = person.get("surname")
        age = person.get("age")
        if not name or not surname:
            raise HTTPException(status_code=400, detail="Name and surname required")
        p = Person(name=name, surname=surname, age=age)
        db.add(p)
        db.commit()
        db.refresh(p)
        return {"id": p.id, "full_name": f"{p.name} {p.surname}"}
    except HTTPException:
        raise
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=str(e))


@app.delete("/persons/{person_id}")
def delete_person(person_id: int, db: Session = Depends(get_db)):
    person = db.query(Person).filter(Person.id == person_id).first()
    if not person:
        raise HTTPException(status_code=404, detail="Person not found")
    try:
        db.delete(person)
        db.commit()
        return {"message": "Deleted"}
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/system/config-export")
def export_system_config(db: Session = Depends(get_db)):
    payload = {
        "format": "ozdekan-config-v1",
        "exported_at": datetime.now().isoformat(),
        "machines": [{
            "name": m.name, "slave_id": m.slave_id, "sensor_ids": m.sensor_ids,
            "sensor_descriptions": m.sensor_descriptions or {}, "sensor_count": m.sensor_count,
            "timer_id": m.timer_id, "is_active": m.is_active,
            "sensor_timeout_minutes": m.sensor_timeout_minutes,
            "sensor_retry_wait_minutes": m.sensor_retry_wait_minutes,
            "sensor_retry_window_minutes": m.sensor_retry_window_minutes
        } for m in db.query(Machine).order_by(Machine.id).all()],
        "persons": [{"name": p.name, "surname": p.surname, "age": p.age}
                    for p in db.query(Person).order_by(Person.id).all()]
    }
    raw = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    return StreamingResponse(io.BytesIO(raw), media_type="application/json",
        headers={"Content-Disposition": "attachment; filename=ozdekan-config.json"})

@app.post("/system/config-import")
async def import_system_config(file: UploadFile = File(...), db: Session = Depends(get_db)):
    try:
        payload = json.loads((await file.read()).decode("utf-8"))
        if payload.get("format") != "ozdekan-config-v1":
            raise ValueError("Geçersiz yapılandırma dosyası")
        for item in payload.get("persons", []):
            person = db.query(Person).filter(Person.name == item.get("name"), Person.surname == item.get("surname")).first()
            if not person:
                db.add(Person(name=item.get("name"), surname=item.get("surname"), age=item.get("age")))
        for item in payload.get("machines", []):
            machine = db.query(Machine).filter(Machine.name == item.get("name")).first()
            if not machine:
                machine = Machine(name=item.get("name"), sensor_ids=[], sensor_count=0)
                db.add(machine)
            for key in ("slave_id", "sensor_ids", "sensor_descriptions", "sensor_count", "timer_id", "is_active", "sensor_timeout_minutes", "sensor_retry_wait_minutes", "sensor_retry_window_minutes"):
                if key in item:
                    setattr(machine, key, item[key])
        db.commit()
        load_machines()
        return {"message": "Yapılandırma birleştirilerek içe aktarıldı; mevcut kayıtlar silinmedi"}
    except (ValueError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        db.rollback()
        raise HTTPException(status_code=400, detail=str(exc))

@app.get("/system/backup")
def download_database_backup():
    source_path = Path(engine.url.database)
    with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as temp:
        backup_path = Path(temp.name)
    try:
        with sqlite3.connect(source_path) as source, sqlite3.connect(backup_path) as target:
            source.backup(target)
        raw = backup_path.read_bytes()
    finally:
        backup_path.unlink(missing_ok=True)
    filename = f"ozdekan-backup-{datetime.now().strftime('%Y%m%d-%H%M%S')}.db"
    return StreamingResponse(io.BytesIO(raw), media_type="application/x-sqlite3",
        headers={"Content-Disposition": f"attachment; filename={filename}"})
@app.get("/cycle-logs")
def get_cycle_logs(include_invalid: bool = False, db: Session = Depends(get_db), pending_only: bool = False):
    """
    Tüm cycle loglarını getirir.
    """
    query = db.query(CycleLog).join(Machine, CycleLog.machine_id == Machine.id).options(joinedload(CycleLog.machine))
    if not include_invalid:
        query = query.filter(CycleLog.is_valid.is_(True))
    if pending_only:
        linked_product = db.query(TaskLog.id).join(Product, TaskLog.product_id == Product.id).filter(
            TaskLog.cycle_id == CycleLog.id, Product.is_reverted.is_(False)
        ).exists()
        query = query.filter(~linked_product)
    cycles = query.order_by(CycleLog.cycle_start_time.asc()).all()
    return [
        {
            "id": c.id,
            "machine_id": c.machine_id,
            "machine_name": c.machine.name if c.machine else f"Makine {c.machine_id}",
            "timestamp": c.timestamp.isoformat() if c.timestamp else None,
            "cycle_start_time": c.cycle_start_time.isoformat() if c.cycle_start_time else None,
            "cycle_end_time": c.cycle_end_time.isoformat() if c.cycle_end_time else None,
            "cooking_duration": c.cooking_duration,
            "renewal_duration": c.renewal_duration,
            "total_duration": c.total_duration,
            "ideal_cooking_time": c.ideal_cooking_time,
            "ideal_renewal_time": c.ideal_renewal_time,
            "ideal_total_time": c.ideal_total_time,
            "cooking_delay": c.cooking_delay,
            "renewal_delay": c.renewal_delay,
            "total_delay": c.total_delay,
            "is_valid": c.is_valid,
            "invalid_reason": c.invalid_reason,
            "closure_type": c.closure_type,
            "machine_context": {
                "order_number": c.order_number,
                "person_id": c.person_id,
                "mold_count": c.mold_count,
                "lot_number": c.lot_number
            }
        }
        for c in cycles
    ]

@app.post("/cycle-logs/{cycle_id}/invalidate")
def invalidate_cycle(cycle_id: int, payload: dict = Body(...), db: Session = Depends(get_db)):
    cycle = db.query(CycleLog).filter(CycleLog.id == cycle_id).first()
    if not cycle:
        raise HTTPException(status_code=404, detail="Cycle bulunamadı")
    if db.query(TaskLog.id).filter(TaskLog.cycle_id == cycle_id, TaskLog.status != "REVERTED").first():
        raise HTTPException(status_code=409, detail="Ürüne bağlı cycle önce geri alınmalıdır")
    reason = str(payload.get("reason") or "").strip()
    if not reason:
        raise HTTPException(status_code=400, detail="Geçersizleştirme nedeni zorunludur")
    cycle.is_valid, cycle.invalid_reason, cycle.invalidated_at = False, reason, datetime.now()
    db.commit()
    return {"message": "Cycle silinmedi; geçersiz işaretlendi"}

@app.post("/cycle-logs/{cycle_id}/close-shift")
def close_cycle_for_shift(cycle_id: int, db: Session = Depends(get_db)):
    cycle = db.query(CycleLog).filter(CycleLog.id == cycle_id, CycleLog.cycle_end_time.is_(None)).first()
    if not cycle:
        raise HTTPException(status_code=404, detail="Açık cycle bulunamadı")
    cycle.cycle_end_time = datetime.now()
    if cycle.cycle_start_time:
        cycle.total_duration = max(0, (cycle.cycle_end_time - cycle.cycle_start_time).total_seconds() / 60)
        cycle.total_delay = cycle.total_duration - (cycle.ideal_total_time or 0)
    cycle.closure_type = "shift_end"
    db.commit()
    return {"message": "Cycle vardiya sonu olarak kapatıldı"}
@app.get("/product-logs")
def get_product_logs(
    ordernumber: Optional[str] = Query(None, description="Sipariş numarası"),
    machine_id: int = Query(None, description="Makine ID"),
    person_id: int = Query(None, description="Personel ID"),
    serialnumber: Optional[str] = Query(None, description="Seri numarası"),
    product_name: Optional[str] = Query(None, description="Ürün adı"),
    productsize: Optional[str] = Query(None, description="Ebat"),
    producttype: Optional[str] = Query(None, description="Tip"),
    moldnumber: Optional[str] = Query(None, description="Kalıp numarası"),
    lot_number: Optional[str] = Query(None, description="Lot numarası"),
    status: Optional[str] = Query(None, description="Ürün durumu"),
    cycle_id: int = Query(None, description="Cycle ID"),
    min_total_duration: Optional[float] = Query(None, description="Minimum toplam cycle süresi"),
    max_total_duration: Optional[float] = Query(None, description="Maksimum toplam cycle süresi"),
    min_total_delay: Optional[float] = Query(None, description="Minimum toplam gecikme"),
    max_total_delay: Optional[float] = Query(None, description="Maksimum toplam gecikme"),
    start_date: Optional[str] = Query(None, description="Başlangıç tarihi (YYYY-MM-DD)"),
    end_date: Optional[str] = Query(None, description="Bitiş tarihi (YYYY-MM-DD)"),
    start_datetime: Optional[str] = Query(None, description="Başlangıç tarih ve saati"),
    end_datetime: Optional[str] = Query(None, description="Bitiş tarih ve saati"),
    sort_by: str = Query("production_time", description="Sıralama: production_time veya serial"),
    include_reverted: bool = Query(False, description="Geri alınmış ürünleri göster"),
    db: Session = Depends(get_db)
):
    """
    Ürün loglarını getirir, sipariş ve tarih filtresi uygulanabilir.
    Order, Product ve TaskLog tablolarını birleştirerek kapsamlı ürün raporu sunar.
    """
    # Product tablosundan başla, Order ile join yap
    query = db.query(
        Product.id.label("product_id"),
        Product.name.label("product_name"),
        Product.serialnumber,
        Product.productsize,
        Product.producttype,
        Product.moldnumber,
        Product.lot_number.label("product_lot_number"),
        Product.timestamp.label("product_timestamp"),
        Product.ordernumber,
        Order.name.label("order_name"),
        Order.quantity.label("order_quantity"),
        Order.lot_number.label("order_lot_number"),
        Order.serial_start,
        Person.name.label("person_name"),
        Person.surname.label("person_surname")
    ).outerjoin(Order, Product.ordernumber == Order.ordernumber)\
     .outerjoin(Person, Order.person_id == Person.id)
    if not include_reverted:
        query = query.filter(Product.is_reverted.is_(False))
    
    # Filtreleme
    if ordernumber is not None:
        query = query.filter(Product.ordernumber == ordernumber)
    if serialnumber is not None:
        query = query.filter(Product.serialnumber == serialnumber)
    if product_name:
        query = query.filter(Product.name.ilike(f"%{product_name}%"))
    if productsize:
        query = query.filter(Product.productsize.ilike(f"%{productsize}%"))
    if producttype:
        query = query.filter(Product.producttype.ilike(f"%{producttype}%"))
    if moldnumber:
        query = query.filter(Product.moldnumber.ilike(f"%{moldnumber}%"))
    if lot_number:
        query = query.filter(or_(
            Product.lot_number.ilike(f"%{lot_number}%"),
            Product.lot_number.is_(None) & Order.lot_number.ilike(f"%{lot_number}%")
        ))
    if machine_id is not None:
        machine_products = db.query(TaskLog.product_id).filter(TaskLog.machine_id == machine_id)
        query = query.filter(Product.id.in_(machine_products))
    if person_id is not None:
        person_products = db.query(TaskLog.product_id).filter(TaskLog.person_id == person_id)
        query = query.filter(Product.id.in_(person_products))
    if status:
        status_products = db.query(TaskLog.product_id).filter(TaskLog.status == status)
        query = query.filter(Product.id.in_(status_products))
    if cycle_id is not None:
        cycle_products = db.query(TaskLog.product_id).filter(TaskLog.cycle_id == cycle_id)
        query = query.filter(Product.id.in_(cycle_products))
    if min_total_duration is not None or max_total_duration is not None:
        duration_products = db.query(TaskLog.product_id).join(CycleLog, TaskLog.cycle_id == CycleLog.id)
        if min_total_duration is not None:
            duration_products = duration_products.filter(CycleLog.total_duration >= min_total_duration)
        if max_total_duration is not None:
            duration_products = duration_products.filter(CycleLog.total_duration <= max_total_duration)
        query = query.filter(Product.id.in_(duration_products))
    if min_total_delay is not None or max_total_delay is not None:
        delayed_products = db.query(TaskLog.product_id).join(CycleLog, TaskLog.cycle_id == CycleLog.id)
        if min_total_delay is not None:
            delayed_products = delayed_products.filter(CycleLog.total_delay >= min_total_delay)
        if max_total_delay is not None:
            delayed_products = delayed_products.filter(CycleLog.total_delay <= max_total_delay)
        query = query.filter(Product.id.in_(delayed_products))
    cycle_time_start = cycle_time_end = None
    if start_datetime:
        try:
            cycle_time_start = datetime.fromisoformat(start_datetime)
        except ValueError:
            raise HTTPException(status_code=400, detail="Geçersiz başlangıç tarih-saat değeri")
    if end_datetime:
        try:
            cycle_time_end = datetime.fromisoformat(end_datetime)
        except ValueError:
            raise HTTPException(status_code=400, detail="Geçersiz bitiş tarih-saat değeri")
    if cycle_time_start or cycle_time_end:
        cycle_query = db.query(TaskLog.product_id).join(CycleLog, TaskLog.cycle_id == CycleLog.id)
        if cycle_time_start:
            cycle_query = cycle_query.filter(CycleLog.cycle_start_time >= cycle_time_start)
        if cycle_time_end:
            cycle_query = cycle_query.filter(CycleLog.cycle_end_time <= cycle_time_end)
        query = query.filter(Product.id.in_(cycle_query.subquery()))
    if start_date:
        query = query.filter(Product.timestamp >= start_date)
    if end_date:
        query = query.filter(Product.timestamp <= end_date)
    
    products = query.order_by(Product.timestamp.desc()).all()
    
    # Her ürün için TaskLog bilgilerini al
    result = []
    for p in products:
        # İlgili task logları getir
        task_query = db.query(TaskLog).filter(TaskLog.product_id == p.product_id)
        if machine_id is not None:
            task_query = task_query.filter(TaskLog.machine_id == machine_id)
        task_logs = task_query.order_by(TaskLog.timestamp.asc(), TaskLog.id.asc()).all()
        
        task_info = []
        task_person_name = None
        for task in task_logs:
            # Machine bilgisi
            machine = db.query(Machine).filter(Machine.id == task.machine_id).first()
            machine_name = machine.name if machine else f"Makine {task.machine_id}"
            task_person = db.query(Person).filter(Person.id == task.person_id).first()
            current_person_name = f"{task_person.name} {task_person.surname}" if task_person else None
            if task_person_name is None and current_person_name:
                task_person_name = current_person_name
            
            # Cycle bilgisi
            cycle_info = None
            if task.cycle_id:
                cycle = db.query(CycleLog).filter(CycleLog.id == task.cycle_id).first()
                if cycle:
                    cycle_info = {
                        "id": cycle.id,
                        "start": cycle.cycle_start_time.isoformat() if cycle.cycle_start_time else None,
                        "end": cycle.cycle_end_time.isoformat() if cycle.cycle_end_time else None,
                        "cooking_duration": cycle.cooking_duration,
                        "renewal_duration": cycle.renewal_duration,
                        "total_duration": cycle.total_duration,
                        "ideal_cooking_time": cycle.ideal_cooking_time,
                        "ideal_renewal_time": cycle.ideal_renewal_time,
                        "ideal_total_time": cycle.ideal_total_time,
                        "cooking_delay": cycle.cooking_delay,
                        "renewal_delay": cycle.renewal_delay,
                        "total_delay": cycle.total_delay
                    }
            
            task_info.append({
                "task_id": task.id,
                "machine_name": machine_name,
                "machine_id": task.machine_id,
                "task_name": task.task_name,
                "status": task.status,
                "timestamp": task.timestamp.isoformat() if task.timestamp else None,
                "details": task.details,
                "person_id": task.person_id,
                "person_name": current_person_name,
                "cycle": cycle_info
            })
        
        result.append({
            "product_id": p.product_id,
            "product_name": p.product_name,
            "serial_number": p.serialnumber,
            "product_size": p.productsize,
            "product_type": p.producttype,
            "mold_number": p.moldnumber,
            "timestamp": p.product_timestamp.isoformat() if p.product_timestamp else None,
            "order_number": p.ordernumber,
            "order_name": p.order_name,
            "order_quantity": p.order_quantity,
            "lot_number": p.product_lot_number if p.product_lot_number is not None else p.order_lot_number,
            "serial_start": p.serial_start,
            "current_status": task_logs[-1].status if task_logs else None,
            # Ürüne özel seçilen TaskLog personeli, sipariş personelinden önceliklidir.
            "person_name": task_person_name or (f"{p.person_name} {p.person_surname}" if p.person_name else None),
            "tasks": task_info,
            "cycle_end_timestamp": next(
                (t["cycle"]["end"] for t in task_info if t.get("cycle") and t["cycle"].get("end")),
                None
            ),
            "production_timestamp": next(
                (t["cycle"]["end"] for t in task_info if t.get("cycle") and t["cycle"].get("end")),
                p.product_timestamp.isoformat() if p.product_timestamp else None
            )
        })
    if sort_by == "serial":
        result.sort(key=lambda item: item.get("serial_number") or "")
    else:
        result.sort(key=lambda item: item.get("production_timestamp") or "", reverse=True)
    return result

@app.get("/sensor-logs")
def get_sensor_logs(
    machine_id: int = Query(None, description="Makine ID"),
    sensor_id: int = Query(None, description="Sensör ID"),
    ordernumber: Optional[str] = Query(None, description="Sipariş numarası"),
    start_date: Optional[str] = Query(None, description="Başlangıç tarihi (YYYY-MM-DD)"),
    end_date: Optional[str] = Query(None, description="Bitiş tarihi (YYYY-MM-DD)"),
    min_value: int = Query(None, description="Minimum değer"),
    max_value: int = Query(None, description="Maximum değer"),
    db: Session = Depends(get_db)
):
    """
    Sensör loglarını getirir, makine, sensör ve tarih filtresi uygulanabilir.
    """
    query = db.query(SensorLog)
    if ordernumber:
        cycle_rows = db.query(TaskLog.cycle_id).join(Product, TaskLog.product_id == Product.id).filter(
            Product.ordernumber == ordernumber, TaskLog.cycle_id.isnot(None)).distinct().all()
        cycles = db.query(CycleLog).filter(CycleLog.id.in_([row[0] for row in cycle_rows])).all()
        conditions = []
        for cycle in cycles:
            if cycle.cycle_start_time and cycle.cycle_end_time:
                conditions.append((SensorLog.machine_id == cycle.machine_id) &
                    (SensorLog.timestamp >= cycle.cycle_start_time) & (SensorLog.timestamp <= cycle.cycle_end_time))
        query = query.filter(or_(*conditions)) if conditions else query.filter(text("1=0"))
    
    if machine_id is not None:
        query = query.filter(SensorLog.machine_id == machine_id)
    if sensor_id is not None:
        query = query.filter(SensorLog.sensor_id == sensor_id)
    if start_date:
        try:
            query = query.filter(SensorLog.timestamp >= datetime.fromisoformat(start_date))
        except ValueError:
            raise HTTPException(status_code=400, detail="Geçersiz başlangıç tarihi")
    if end_date:
        try:
            end_value = datetime.fromisoformat(end_date)
            if len(end_date) == 10:
                end_value = end_value.replace(hour=23, minute=59, second=59, microsecond=999999)
            query = query.filter(SensorLog.timestamp <= end_value)
        except ValueError:
            raise HTTPException(status_code=400, detail="Geçersiz bitiş tarihi")
    if min_value is not None:
        query = query.filter(SensorLog.value >= min_value)
    if max_value is not None:
        query = query.filter(SensorLog.value <= max_value)
    
    logs = query.order_by(SensorLog.timestamp.desc()).all()
    
    result = []
    for log in logs:
        # Makine bilgisi
        machine = db.query(Machine).filter(Machine.id == log.machine_id).first()
        machine_name = machine.name if machine else f"Makine {log.machine_id}"
        
        result.append({
            "id": log.id,
            "machine_id": log.machine_id,
            "machine_name": machine_name,
            "sensor_id": log.sensor_id,
            "timestamp": log.timestamp.isoformat() if log.timestamp else None,
            "value": log.value,
            "value_set": log.value_set,
            "value_histerezis": log.value_histerezis,
            "deviation": log.value - log.value_set  # Sapma
        })
    
    return result

@app.get("/task-logs")
def get_task_logs(
    machine_id: int = Query(None, description="Makine ID"),
    person_id: int = Query(None, description="Personel ID"),
    product_id: int = Query(None, description="Ürün ID"),
    cycle_id: int = Query(None, description="CycleLog ID"),
    status: str = Query(None, description="Durum"),
    start: Optional[str] = Query(None, description="Başlangıç tarihi (YYYY-MM-DD)"),
    end: Optional[str] = Query(None, description="Bitiş tarihi (YYYY-MM-DD)"),
    db: Session = Depends(get_db)
):
    """
    TaskLog tablosundan verileri getirir, filtre uygulanabilir.
    """
    query = db.query(TaskLog)
    if machine_id is not None:
        query = query.filter(TaskLog.machine_id == machine_id)
    if person_id is not None:
        query = query.filter(TaskLog.person_id == person_id)
    if product_id is not None:
        query = query.filter(TaskLog.product_id == product_id)
    if cycle_id is not None and cycle_id != "":
        query = query.filter(TaskLog.cycle_id == cycle_id)
    if status is not None:
        query = query.filter(TaskLog.status == status)
    if start:
        query = query.filter(TaskLog.timestamp >= start)
    if end:
        query = query.filter(TaskLog.timestamp <= end)
    logs = query.order_by(TaskLog.timestamp.asc()).all()
    enriched_logs = []
    for log in logs:
        cycle = db.query(CycleLog).filter(CycleLog.id == log.cycle_id).first() if log.cycle_id else None
        delay = getattr(cycle, "delay", None) if cycle else None
        # TimerLog'dan T1 ve T2 değerlerini al
        t1 = None
        t2 = None
        if cycle:
            # Build TimerLog query filters safely: only compare end time when it's not None
            tl_query = db.query(TimerLog).filter(
                TimerLog.machine_id == log.machine_id,
                TimerLog.timestamp >= cycle.cycle_start_time,
            )
            if cycle.cycle_end_time is not None:
                tl_query = tl_query.filter(TimerLog.timestamp <= cycle.cycle_end_time)
            timer_logs = tl_query.all()
            if timer_logs:
                t1_vals = [tl.value_T1 for tl in timer_logs if tl.value_T1 is not None]
                t2_vals = [tl.value_T2 for tl in timer_logs if tl.value_T2 is not None]
                t1 = round(sum(t1_vals) / len(t1_vals), 2) if t1_vals else None
                t2 = round(sum(t2_vals) / len(t2_vals), 2) if t2_vals else None
        enriched_logs.append({
            "id": log.id,
            "machine_id": log.machine_id,
            "person_id": log.person_id,
            "product_id": log.product_id,
            "cycle_id": log.cycle_id,
            "timestamp": log.timestamp.isoformat(),
            "task_name": log.task_name,
            "status": log.status,
            "details": log.details,
            "delay": delay,
            "t1": t1,
            "t2": t2
            , "product_reverted": bool(getattr(db.query(Product).filter(Product.id == log.product_id).first(), "is_reverted", False))
        })
    return enriched_logs

@app.get("/task-logs/export")
def export_task_logs(
    machine_id: int = Query(None, description="Makine ID"),
    person_id: int = Query(None, description="Personel ID"),
    product_id: int = Query(None, description="Ürün ID"),
    cycle_id: int = Query(None, description="CycleLog ID"),
    status: str = Query(None, description="Durum"),
    start: Optional[str] = Query(None, description="Başlangıç tarihi (YYYY-MM-DD)"),
    end: Optional[str] = Query(None, description="Bitiş tarihi (YYYY-MM-DD)"),
    db: Session = Depends(get_db)
):
    """
    TaskLog tablosundan verileri Excel olarak export eder.
    """
    query = db.query(TaskLog)
    if machine_id is not None:
        query = query.filter(TaskLog.machine_id == machine_id)
    if person_id is not None:
        query = query.filter(TaskLog.person_id == person_id)
    if product_id is not None:
        query = query.filter(TaskLog.product_id == product_id)
    if cycle_id is not None:
        query = query.filter(TaskLog.cycle_id == cycle_id)
    if status is not None:
        query = query.filter(TaskLog.status == status)
    if start:
        query = query.filter(TaskLog.timestamp >= start)
    if end:
        query = query.filter(TaskLog.timestamp <= end)
    logs = query.order_by(TaskLog.timestamp.asc()).all()
    data = [
        {
            "ID": log.id,
            "Makine ID": log.machine_id,
            "Personel ID": log.person_id,
            "Ürün ID": log.product_id,
            "CycleLog ID": log.cycle_id,
            "Tarih": log.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
            "Görev Adı": log.task_name,
            "Durum": log.status,
            "Açıklama": log.details
        }
        for log in logs
    ]
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="TaskLogs")
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=task_logs.xlsx"}
    )

@app.post("/products")
def create_product(product: dict = Body(...), db: Session = Depends(get_db)):
    """
    Yeni ürün kaydı oluşturur.
    """
    order = db.query(Order).filter(Order.ordernumber == product.get("ordernumber")).first() if product.get("ordernumber") else None
    serial = str(product.get("serialnumber") or "").strip()
    if not serial:
        raise HTTPException(status_code=400, detail="Tam seri numarası zorunludur")
    if order:
        if order.is_manually_completed:
            raise HTTPException(status_code=409, detail="Tamamlanmış siparişe ürün eklenemez")
        existing_serials = db.query(Product.serialnumber).filter(Product.ordernumber == order.ordernumber).all()
        used = []
        for (serial_value,) in existing_serials:
            try:
                used.append(int(str(serial_value).rsplit('/', 1)[1]))
            except (ValueError, IndexError):
                pass
        next_sequence = max(used, default=(order.serial_start or 1) - 1) + 1
        expected = f"{order.ordernumber}/{next_sequence}"
        if serial != expected:
            raise HTTPException(status_code=409, detail=f"Sıradaki seri numarası {expected} olmalıdır")
    prod = Product(
        name=product.get("name"),
        serialnumber=serial,
        productsize=product.get("productsize"),
        producttype=product.get("producttype"),
        moldnumber=product.get("moldnumber"),
        lot_number=(product.get("lot_number") if product.get("lot_number") is not None else (order.lot_number if order else None)),
        ordernumber=product.get("ordernumber"),
        timestamp=datetime.now()
    )
    db.add(prod)
    db.flush()
    if product.get("machine_id") is not None:
        if not product.get("person_id"):
            db.rollback()
            raise HTTPException(status_code=400, detail="Personel zorunludur")
        db.add(TaskLog(machine_id=int(product["machine_id"]), person_id=int(product["person_id"]),
            product_id=prod.id, cycle_id=product.get("cycle_id"), timestamp=datetime.now(),
            task_name=str(product.get("task_name") or f"Ürün - {prod.name}"),
            status=product.get("status") or "SUCCESS", details=product.get("details")))
    db.commit()
    db.refresh(prod)
    return {"id": prod.id}

@app.post("/products/manual")
def create_manual_products(products: List[dict] = Body(...), db: Session = Depends(get_db)):
    """Manuel girilen ürünleri ve görev kayıtlarını tek transaction'da oluşturur."""
    if not products:
        raise HTTPException(status_code=400, detail="En az bir ürün girilmelidir")
    if len(products) > 100:
        raise HTTPException(status_code=400, detail="Tek seferde en fazla 100 ürün girilebilir")

    allowed_statuses = {"SUCCESS", "FAILED", "CONDITIONAL"}
    payload_serials = []

    try:
        for index, item in enumerate(products, start=1):
            required = ("name", "serialnumber", "productsize", "producttype", "moldnumber", "machine_id", "person_id")
            missing = [key for key in required if item.get(key) in (None, "")]
            if missing:
                raise ValueError(f"Ürün {index}: eksik alanlar: {', '.join(missing)}")

            serialnumber = str(item["serialnumber"]).strip()
            machine_id = int(item["machine_id"])
            person_id = int(item["person_id"])
            ordernumber = str(item["ordernumber"]).strip() if item.get("ordernumber") not in (None, "") else None
            status = item.get("status", "SUCCESS")

            if status not in allowed_statuses:
                raise ValueError(f"Ürün {index}: geçersiz durum")
            if not db.query(Machine.id).filter(Machine.id == machine_id).first():
                raise ValueError(f"Ürün {index}: makine bulunamadı")
            if not db.query(Person.id).filter(Person.id == person_id).first():
                raise ValueError(f"Ürün {index}: personel bulunamadı")
            if ordernumber is not None and not db.query(Order.id).filter(Order.ordernumber == ordernumber).first():
                raise ValueError(f"Ürün {index}: sipariş bulunamadı")
            payload_serials.append(serialnumber)

        if len(payload_serials) != len(set(payload_serials)):
            raise ValueError("Aynı seri numarası birden fazla kartta kullanılamaz")
        existing = db.query(Product.serialnumber).filter(Product.serialnumber.in_(payload_serials)).first()
        if existing:
            raise ValueError(f"Seri numarası {existing[0]} zaten kayıtlı")

        created_ids = []
        for item in products:
            prod = Product(
                name=str(item["name"]).strip(),
                serialnumber=str(item["serialnumber"]).strip(),
                productsize=str(item["productsize"]).strip(),
                producttype=str(item["producttype"]).strip(),
                moldnumber=str(item["moldnumber"]).strip(),
                lot_number=(str(item.get("lot_number", "")).strip()),
                ordernumber=str(item["ordernumber"]).strip() if item.get("ordernumber") not in (None, "") else None,
                timestamp=datetime.now()
            )
            db.add(prod)
            db.flush()

            log = TaskLog(
                machine_id=int(item["machine_id"]),
                person_id=int(item["person_id"]),
                product_id=prod.id,
                # Manuel ürünler fiziksel bir cycle'dan gelmediği için cycle aranmaz.
                cycle_id=None,
                timestamp=datetime.now(),
                task_name=(str(item.get("task_name") or "Manuel ürün girişi").strip()),
                status=item.get("status", "SUCCESS"),
                details=(str(item["details"]).strip() if item.get("details") else None)
            )
            db.add(log)
            created_ids.append(prod.id)

        db.commit()
        return {"message": f"{len(created_ids)} ürün kaydedildi", "product_ids": created_ids}
    except (ValueError, TypeError) as e:
        db.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    except SQLAlchemyError as e:
        db.rollback()
        raise HTTPException(status_code=400, detail=f"Veritabanı hatası: {str(e)}")

@app.patch("/products/{product_id}/lot-number")
def update_product_lot_number(product_id: int, payload: dict = Body(...), db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id).first()
    if not product:
        raise HTTPException(status_code=404, detail="Ürün bulunamadı")
    lot_number = payload.get("lot_number")
    # Boş metin, siparişten gelen lotu ürün bazında bilinçli olarak temizler.
    product.lot_number = str(lot_number).strip() if lot_number is not None else ""
    db.commit()
    return {"message": "Lot numarası güncellendi", "lot_number": product.lot_number}

@app.patch("/products/{product_id}")
def update_product(product_id: int, payload: dict = Body(...), db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id, Product.is_reverted.is_(False)).first()
    if not product:
        raise HTTPException(status_code=404, detail="Aktif ürün bulunamadı")
    latest = db.query(TaskLog).filter(TaskLog.product_id == product_id).order_by(TaskLog.timestamp.desc(), TaskLog.id.desc()).first()
    for key in ("name", "productsize", "producttype", "moldnumber", "lot_number", "ordernumber"):
        if key in payload:
            setattr(product, key, str(payload[key]).strip() if payload[key] is not None else None)
    if latest and payload.get("person_id") is not None:
        if not db.query(Person.id).filter(Person.id == int(payload["person_id"])).first():
            raise HTTPException(status_code=400, detail="Personel bulunamadı")
        latest.person_id = int(payload["person_id"])
    db.commit()
    return {"message": "Ürün bilgileri güncellendi"}

@app.post("/products/{product_id}/revert-to-cycle")
def revert_product_to_cycle(product_id: int, payload: dict = Body(...), db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id, Product.is_reverted.is_(False)).first()
    latest = db.query(TaskLog).filter(TaskLog.product_id == product_id).order_by(TaskLog.timestamp.desc(), TaskLog.id.desc()).first()
    if not product or not latest or not latest.cycle_id:
        raise HTTPException(status_code=404, detail="Cycle bağlantılı aktif ürün bulunamadı")
    reason = str(payload.get("reason") or "").strip()
    if not reason:
        raise HTTPException(status_code=400, detail="Geri alma nedeni zorunludur")
    product.is_reverted, product.reverted_at, product.revert_reason = True, datetime.now(), reason
    db.add(TaskLog(machine_id=latest.machine_id, person_id=latest.person_id, product_id=product.id,
        cycle_id=latest.cycle_id, timestamp=datetime.now(), task_name="Üründen cycle'a geri alma",
        status="REVERTED", details=reason))
    order = db.query(Order).filter(Order.ordernumber == product.ordernumber).first() if product.ordernumber else None
    if order and order.is_manually_completed:
        order.is_manually_completed, order.completed_at = False, None
        order.completion_note = "Ürün geri alındığı için sipariş yeniden açıldı."
    db.commit()
    return {"message": "Ürün silinmeden geri alındı; cycle yeniden kullanılabilir", "cycle_id": latest.cycle_id}

@app.post("/task-logs")
def create_task_log(task: dict = Body(...), db: Session = Depends(get_db)):
    """
    Yeni görev kaydı oluşturur.
    """
    log = TaskLog(
        machine_id=task.get("machine_id"),
        person_id=task.get("person_id"),
        product_id=task.get("product_id"),
        cycle_id=task.get("cycle_id"),
        timestamp=datetime.now(),
        task_name=task.get("task_name"),
        status=task.get("status"),
        details=task.get("details")
    )
    db.add(log)
    db.commit()
    db.refresh(log)
    return {"id": log.id}

def _order_counts_many(db: Session, ordernumber=None):
    # Son durum: önce zaman, aynı zamanda büyük ID. Geri alınmış ürünleri sayma.
    latest_query = db.query(
        TaskLog.product_id.label("product_id"), TaskLog.status.label("status"),
        func.row_number().over(partition_by=TaskLog.product_id,
            order_by=(TaskLog.timestamp.desc(), TaskLog.id.desc())).label("position")
    )
    if ordernumber is not None:
        latest_query = latest_query.join(Product, TaskLog.product_id == Product.id).filter(
            Product.ordernumber == ordernumber, Product.is_reverted.is_(False))
    latest = latest_query.subquery()
    query = db.query(Product.ordernumber, latest.c.status, func.count(Product.id)).outerjoin(
        latest, and_(latest.c.product_id == Product.id, latest.c.position == 1)
    ).filter(Product.is_reverted.is_(False))
    if ordernumber is not None:
        query = query.filter(Product.ordernumber == ordernumber)
    result = {}
    for number, status, count in query.group_by(Product.ordernumber, latest.c.status).all():
        produced, counts = result.setdefault(number, [0, {"SUCCESS": 0, "FAILED": 0, "CONDITIONAL": 0}])
        result[number][0] += count
        if status in counts:
            counts[status] += count
    return result


def _order_counts(db: Session, ordernumber: str):
    produced, counts = _order_counts_many(db, ordernumber).get(
        ordernumber, (0, {"SUCCESS": 0, "FAILED": 0, "CONDITIONAL": 0}))
    return produced, counts

@app.get("/order-status/{ordernumber}")
def get_order_status(ordernumber: str, db: Session = Depends(get_db)):
    """
    Sipariş durumu ve üretim sayısını getirir.
    """
    order = db.query(Order).filter(Order.ordernumber == ordernumber).first()
    if not order:
        return {"error": "Sipariş bulunamadı"}
    
    # Bu siparişten kaç ürün üretilmiş
    produced_count, counts = _order_counts(db, ordernumber)
    all_serials = db.query(Product.serialnumber).filter(Product.ordernumber == ordernumber).all()
    used_sequences = []
    for (serial_value,) in all_serials:
        try:
            used_sequences.append(int(str(serial_value).rsplit('/', 1)[1]))
        except (ValueError, IndexError):
            pass
    next_sequence = max(used_sequences, default=(order.serial_start or 1) - 1) + 1
    next_serial = f"{order.ordernumber}/{next_sequence}"
    accepted_count = counts["SUCCESS"]
    
    return {
        "ordernumber": order.ordernumber,
        "name": order.name,
        "quantity": order.quantity,
        "produced_count": produced_count,
        "accepted_count": accepted_count,
        "failed_count": counts["FAILED"],
        "conditional_count": counts["CONDITIONAL"],
        "remaining": max(0, order.quantity - accepted_count),
        "next_serial": next_serial,
        "next_sequence": next_sequence,
        "is_completed": bool(order.is_manually_completed),
        "can_complete": accepted_count >= order.quantity,
        "total_volume": produced_count * (order.volume_per_piece or 0),
        "accepted_volume": accepted_count * (order.volume_per_piece or 0)
    }

@app.post("/orders/{ordernumber}/complete")
def complete_order(ordernumber: str, payload: dict = Body(default={}), db: Session = Depends(get_db)):
    order = db.query(Order).filter(Order.ordernumber == ordernumber).first()
    if not order:
        raise HTTPException(status_code=404, detail="Sipariş bulunamadı")
    _, counts = _order_counts(db, ordernumber)
    if counts["SUCCESS"] < order.quantity and not payload.get("force"):
        raise HTTPException(status_code=409, detail=f"Sağlam ürün sayısı yetersiz ({counts['SUCCESS']}/{order.quantity}). Yetkili eksik kapatma onayı gerekir.")
    note = str(payload.get("note") or "").strip()
    if counts["SUCCESS"] < order.quantity and not note:
        raise HTTPException(status_code=400, detail="Eksik kapatma için gerekçe zorunludur")
    order.is_manually_completed = True
    order.completed_at = datetime.now()
    order.completion_note = note or None
    db.commit()
    return {"message": "Sipariş tamamlandı"}

@app.post("/orders/{ordernumber}/reopen")
def reopen_order(ordernumber: str, payload: dict = Body(default={}), db: Session = Depends(get_db)):
    order = db.query(Order).filter(Order.ordernumber == ordernumber).first()
    if not order:
        raise HTTPException(status_code=404, detail="Sipariş bulunamadı")
    if not order.is_manually_completed:
        return {"message": "Sipariş zaten açık", "reopened": False}
    reason = str(payload.get("reason") or "Yeni ürün eklemek için yeniden açıldı").strip()
    order.is_manually_completed = False
    order.completed_at = None
    order.completion_note = reason
    db.commit()
    return {"message": "Sipariş yeniden açıldı", "reopened": True}

@app.patch("/products/{product_id}/status")
def update_product_status(product_id: int, payload: dict = Body(...), db: Session = Depends(get_db)):
    status = payload.get("status")
    if status not in {"SUCCESS", "FAILED", "CONDITIONAL"}:
        raise HTTPException(status_code=400, detail="Geçersiz ürün durumu")
    product = db.query(Product).filter(Product.id == product_id).first()
    latest = db.query(TaskLog).filter(TaskLog.product_id == product_id).order_by(TaskLog.timestamp.desc(), TaskLog.id.desc()).first()
    if not product or not latest:
        raise HTTPException(status_code=404, detail="Ürün veya işlem kaydı bulunamadı")
    old_status = latest.status
    reason = str(payload.get("reason") or "").strip()
    status_names = {"SUCCESS": "Başarılı", "FAILED": "Başarısız", "CONDITIONAL": "Şartlı"}
    product.status_change = f"{status_names.get(old_status, old_status)} → {status_names.get(status, status)}"
    product.status_change_reason = reason or "Sebep belirtilmedi"
    product.status_changed_at = datetime.now()
    db.add(TaskLog(machine_id=latest.machine_id, person_id=latest.person_id, product_id=product_id,
        cycle_id=latest.cycle_id, timestamp=datetime.now(), task_name="Kalite durumu değişikliği",
        status=status, details=reason or "Sebep belirtilmedi"))
    order = db.query(Order).filter(Order.ordernumber == product.ordernumber).first() if product.ordernumber else None
    order_reopened = False
    if order and order.is_manually_completed and old_status == "SUCCESS" and status != "SUCCESS":
        order.is_manually_completed = False
        order.completed_at = None
        order.completion_note = "Ürün durumu değiştiği için sipariş otomatik yeniden açıldı."
        order_reopened = True
    db.commit()
    return {"message": "Ürün durumu güncellendi", "order_reopened": order_reopened}

@app.get("/orders")
def get_orders(db: Session = Depends(get_db)):
    """
    Tüm siparişleri getirir.
    """
    orders = db.query(Order).options(joinedload(Order.person)).order_by(Order.timestamp.desc()).all()
    all_counts = _order_counts_many(db)
    result = []
    for o in orders:
        produced_count, counts = all_counts.get(o.ordernumber, (0, {"SUCCESS": 0, "FAILED": 0, "CONDITIONAL": 0}))
        result.append({
            "id": o.id,
            "name": o.name,
            "productsize": o.productsize,
            "producttype": o.producttype,
            "moldnumber": o.moldnumber,
            "ordernumber": o.ordernumber,
            "quantity": o.quantity,
            "unit": o.unit,
            "volume_per_piece": o.volume_per_piece,
            "produced_count": produced_count,
            "accepted_count": counts["SUCCESS"],
            "failed_count": counts["FAILED"],
            "conditional_count": counts["CONDITIONAL"],
            "remaining": max(0, o.quantity - counts["SUCCESS"]),
            "is_completed": bool(o.is_manually_completed),
            "serial_start": o.serial_start,
            "lot_number": o.lot_number,
            "person_id": o.person_id,
            "person_name": f"{o.person.name} {o.person.surname}" if o.person else None,
            "timestamp": o.timestamp.isoformat() if o.timestamp else None
        })
    return result

@app.get("/products/{product_id}/sensor-deviations")
def get_product_sensor_deviations(product_id: int, db: Session = Depends(get_db)):
    """Ürünün pişirme aralığında yalnızca set ± histerezis dışındaki sensör olaylarını döndürür."""
    task = db.query(TaskLog).filter(TaskLog.product_id == product_id, TaskLog.cycle_id.isnot(None)).order_by(TaskLog.timestamp.desc()).first()
    if not task:
        return {"events": [], "message": "Ürüne bağlı cycle bulunamadı"}
    cycle = db.query(CycleLog).filter(CycleLog.id == task.cycle_id).first()
    if not cycle or not cycle.cycle_start_time or not cycle.cycle_end_time:
        return {"events": [], "message": "Cycle zaman aralığı tamamlanmamış"}
    start = cycle.cooking_start_time or cycle.cycle_start_time
    logs = db.query(SensorLog).filter(SensorLog.machine_id == task.machine_id,
        SensorLog.timestamp >= start, SensorLog.timestamp <= cycle.cycle_end_time).order_by(SensorLog.sensor_id, SensorLog.timestamp).all()
    outside = []
    for log in logs:
        low, high = log.value_set - abs(log.value_histerezis), log.value_set + abs(log.value_histerezis)
        if log.value < low or log.value > high:
            outside.append({"sensor_id": log.sensor_id, "timestamp": log.timestamp, "value": log.value,
                "low_limit": low, "high_limit": high, "kind": "LOW" if log.value < low else "HIGH"})
    events = []
    for item in outside:
        previous = events[-1] if events else None
        # Aynı sensör ve yöndeki ardışık limit ihlallerini tek olay olarak grupla.
        if previous and previous["sensor_id"] == item["sensor_id"] and previous["kind"] == item["kind"] and (item["timestamp"] - previous["end_raw"]).total_seconds() <= 60:
            previous["end_raw"] = item["timestamp"]
            previous["end"] = item["timestamp"].isoformat()
            previous["duration_seconds"] = int((item["timestamp"] - previous["start_raw"]).total_seconds())
            previous["extreme_value"] = min(previous["extreme_value"], item["value"]) if item["kind"] == "LOW" else max(previous["extreme_value"], item["value"])
            previous["sample_count"] += 1
        else:
            events.append({"sensor_id": item["sensor_id"], "kind": item["kind"],
                "start_raw": item["timestamp"], "end_raw": item["timestamp"], "start": item["timestamp"].isoformat(),
                "end": item["timestamp"].isoformat(), "duration_seconds": 0, "extreme_value": item["value"],
                "low_limit": item["low_limit"], "high_limit": item["high_limit"], "sample_count": 1})
    for event in events:
        event.pop("start_raw", None); event.pop("end_raw", None)
    machine = db.query(Machine).filter(Machine.id == task.machine_id).first()
    return {"product_id": product_id, "cycle_id": cycle.id, "machine_name": machine.name if machine else str(task.machine_id),
        "cooking_start": start.isoformat(), "cycle_end": cycle.cycle_end_time.isoformat(), "events": events,
        "message": None if events else "Bu ürünün pişirme sürecinde sensör sapması bulunamadı"}

@app.post("/product-logs/export")
def export_product_logs(payload: dict = Body(...), db: Session = Depends(get_db)):
    """Filtrelenmiş ID'leri veya filtresiz durumda bütün ürünleri XLSX olarak döndürür."""
    filtered = bool(payload.get("filtered"))
    product_ids = payload.get("product_ids") or []
    query = db.query(Product).order_by(Product.timestamp.desc())
    query = query.filter(Product.is_reverted.is_(False))
    if filtered:
        query = query.filter(Product.id.in_(product_ids)) if product_ids else query.filter(text("1=0"))

    rows = []
    for product in query.all():
        task_logs = db.query(TaskLog).filter(TaskLog.product_id == product.id)\
            .order_by(TaskLog.timestamp.desc(), TaskLog.id.desc()).all()
        task = task_logs[0] if task_logs else None
        status_names = {"SUCCESS": "Başarılı", "FAILED": "Başarısız", "CONDITIONAL": "Şartlı"}
        order = db.query(Order).filter(Order.ordernumber == product.ordernumber).first() if product.ordernumber else None
        person = db.query(Person).filter(Person.id == task.person_id).first() if task else None
        machine = db.query(Machine).filter(Machine.id == task.machine_id).first() if task else None
        cycle = db.query(CycleLog).filter(CycleLog.id == task.cycle_id).first() if task and task.cycle_id else None
        rows.append({
            "ID": product.id, "Seri No": product.serialnumber, "Ürün Adı": product.name,
            "Ebat": product.productsize, "Tip": product.producttype, "Kalıp No": product.moldnumber,
            "Sipariş No": product.ordernumber, "Sipariş Adı": order.name if order else None,
            "Lot No": product.lot_number if product.lot_number is not None else (order.lot_number if order else None),
            "Makine": machine.name if machine else None,
            "Personel": f"{person.name} {person.surname}" if person else None,
            "Durum": status_names.get(task.status, task.status) if task else None,
            "Yapılan Durum Değişikliği": product.status_change,
            "Değişiklik Sebebi": product.status_change_reason,
            "Değişiklik Tarihi": product.status_changed_at,
            "Ürün Tarihi": product.timestamp,
            "Cycle ID": cycle.id if cycle else None, "Cycle Başlangıç": cycle.cycle_start_time if cycle else None,
            "Cycle Bitiş": cycle.cycle_end_time if cycle else None,
            "Pişirme (dk)": cycle.cooking_duration if cycle else None, "Yenileme (dk)": cycle.renewal_duration if cycle else None,
            "Toplam (dk)": cycle.total_duration if cycle else None, "İdeal Pişirme": cycle.ideal_cooking_time if cycle else None,
            "İdeal Yenileme": cycle.ideal_renewal_time if cycle else None, "İdeal Toplam": cycle.ideal_total_time if cycle else None,
            "Pişirme Gecikme": cycle.cooking_delay if cycle else None, "Yenileme Gecikme": cycle.renewal_delay if cycle else None,
            "Toplam Gecikme": cycle.total_delay if cycle else None
        })
    output = create_simple_xlsx(rows, "Urun Loglari")
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=urun_loglari.xlsx"})

@app.post("/product-logs/delete")
def delete_product_logs(payload: dict = Body(...), db: Session = Depends(get_db)):
    product_ids = payload.get("product_ids") or []
    if not product_ids:
        raise HTTPException(status_code=400, detail="Silinecek ürün logu seçilmedi")
    try:
        products = db.query(Product).filter(Product.id.in_(product_ids), Product.is_reverted.is_(False)).all()
        reason = str(payload.get("reason") or "Rapor ekranından geri alındı").strip()
        for product in products:
            product.is_reverted, product.reverted_at, product.revert_reason = True, datetime.now(), reason
            latest = db.query(TaskLog).filter(TaskLog.product_id == product.id).order_by(TaskLog.id.desc()).first()
            if latest:
                db.add(TaskLog(machine_id=latest.machine_id, person_id=latest.person_id, product_id=product.id,
                    cycle_id=latest.cycle_id, timestamp=datetime.now(), task_name="Ürün kaydını güvenli geri alma",
                    status="REVERTED", details=reason))
        db.commit()
        return {"message": f"{len(products)} ürün silinmeden geri alındı", "deleted": len(products)}
    except SQLAlchemyError as e:
        db.rollback()
        raise HTTPException(status_code=400, detail=f"Silme işlemi başarısız: {str(e)}")

@app.post("/orders")
def create_order(order: dict = Body(...), db: Session = Depends(get_db)):
    """
    Yeni sipariş kaydı oluşturur.
    """
    ordernumber = str(order.get("ordernumber") or "").strip()
    if not re.fullmatch(r"\d+(?:-\d+)?", ordernumber):
        raise HTTPException(status_code=400, detail="Sipariş numarası 5315 veya 5312-1 biçiminde olmalıdır")
    if db.query(Order.id).filter(Order.ordernumber == ordernumber).first():
        raise HTTPException(status_code=409, detail="Bu sipariş/kalem numarası zaten kayıtlı")
    try:
        quantity = int(order.get("quantity"))
        volume_per_piece = float(order.get("volume_per_piece"))
        serial_start = int(order.get("serial_start") or 1)
        if quantity <= 0 or volume_per_piece <= 0 or serial_start <= 0:
            raise ValueError
    except (TypeError, ValueError):
        raise HTTPException(status_code=400, detail="Adet, dm³/adet ve seri başlangıcı sıfırdan büyük olmalıdır")
    new_order = Order(
        name=order.get("name"),
        productsize=order.get("productsize"),
        producttype=order.get("producttype"),
        moldnumber=order.get("moldnumber"),
        ordernumber=ordernumber,
        quantity=quantity,
        unit=(str(order.get("unit")).strip() if order.get("unit") else None),
        volume_per_piece=volume_per_piece,
        serial_start=serial_start,
        lot_number=order.get("lot_number"),
        person_id=order.get("person_id"),
        timestamp=datetime.now()
    )
    db.add(new_order)
    db.commit()
    db.refresh(new_order)
    return {"id": new_order.id}

@app.post("/orders/bulk")
def create_orders_bulk(orders: List[dict] = Body(...), db: Session = Depends(get_db)):
    """Aynı ana siparişin birbirinden bağımsız ürün kalemlerini tek transaction'da oluşturur."""
    if not orders or len(orders) > 100:
        raise HTTPException(status_code=400, detail="1 ile 100 arasında sipariş kalemi gönderilmelidir")
    numbers = [str(item.get("ordernumber") or "").strip() for item in orders]
    if len(numbers) != len(set(numbers)):
        raise HTTPException(status_code=400, detail="Aynı kalem numarası birden fazla kez kullanılamaz")
    if any(not re.fullmatch(r"\d+-[1-9]\d*", number) for number in numbers):
        raise HTTPException(status_code=400, detail="Kalem numaraları 254234-1, 254234-2 biçiminde olmalıdır")
    bases = {number.rsplit('-', 1)[0] for number in numbers}
    if len(bases) != 1:
        raise HTTPException(status_code=400, detail="Gönderilen kalemlerin tamamı aynı ana siparişe ait olmalıdır")
    base = next(iter(bases))
    existing_numbers = [row[0] for row in db.query(Order.ordernumber).filter(Order.ordernumber.like(f"{base}-%")).all()]
    existing_suffixes = []
    for number in existing_numbers:
        try:
            existing_suffixes.append(int(str(number).rsplit('-', 1)[1]))
        except (ValueError, IndexError):
            continue
    expected_start = max(existing_suffixes, default=0) + 1
    incoming_suffixes = sorted(int(number.rsplit('-', 1)[1]) for number in numbers)
    expected_suffixes = list(range(expected_start, expected_start + len(numbers)))
    if incoming_suffixes != expected_suffixes:
        expected_label = f"{base}-{expected_start}"
        raise HTTPException(status_code=400, detail=f"Yeni kalemler {expected_label} numarasından başlayarak sıralı olmalıdır")
    if db.query(Order.id).filter(Order.ordernumber.in_(numbers)).first():
        raise HTTPException(status_code=409, detail="Bu siparişin kalemlerinden en az biri zaten kayıtlı")
    created = []
    try:
        for item, number in zip(orders, numbers):
            quantity = int(item.get("quantity")); volume = float(item.get("volume_per_piece")); serial_start = int(item.get("serial_start") or 1)
            required = [item.get("name"), item.get("productsize"), item.get("producttype"), item.get("moldnumber")]
            if not all(str(field or "").strip() for field in required) or quantity <= 0 or volume <= 0 or serial_start <= 0:
                raise ValueError(f"{number} kaleminde eksik veya geçersiz bilgi var")
            row = Order(name=str(item["name"]).strip(), productsize=str(item["productsize"]).strip(),
                producttype=str(item["producttype"]).strip(), moldnumber=str(item["moldnumber"]).strip(),
                ordernumber=number, quantity=quantity, unit="dm³", volume_per_piece=volume,
                serial_start=serial_start, lot_number=str(item.get("lot_number") or "").strip() or None,
                person_id=item.get("person_id"), timestamp=datetime.now())
            db.add(row); created.append(row)
        db.commit()
        return {"message": f"{len(created)} sipariş kalemi kaydedildi", "ids": [row.id for row in created]}
    except (ValueError, TypeError) as exc:
        db.rollback()
        raise HTTPException(status_code=400, detail=str(exc))
    except SQLAlchemyError as exc:
        db.rollback()
        raise HTTPException(status_code=400, detail=f"Sipariş kalemleri kaydedilemedi: {exc}")
