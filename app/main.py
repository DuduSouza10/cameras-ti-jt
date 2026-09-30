from __future__ import annotations

import hashlib
import io
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from sqlalchemy import DateTime, ForeignKey, Integer, String, Text, create_engine, select
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column, relationship, sessionmaker

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
SEED_FILE = DATA_DIR / "seed.xlsx"

DATABASE_URL = os.getenv("DATABASE_URL", f"sqlite:///{DATA_DIR / 'dashboard.db'}")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+psycopg://", 1)
elif DATABASE_URL.startswith("postgresql://") and "+psycopg" not in DATABASE_URL:
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg://", 1)

connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine = create_engine(DATABASE_URL, pool_pre_ping=True, connect_args=connect_args)
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)


class Base(DeclarativeBase):
    pass


class Snapshot(Base):
    __tablename__ = "snapshots"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    captured_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc), index=True)
    source: Mapped[str] = mapped_column(String(40), default="excel")
    source_date: Mapped[str | None] = mapped_column(String(20), nullable=True)
    row_count: Mapped[int] = mapped_column(Integer, default=0)
    data_hash: Mapped[str] = mapped_column(String(64), index=True)
    records: Mapped[list["BaseRecord"]] = relationship(back_populates="snapshot", cascade="all, delete-orphan")


class BaseRecord(Base):
    __tablename__ = "base_records"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    snapshot_id: Mapped[int] = mapped_column(ForeignKey("snapshots.id", ondelete="CASCADE"), index=True)
    base: Mapped[str] = mapped_column(String(120), index=True)
    recording_days_raw: Mapped[str | None] = mapped_column(String(120), nullable=True)
    recording_days: Mapped[int | None] = mapped_column(Integer, nullable=True)
    camera_count_raw: Mapped[str | None] = mapped_column(String(120), nullable=True)
    camera_count: Mapped[int | None] = mapped_column(Integer, nullable=True)
    offline_raw: Mapped[str | None] = mapped_column(String(120), nullable=True)
    offline_count: Mapped[int | None] = mapped_column(Integer, nullable=True)
    status: Mapped[str] = mapped_column(String(120), index=True)
    is_implantation: Mapped[int] = mapped_column(Integer, default=0)
    quality_note: Mapped[str | None] = mapped_column(Text, nullable=True)

    snapshot: Mapped[Snapshot] = relationship(back_populates="records")


Base.metadata.create_all(engine)

app = FastAPI(title="Acompanhamento de Bases", version="1.0.0")
app.mount("/static", StaticFiles(directory=BASE_DIR / "app" / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "app" / "templates")


def s(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text if text else None


def parse_days(value: Any) -> int | None:
    text = (s(value) or "").upper()
    if not text or text == "X" or "IMPLANTA" in text:
        return None
    match = re.search(r"\d+", text)
    return int(match.group()) if match else None


def parse_camera_number(value: Any) -> int | None:
    text = (s(value) or "").upper()
    if not text or text == "X":
        return None
    # Aceita CAM / câmera e também o erro comum ACM, mas evita interpretar nomes de localidades.
    if "CAM" not in text and "ACM" not in text and not text.isdigit():
        return None
    match = re.search(r"\d+", text)
    return int(match.group()) if match else None


def normalize_status(value: Any) -> str:
    return (s(value) or "NÃO INFORMADO").upper()


def quality_note(recording_raw: str | None, camera_raw: str | None, offline_raw: str | None) -> str | None:
    notes: list[str] = []
    if camera_raw and parse_camera_number(camera_raw) is None and camera_raw.upper() != "X":
        notes.append("Qtd. de câmeras não numérica")
    if offline_raw and parse_camera_number(offline_raw) is None and offline_raw.upper() != "X":
        notes.append("Câmeras off-line não numérica")
    if offline_raw and "ACM" in offline_raw.upper():
        notes.append("Possível digitação 'ACM' em vez de 'CAM'")
    if not recording_raw:
        notes.append("Dias de gravação não informado")
    return "; ".join(notes) if notes else None


def workbook_to_payload(content: bytes) -> tuple[str | None, list[dict[str, Any]]]:
    try:
        wb = load_workbook(io.BytesIO(content), data_only=True)
    except Exception as exc:
        raise ValueError(f"Não foi possível abrir o Excel: {exc}") from exc

    ws = wb.active
    source_date = None
    header_date = ws.cell(1, 5).value
    if isinstance(header_date, datetime):
        source_date = header_date.strftime("%d/%m/%Y")
    elif header_date is not None:
        source_date = str(header_date).strip()

    rows: list[dict[str, Any]] = []
    for values in ws.iter_rows(min_row=2, min_col=1, max_col=5, values_only=True):
        base_name = s(values[0])
        if not base_name:
            continue
        rec_raw = s(values[1])
        cam_raw = s(values[2])
        off_raw = s(values[3])
        status = normalize_status(values[4])
        rows.append({
            "base": base_name,
            "recording_days_raw": rec_raw,
            "recording_days": parse_days(rec_raw),
            "camera_count_raw": cam_raw,
            "camera_count": parse_camera_number(cam_raw),
            "offline_raw": off_raw,
            "offline_count": parse_camera_number(off_raw),
            "status": status,
            "is_implantation": 1 if "IMPLANTA" in (rec_raw or "").upper() else 0,
            "quality_note": quality_note(rec_raw, cam_raw, off_raw),
        })
    if not rows:
        raise ValueError("A planilha não possui bases válidas a partir da linha 2.")
    return source_date, rows


def payload_hash(rows: list[dict[str, Any]]) -> str:
    normalized = json.dumps(rows, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
    return hashlib.sha256(normalized).hexdigest()


def save_snapshot(rows: list[dict[str, Any]], source: str, source_date: str | None, captured_at: datetime | None = None) -> Snapshot:
    snap = Snapshot(
        captured_at=captured_at or datetime.now(timezone.utc),
        source=source,
        source_date=source_date,
        row_count=len(rows),
        data_hash=payload_hash(rows),
    )
    with SessionLocal() as db:
        db.add(snap)
        db.flush()
        for row in rows:
            db.add(BaseRecord(snapshot_id=snap.id, **row))
        db.commit()
        db.refresh(snap)
    return snap


def latest_snapshot(db: Session) -> Snapshot | None:
    return db.scalar(select(Snapshot).order_by(Snapshot.captured_at.desc(), Snapshot.id.desc()).limit(1))


def previous_snapshot(db: Session, current_id: int) -> Snapshot | None:
    return db.scalar(select(Snapshot).where(Snapshot.id != current_id).order_by(Snapshot.captured_at.desc(), Snapshot.id.desc()).limit(1))


def classify_status(status: str) -> str:
    u = status.upper()
    if u == "EM FUNCIONAMENTO":
        return "operational"
    if "PARCIAL" in u:
        return "partial"
    if "DESATIV" in u:
        return "deactivated"
    if "SEM FUNCIONAMENTO" in u:
        return "offline"
    return "other"


def snapshot_metrics(records: list[BaseRecord]) -> dict[str, Any]:
    counts = {"operational": 0, "partial": 0, "offline": 0, "deactivated": 0, "other": 0}
    for r in records:
        counts[classify_status(r.status)] += 1
    cameras_total = sum((r.camera_count or 0) for r in records)
    cameras_offline = sum((r.offline_count or 0) for r in records)
    cameras_online = max(cameras_total - cameras_offline, 0)
    known_cam_bases = sum(1 for r in records if r.camera_count is not None)
    implant = sum(r.is_implantation for r in records)
    quality = sum(1 for r in records if r.quality_note)
    total = len(records)
    return {
        "total": total,
        **counts,
        "operational_rate": round((counts["operational"] / total * 100), 1) if total else 0,
        "cameras_total": cameras_total,
        "cameras_offline": cameras_offline,
        "cameras_online": cameras_online,
        "known_cam_bases": known_cam_bases,
        "implantation": implant,
        "quality_issues": quality,
    }


def build_changes(current: list[BaseRecord], previous: list[BaseRecord] | None) -> list[dict[str, Any]]:
    if not previous:
        return []
    old = {r.base: r for r in previous}
    new = {r.base: r for r in current}
    changes: list[dict[str, Any]] = []
    for base_name in sorted(set(old) | set(new)):
        a, b = old.get(base_name), new.get(base_name)
        if a is None:
            changes.append({"base": base_name, "type": "Nova base", "before": "—", "after": "Incluída"})
            continue
        if b is None:
            changes.append({"base": base_name, "type": "Base removida", "before": "Existia", "after": "—"})
            continue
        comparisons = [
            ("Status", a.status, b.status),
            ("Dias de gravação", a.recording_days_raw or "—", b.recording_days_raw or "—"),
            ("Qtd. câmeras", a.camera_count_raw or "—", b.camera_count_raw or "—"),
            ("Câmeras off-line", a.offline_raw or "—", b.offline_raw or "—"),
        ]
        for kind, before, after in comparisons:
            if before != after:
                changes.append({"base": base_name, "type": kind, "before": before, "after": after})
    return changes


def serialize_record(r: BaseRecord) -> dict[str, Any]:
    return {
        "base": r.base,
        "recording_days_raw": r.recording_days_raw or "—",
        "recording_days": r.recording_days,
        "camera_count_raw": r.camera_count_raw or "—",
        "camera_count": r.camera_count,
        "offline_raw": r.offline_raw or "—",
        "offline_count": r.offline_count,
        "status": r.status,
        "is_implantation": bool(r.is_implantation),
        "quality_note": r.quality_note or "",
    }


def iso_utc(dt: datetime) -> str:
    """Serializa timestamps do banco sempre como UTC para o navegador converter corretamente."""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc).isoformat()


def dashboard_context() -> dict[str, Any]:
    with SessionLocal() as db:
        snap = latest_snapshot(db)
        if not snap:
            return {"snapshot": None}
        current_records = list(db.scalars(select(BaseRecord).where(BaseRecord.snapshot_id == snap.id).order_by(BaseRecord.base)))
        prev = previous_snapshot(db, snap.id)
        prev_records = list(db.scalars(select(BaseRecord).where(BaseRecord.snapshot_id == prev.id))) if prev else None
        metrics = snapshot_metrics(current_records)
        prev_metrics = snapshot_metrics(prev_records) if prev_records else None
        changes = build_changes(current_records, prev_records)

        statuses: dict[str, int] = {}
        day_bands = {"0–14 dias": 0, "15–30 dias": 0, "31–45 dias": 0, "46+ dias": 0, "Sem dado": 0, "Em implantação": 0}
        for r in current_records:
            statuses[r.status] = statuses.get(r.status, 0) + 1
            if r.is_implantation:
                day_bands["Em implantação"] += 1
            elif r.recording_days is None:
                day_bands["Sem dado"] += 1
            elif r.recording_days <= 14:
                day_bands["0–14 dias"] += 1
            elif r.recording_days <= 30:
                day_bands["15–30 dias"] += 1
            elif r.recording_days <= 45:
                day_bands["31–45 dias"] += 1
            else:
                day_bands["46+ dias"] += 1

        alerts = sorted(
            [serialize_record(r) for r in current_records if classify_status(r.status) != "operational" or (r.offline_count or 0) > 0],
            key=lambda x: (0 if (x["offline_count"] or 0) > 0 else 1, x["status"], x["base"]),
        )

        snaps = list(reversed(list(db.scalars(select(Snapshot).order_by(Snapshot.captured_at.desc(), Snapshot.id.desc()).limit(60)))))
        history = []
        for h in snaps:
            rs = list(db.scalars(select(BaseRecord).where(BaseRecord.snapshot_id == h.id)))
            hm = snapshot_metrics(rs)
            history.append({
                "id": h.id,
                "captured_at": iso_utc(h.captured_at),
                "source_date": h.source_date,
                "source": h.source,
                "total": hm["total"],
                "operational": hm["operational"],
                "operational_rate": hm["operational_rate"],
                "cameras_offline": hm["cameras_offline"],
                "changes": None,
            })
        if history:
            for idx, item in enumerate(history):
                item["changes"] = 0 if idx == 0 else None

        return {
            "snapshot": {
                "id": snap.id,
                "captured_at": iso_utc(snap.captured_at),
                "source_date": snap.source_date,
                "source": snap.source,
                "row_count": snap.row_count,
            },
            "metrics": metrics,
            "prev_metrics": prev_metrics,
            "records": [serialize_record(r) for r in current_records],
            "alerts": alerts,
            "changes": changes,
            "statuses": statuses,
            "day_bands": day_bands,
            "history": history,
        }


@app.on_event("startup")
def startup_seed() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with SessionLocal() as db:
        exists = db.scalar(select(Snapshot.id).limit(1))
    if not exists and SEED_FILE.exists():
        content = SEED_FILE.read_bytes()
        source_date, rows = workbook_to_payload(content)
        # A medição vem da planilha; a captura é o momento em que o app inicializa.
        save_snapshot(rows, source="excel-inicial", source_date=source_date)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request):
    context = dashboard_context()
    return templates.TemplateResponse("dashboard.html", {"request": request, **context})


@app.post("/importar")
async def importar(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        return RedirectResponse(url="/?import=erro-formato", status_code=303)
    content = await file.read()
    try:
        source_date, rows = workbook_to_payload(content)
        snap = save_snapshot(rows, source="excel-upload", source_date=source_date)
    except Exception:
        return RedirectResponse(url="/?import=erro", status_code=303)
    return RedirectResponse(url=f"/?import=ok&snapshot={snap.id}", status_code=303)


@app.post("/api/ingest")
async def api_ingest(request: Request):
    """Ponto pronto para a futura integração Feishu.

    Espera JSON: {source_date, source, rows:[{base, recording_days, camera_count, offline_count, status}]}
    Os valores podem vir crus (strings) e são normalizados do mesmo modo que o Excel.
    """
    payload = await request.json()
    incoming = payload.get("rows")
    if not isinstance(incoming, list) or not incoming:
        raise HTTPException(400, "rows é obrigatório e deve conter registros")
    rows: list[dict[str, Any]] = []
    for item in incoming:
        if not isinstance(item, dict) or not s(item.get("base")):
            continue
        rec_raw = s(item.get("recording_days") or item.get("recording_days_raw"))
        cam_raw = s(item.get("camera_count") or item.get("camera_count_raw"))
        off_raw = s(item.get("offline_count") or item.get("offline_raw"))
        status = normalize_status(item.get("status"))
        rows.append({
            "base": s(item.get("base")),
            "recording_days_raw": rec_raw,
            "recording_days": parse_days(rec_raw),
            "camera_count_raw": cam_raw,
            "camera_count": parse_camera_number(cam_raw),
            "offline_raw": off_raw,
            "offline_count": parse_camera_number(off_raw),
            "status": status,
            "is_implantation": 1 if "IMPLANTA" in (rec_raw or "").upper() else 0,
            "quality_note": quality_note(rec_raw, cam_raw, off_raw),
        })
    if not rows:
        raise HTTPException(400, "nenhuma base válida recebida")
    snap = save_snapshot(rows, source=s(payload.get("source")) or "feishu-api", source_date=s(payload.get("source_date")))
    return {"ok": True, "snapshot_id": snap.id, "rows": len(rows), "captured_at": iso_utc(snap.captured_at)}


@app.get("/api/dashboard")
def api_dashboard():
    return dashboard_context()


@app.get("/relatorio.xlsx")
def relatorio_excel():
    ctx = dashboard_context()
    if not ctx.get("snapshot"):
        raise HTTPException(404, "Sem dados")
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"
    title_fill = PatternFill("solid", fgColor="C9152B")
    header_fill = PatternFill("solid", fgColor="1F2937")
    white = "FFFFFF"
    ws["A1"] = "Acompanhamento de Bases — Relatório"
    ws["A1"].font = Font(bold=True, size=16, color=white)
    ws["A1"].fill = title_fill
    ws.merge_cells("A1:D1")
    rows = [
        ("Medição da planilha", ctx["snapshot"]["source_date"] or "—"),
        ("Momento da captura", ctx["snapshot"]["captured_at"]),
        ("Total de bases", ctx["metrics"]["total"]),
        ("Em funcionamento", ctx["metrics"]["operational"]),
        ("Sem funcionamento", ctx["metrics"]["offline"]),
        ("Em implantação", ctx["metrics"]["implantation"]),
        ("Câmeras conhecidas", ctx["metrics"]["cameras_total"]),
        ("Câmeras off-line", ctx["metrics"]["cameras_offline"]),
        ("Alertas de qualidade", ctx["metrics"]["quality_issues"]),
    ]
    for idx, (label, value) in enumerate(rows, 3):
        ws.cell(idx, 1, label).font = Font(bold=True)
        ws.cell(idx, 2, value)
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 28

    data_ws = wb.create_sheet("Bases atuais")
    headers = ["Base", "Dias de gravação", "Qtd. câmeras", "Câmeras off-line", "Status", "Em implantação", "Observação de qualidade"]
    data_ws.append(headers)
    for cell in data_ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color=white)
    for r in ctx["records"]:
        data_ws.append([r["base"], r["recording_days_raw"], r["camera_count_raw"], r["offline_raw"], r["status"], "SIM" if r["is_implantation"] else "NÃO", r["quality_note"]])
    for col, width in {"A":18,"B":20,"C":18,"D":20,"E":32,"F":16,"G":48}.items():
        data_ws.column_dimensions[col].width = width
    for row in data_ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ch_ws = wb.create_sheet("Alterações")
    ch_ws.append(["Base", "Campo", "Antes", "Depois"])
    for cell in ch_ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color=white)
    for c in ctx["changes"]:
        ch_ws.append([c["base"], c["type"], c["before"], c["after"]])
    for col, width in {"A":18,"B":24,"C":32,"D":32}.items():
        ch_ws.column_dimensions[col].width = width

    hist_ws = wb.create_sheet("Histórico")
    hist_ws.append(["Snapshot", "Momento", "Data fonte", "Origem", "Bases", "Em funcionamento", "% funcionamento", "Câmeras off-line"])
    for cell in hist_ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color=white)
    for h in ctx["history"]:
        hist_ws.append([h["id"], h["captured_at"], h["source_date"], h["source"], h["total"], h["operational"], h["operational_rate"], h["cameras_offline"]])
    for col, width in {"A":12,"B":28,"C":16,"D":20,"E":12,"F":20,"G":18,"H":20}.items():
        hist_ws.column_dimensions[col].width = width

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    filename = f"relatorio_acompanhamento_bases_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return StreamingResponse(out, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="{filename}"'})
