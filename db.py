import os
from pathlib import Path
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

# .env کنار همین فایل
env_path = Path(__file__).resolve().parent / ".env"
load_dotenv(dotenv_path=env_path, override=True)

raw_url = (os.getenv("DATABASE_URL") or "").strip()
if not raw_url:
    raise RuntimeError("DATABASE_URL is missing. Check your .env")

url = raw_url
if url.startswith("postgres://"):
    url = url.replace("postgres://", "postgresql+psycopg2://", 1)
elif url.startswith("postgresql://") and "+psycopg2" not in url:
    url = url.replace("postgresql://", "postgresql+psycopg2://", 1)

engine = create_engine(url, pool_pre_ping=True, future=True)
SessionLocal = sessionmaker(bind=engine, autocommit=False, autoflush=False, future=True)

# ❗️فقط همین‌جا Base را تعریف کن. هیچ ایمپورتی از models نکن.
Base = declarative_base()

# db.py
# ... همین هایی که قبلاً گذاشتیم: engine, SessionLocal, Base

def init_db():
    import models  # ← ثبت همه مدل‌ها
    Base.metadata.create_all(bind=engine)
