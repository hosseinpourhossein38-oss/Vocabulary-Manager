import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from models import Base

# Database configuration
DATABASE_URL = os.environ.get('DATABASE_URL')
if DATABASE_URL:
    # For production (Render) - use PostgreSQL
    if DATABASE_URL.startswith('postgres://'):
        DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)
    engine = create_engine(DATABASE_URL)
else:
    # For development - use SQLite
    engine = create_engine('sqlite:///data/app.db')

# Create session factory
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

def get_db():
    """Dependency to get database session"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def init_db():
    """Initialize database tables"""
    # Ensure data directory exists
    os.makedirs('data', exist_ok=True)
    
    # Create all tables
    Base.metadata.create_all(bind=engine)
