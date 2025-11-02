# models.py
from datetime import datetime
from sqlalchemy import Column, Integer, String, Text, Boolean, DateTime, ForeignKey
from sqlalchemy.orm import relationship

# ❗️ نکته مهم: Base را از db ایمپورت کن، نه declarative_base()
from db import Base


# -----------------------
# Vocabulary (جدید)
# -----------------------
class Vocabulary(Base):
    __tablename__ = "vocabulary"
    id = Column(Integer, primary_key=True)
    word = Column(String(128), nullable=False, index=True, unique=False)
    sentence = Column(Text)
    synonym = Column(String(256))
    native_meaning = Column(Text)
    pos = Column(String(32))         # همان type در UI
    base_word = Column(String(128))
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False, index=True)


    def __repr__(self):
        return f"<Vocabulary(id={self.id}, word='{self.word}')>"


# -----------------------
# Goals / TimeEntries
# -----------------------
class Goal(Base):
    __tablename__ = "goals"

    id = Column(Integer, primary_key=True)
    title = Column(String(255), nullable=False)
    category = Column(String(100), nullable=True)
    notes = Column(Text, nullable=True)
    is_active = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    time_entries = relationship("TimeEntry", back_populates="goal", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Goal(id={self.id}, title='{self.title}', is_active={self.is_active})>"


class TimeEntry(Base):
    __tablename__ = "time_entries"

    id = Column(Integer, primary_key=True)
    goal_id = Column(Integer, ForeignKey('goals.id'), nullable=False)
    started_at = Column(DateTime, nullable=False)
    minutes = Column(Integer, nullable=False)
    note = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    goal = relationship("Goal", back_populates="time_entries")

    def __repr__(self):
        return f"<TimeEntry(id={self.id}, goal_id={self.goal_id}, minutes={self.minutes})>"
