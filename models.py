from datetime import datetime
from sqlalchemy import Column, Integer, String, Text, Boolean, DateTime, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship

Base = declarative_base()


class Goal(Base):
    __tablename__ = 'goals'
    
    id = Column(Integer, primary_key=True)
    title = Column(String(255), nullable=False)
    category = Column(String(100), nullable=True)
    notes = Column(Text, nullable=True)
    is_active = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    
    # Relationship to time entries
    time_entries = relationship("TimeEntry", back_populates="goal", cascade="all, delete-orphan")
    
    def __repr__(self):
        return f"<Goal(id={self.id}, title='{self.title}', is_active={self.is_active})>"


class TimeEntry(Base):
    __tablename__ = 'time_entries'
    
    id = Column(Integer, primary_key=True)
    goal_id = Column(Integer, ForeignKey('goals.id'), nullable=False)
    started_at = Column(DateTime, nullable=False)
    minutes = Column(Integer, nullable=False)
    note = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    
    # Relationship to goal
    goal = relationship("Goal", back_populates="time_entries")
    
    def __repr__(self):
        return f"<TimeEntry(id={self.id}, goal_id={self.goal_id}, minutes={self.minutes})>"
