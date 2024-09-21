from sqlalchemy import Column, Integer, String, ForeignKey, Table
from sqlalchemy.orm import relationship
from database import Base

class Class(Base):
    __tablename__ = 'classes'
    id = Column(Integer, primary_key=True)
    name = Column(String, index=True, unique=True)

class Schedule(Base):
    __tablename__ = 'schedules'
    id = Column(Integer, primary_key=True, index=True)
    class_id = Column(Integer, ForeignKey('classes.id'))
    day = Column(String)
    lesson_number = Column(Integer)
    subject_id = Column(Integer, ForeignKey('subjects.id'))
    cabinet_id = Column(Integer, ForeignKey('cabinets.id'))
    teacher_id = Column(Integer, ForeignKey('teachers.id'))

    class_ = relationship("Class", back_populates="schedules")
    subject = relationship("Subject", back_populates="schedules")
    cabinet = relationship("Cabinet", back_populates="schedules")
    teacher = relationship("Teacher", back_populates="schedules")

class Teacher(Base):
    __tablename__ = 'teachers'
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, index=True, unique=True)
    specializations = relationship("TeacherSpecialization", back_populates="teacher")
    schedules = relationship("Schedule", back_populates="teacher")

class TeacherSpecialization(Base):
    __tablename__ = 'teacher_specializations'
    id = Column(Integer, primary_key=True, index=True)
    teacher_id = Column(Integer, ForeignKey('teachers.id'))
    subject_id = Column(Integer, ForeignKey('subjects.id'))

    teacher = relationship("Teacher", back_populates="specializations")
    subject = relationship("Subject", back_populates="teacher_specializations")

class TeacherSchedule(Base):
    __tablename__ = 'teacher_schedules'
    id = Column(Integer, primary_key=True)
    teacher_id = Column(Integer, ForeignKey('teachers.id'))
    day = Column(String, nullable=False)
    lesson_number = Column(Integer, nullable=False)
    class_id = Column(Integer, ForeignKey('classes.id'))
    subject_id = Column(Integer, ForeignKey('subjects.id'))
    cabinet_id = Column(Integer, ForeignKey('cabinets.id'))
    teacher = relationship('Teacher', back_populates='schedules')

class Subject(Base):
    __tablename__ = 'subjects'
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, index=True, unique=True)
    schedules = relationship("Schedule", back_populates="subject")
    teacher_specializations = relationship("TeacherSpecialization", back_populates="subject")

class Cabinet(Base):
    __tablename__ = 'cabinets'
    id = Column(Integer, primary_key=True, index=True)
    number = Column(String, index=True, unique=True)
    schedules = relationship("Schedule", back_populates="cabinet")
