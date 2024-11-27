from sqlalchemy import Column, Integer, String, ForeignKey, Enum
from sqlalchemy.orm import relationship
from database import Base


class Class(Base):
    __tablename__ = 'classes'
    id = Column(Integer, primary_key=True)
    name = Column(String, index=True, unique=True)
    shift = Column(Integer)
    quantity = Column(Integer)
    lessons_min = Column(Integer)
    lessons_max = Column(Integer)


class Lesson(Base):
    __tablename__ = 'lessons'
    id = Column(Integer, primary_key=True, index=True)
    class_id = Column(Integer, ForeignKey('classes.id'))
    day = Column(String)
    lesson_number = Column(Integer)
    subject_id = Column(Integer, ForeignKey('subjects.id'))
    cabinet_id = Column(Integer, ForeignKey('cabinets.id'))
    teacher_id = Column(Integer, ForeignKey('teachers.id'))

    # class_ = relationship("Class", back_populates="schedules")
    # subject = relationship("Subject", back_populates="schedules")
    # cabinet = relationship("Cabinet", back_populates="schedules")
    # teacher = relationship("Teacher", back_populates="schedules")

class Teacher(Base):
    __tablename__ = 'teachers'
    id = Column(Integer, primary_key=True, index=True)
    surname = Column(String, index=True)
    name_last_name = Column(String)
    base_cabinet = Column(Integer, ForeignKey('cabinets.id'))
    # specializations = relationship("TeacherSpecialization", back_populates="teacher")
    # schedules = relationship("Schedule", back_populates="teacher")

    def __str__(self):
        return f'{self.surname} {self.name_last_name}'

    def __repr__(self):
        return self.__str__()


class TeacherSpecialization(Base):
    __tablename__ = 'teacher_specializations'
    id = Column(Integer, primary_key=True, index=True)
    teacher_id = Column(Integer, ForeignKey('teachers.id'))
    subject_id = Column(Integer, ForeignKey('subjects.id'))

    # teacher = relationship("Teacher", back_populates="specializations")
    # subject = relationship("Subject", back_populates="teacher_specializations")

class ClassTeacherSubjectConnection(Base):
    __tablename__ = 'class_teacher_subject_connections'
    id = Column(Integer, primary_key=True)
    class_id = Column(Integer, ForeignKey('classes.id'))
    times_a_week = Column(Integer)
    subject_id = Column(Integer, ForeignKey('subjects.id'))
    teacher_id = Column(Integer, ForeignKey('teachers.id'))
    cabinet_id = Column(Integer, ForeignKey('cabinets.id'), nullable=True)
    weekly = Column(Enum('every', 'even', 'odd', name='weekly'), default='every')


class Subject(Base):
    __tablename__ = 'subjects'
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, index=True, unique=True)
    # schedules = relationship("Schedule", back_populates="subject")
    # teacher_specializations = relationship("TeacherSpecialization", back_populates="subject")

class Cabinet(Base):
    __tablename__ = 'cabinets'
    id = Column(Integer, primary_key=True, index=True)
    number = Column(String, index=True, unique=True)
    capacity = Column(Integer)

    # schedules = relationship("Schedule", back_populates="cabinet")
