from database import get_db
from db_models import Class, Schedule, Teacher, TeacherSchedule, Cabinet

class MainController:
    def __init__(self):
        self.db = next(get_db())

    def run(self):
        # Пример запуска контроллера
        self.add_example_data()
        self.get_free_teachers(day='Понедельник', lesson_number=1)

    def add_example_data(self):
        # Пример добавления данных
        class_ = Class(name='10A')
        self.db.add(class_)
        self.db.commit()

        schedule = Schedule(class_id=class_.id, day='Понедельник', lesson_number=1, subject='Математика', cabinet='101')
        self.db.add(schedule)
        self.db.commit()

        teacher = Teacher(last_name='Иванов', first_name='Иван', middle_name='Иванович')
        self.db.add(teacher)
        self.db.commit()

        teacher_schedule = TeacherSchedule(teacher_id=teacher.id, day='Понедельник', lesson_number=1, class_name='10A', subject='Математика', cabinet='101')
        self.db.add(teacher_schedule)
        self.db.commit()

    def get_free_teachers(self, day, lesson_number):
        # Получение списка свободных учителей на определенный урок
        busy_teachers = self.db.query(TeacherSchedule.teacher_id).filter_by(day=day, lesson_number=lesson_number).all()
        busy_teacher_ids = [teacher.teacher_id for teacher in busy_teachers]
        free_teachers = self.db.query(Teacher).filter(~Teacher.id.in_(busy_teacher_ids)).all()

        for teacher in free_teachers:
            print(f"{teacher.last_name} {teacher.first_name} {teacher.middle_name}")

    def close(self):
        self.db.close()
