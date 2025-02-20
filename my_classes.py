class TeacherChanges:
    __slots__ = ('lesson_n', 'prev_subj', 'prev_class', 'prev_cabinet',
                 'new_subj', 'new_class', 'new_cabinet')

    def __init__(self, lesson_n:int, prev_subj=None, prev_class=None, prev_cabinet=None,
                 new_subj=None, new_class=None, new_cabinet=None):
        self.lesson_n = lesson_n
        self.prev_subj = prev_subj
        self.prev_class = prev_class
        self.prev_cabinet = prev_cabinet
        self.new_subj = new_subj
        self.new_class = new_class
        self.new_cabinet = new_cabinet
