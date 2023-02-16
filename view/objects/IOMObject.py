class Question:
    def __init__(self, field):
        self.name = field.Name
        self.fullname = field.FullName
        self.label = field.Label
        self.questions = Questions()
    
class Questions:
    def __init__(self):
        self.questions = list()

    def add_question(self, question):
        self.questions.append(question)

    def find_question(self, question):
        for qre in self.questions:
            if qre.Name == question.Name and qre.FullName == question.FullName:
                return qre
        return None
        