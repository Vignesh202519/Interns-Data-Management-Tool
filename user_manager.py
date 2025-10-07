class UserManager:
    def __init__(self):
        self.users = {
            "admin": "123",
            "vignesh": "password"
        }

    def validate_user(self, username, password):
        return self.users.get(username) == password


class StudentManager:
    def __init__(self):
        # store student records as list of dicts
        self.students = []
        self.next_id = 1

    def add_student(self, student_data):
        student_data["id"] = self.next_id
        self.students.append(student_data)
        self.next_id += 1

    def get_all_students(self):
        return self.students

    def get_student(self, student_id):
        for student in self.students:
            if student["id"] == student_id:
                return student
        return None

    def update_student(self, student_id, updated_data):
        for student in self.students:
            if student["id"] == student_id:
                student.update(updated_data)
                return True
        return False
    
    def delete_student(self, student_id):
        self.students = [s for s in self.students if s["id"] != student_id]
