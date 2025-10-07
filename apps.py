from flask import Flask, render_template, request, redirect, url_for, session, flash
import pandas as pd, os, re, secrets
from datetime import datetime, timedelta
import matplotlib.pyplot as plt, io, base64
from collections import Counter

app = Flask(__name__)
app.secret_key = "secret123"
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(days=7)

EXCEL_FILE = "interns.xlsx"
REGISTRATION_FILE = "registers.xlsx"
users_db, reset_tokens = [], {}
users = {
    "user1@example.com": {"username": "user1", "password": "pass1", "name": "Vignesh"},
    "admin": {"username": "admin", "password": "admin123", "name": "Administrator"},
    "user1": {"username": "user1", "password": "pass1", "name": "Vignesh"}
}

# ----------------- EXCEL FUNCTIONS FOR USER REGISTRATION -----------------
def save_user_to_excel(user_data):
    """Save user registration data to Excel file"""
    try:
        df_new = pd.DataFrame([user_data])
        
        if os.path.exists(REGISTRATION_FILE):
            # Append to existing file
            df_existing = pd.read_excel(REGISTRATION_FILE)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            # Create new file
            df_combined = df_new
        
        # Save to Excel
        df_combined.to_excel(REGISTRATION_FILE, index=False)
        return True
    except Exception as e:
        print(f"Error saving user to Excel: {e}")
        return False

def load_users_from_excel():
    """Load users from Excel file into memory"""
    if not os.path.exists(REGISTRATION_FILE):
        return []
    
    try:
        df = pd.read_excel(REGISTRATION_FILE)
        users = df.to_dict('records')
        return users
    except Exception as e:
        print(f"Error loading users from Excel: {e}")
        return []

def check_existing_user(username, email):
    """Check if username or email already exists in Excel"""
    if not os.path.exists(REGISTRATION_FILE):
        return False
    
    try:
        df = pd.read_excel(REGISTRATION_FILE)
        
        # Check if username exists
        if not df[df['username'] == username].empty:
            return True
        
        # Check if email exists
        if not df[df['email'] == email].empty:
            return True
        
        return False
    except Exception as e:
        print(f"Error checking existing user: {e}")
        return False

def authenticate_from_excel(username, password):
    """Authenticate user from Excel file"""
    if not os.path.exists(REGISTRATION_FILE):
        print("Excel file not found")
        return None
    
    try:
        df = pd.read_excel(REGISTRATION_FILE)
        print(f"Excel columns: {df.columns.tolist()}")
        print(f"Looking for user: {username}")
        
        # Check if username/email and password match
        user_row = df[
            ((df['username'] == username) | (df['email'] == username)) & 
            (df['password'] == password)
        ]
        
        print(f"Found {len(user_row)} matching users")
        
        if not user_row.empty:
            user_data = user_row.iloc[0].to_dict()
            print(f"User found: {user_data}")
            return {
                "fullname": user_data.get('fullname', ''),
                "username": user_data.get('username', ''),
                "email": user_data.get('email', ''),
                "password": user_data.get('password', ''),
                "created_at": user_data.get('created_at', '')
            }
        
        return None
    except Exception as e:
        print(f"Error reading Excel file for authentication: {e}")
        return None

# Load existing users from Excel on startup
def load_all_users():
    """Load all registered users from Excel on startup"""
    global users_db
    
    excel_users = load_users_from_excel()
    if excel_users:
        users_db.extend(excel_users)
        
        # Also populate the users dict for backward compatibility
        for user in excel_users:
            users[user['email']] = {
                "username": user['username'],
                "password": user['password'],
                "name": user['fullname']
            }
        print(f"Loaded {len(excel_users)} users from Excel file")

# Initialize users on app start
load_all_users()

# ----------------- STUDENT MANAGEMENT -----------------
class StudentManager:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.students = []
        self.next_id = 1
        self.load_from_excel()

    def load_from_excel(self):
        """Load students from Excel file on startup"""
        if os.path.exists(self.excel_file):
            df = pd.read_excel(self.excel_file)
            if not df.empty:
                # Convert DataFrame to list of dictionaries
                self.students = df.to_dict('records')
                # Set next_id based on existing data
                if self.students:
                    max_id = max(int(student.get('id', 0)) for student in self.students)
                    self.next_id = max_id + 1
                # Ensure all records have id field
                for i, student in enumerate(self.students):
                    if 'id' not in student:
                        student['id'] = i + 1

    def save_to_excel(self):
        """Save current students to Excel file"""
        if self.students:
            df = pd.DataFrame(self.students)
            df.to_excel(self.excel_file, index=False)

    def add_student(self, data):
        data['id'] = self.next_id
        self.students.append(data)
        self.next_id += 1
        self.save_to_excel()  # Save to Excel after adding

    def all_students(self):
        return self.students

    def get_student(self, sid):
        return next((s for s in self.students if s['id'] == sid), None)

    def update_student(self, sid, data):
        for s in self.students:
            if s['id'] == sid:
                s.update(data)
                break  # Break after updating the student
        self.save_to_excel()  # Save to Excel after updating

    def delete_student(self, sid):
        self.students = [s for s in self.students if s['id'] != sid]
        self.save_to_excel()  # Save to Excel after deleting

# Initialize with Excel file
student_manager = StudentManager(EXCEL_FILE)

# ----------------- UTILITIES -----------------
def extract_interests(txt):
    if not txt:
        return []
    
    # Convert to string in case it's a float or other type
    txt = str(txt) if txt is not None else ""
    
    return [i.strip().title() for i in re.split(r'[,;|/\n]', txt) if len(i.strip()) > 2]

def generate_chart():
    if not os.path.exists(EXCEL_FILE):
        return None
    
    try:
        df = pd.read_excel(EXCEL_FILE)
        
        # Check if the Excel file has data
        if df.empty:
            return None
        
        # Use 'interest' column instead of 'Area of Interest'
        interests_column = 'interest'
        
        if interests_column not in df.columns:
            # If 'interest' column doesn't exist, try to find a similar column
            possible_columns = ['interest', 'Area of Interest', 'area_of_interest', 'interests']
            for col in possible_columns:
                if col in df.columns:
                    interests_column = col
                    break
            else:
                # No interest column found
                return None
        
        interests = []
        for t in df[interests_column].dropna():
            # Ensure we're passing a string to extract_interests
            interests.extend(extract_interests(str(t) if t is not None else ""))
        
        if not interests:
            return None
        
        top = Counter(interests).most_common(10)
        plt.figure(figsize=(8,5))
        names, counts = zip(*top)
        bars = plt.bar(names, counts, color=plt.cm.Set3(range(len(names))), edgecolor='black')
        plt.xticks(rotation=45, ha='right')
        for bar, count in zip(bars, counts):
            plt.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.1, str(count), ha='center', va='bottom')
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight')
        img.seek(0)
        plt.close()
        return f"data:image/png;base64,{base64.b64encode(img.getvalue()).decode()}"
    
    except Exception as e:
        print(f"Error generating chart: {e}")
        return None

# ----------------- ROUTES -----------------

@app.route("/")
def index():
    return redirect(url_for("login"))

@app.route("/login", methods=["GET","POST"])
def login():
    if "user" in session:
        return redirect(url_for("home"))

    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")
        
        # First check in-memory users
        user = next((u for u in users_db if (u['username']==username or u['email']==username) and u['password']==password), None)
        
        # If not found in memory, check Excel file
        if not user:
            user = authenticate_from_excel(username, password)
            # If found in Excel, add to in-memory storage
            if user and user not in users_db:
                users_db.append(user)
        
        # Check default users
        if not user and username in users and users[username]["password"] == password:
            user = users[username]
            
        if user:
            session["user"] = user.get('username', username)
            session["user_name"] = user.get('name', user.get('fullname', 'User'))
            session.permanent = bool(request.form.get("remember"))
            return redirect(url_for("home"))
        return render_template("login.html", error="Invalid credentials!")
    return render_template("login.html")

@app.route("/register", methods=["GET","POST"])
def register():
    if "user" in session:
        return redirect(url_for("home"))

    if request.method == "POST":
        fullname = request.form.get('fullname', '').strip()
        username = request.form.get('username', '').strip()
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')

        errors = []
        if not fullname: errors.append("Full name required")
        if not username: errors.append("Username required")
        if not email: errors.append("Email required")
        if not password: errors.append("Password required")

        # Check if user already exists in both memory and Excel
        if any(u['username'] == username for u in users_db) or check_existing_user(username, email):
            errors.append("Username already exists")
        
        if any(u['email'] == email for u in users_db) or check_existing_user(username, email):
            errors.append("Email already registered")

        if errors:
            for e in errors: flash(e, 'error')
            return render_template("register.html", fullname=fullname, username=username, email=email)

        # Prepare user data
        user_data = {
            "fullname": fullname,
            "username": username,
            "email": email,
            "password": password,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        # Save to Excel file
        if save_user_to_excel(user_data):
            # Also add to in-memory storage for immediate access
            users_db.append(user_data)
            users[email] = {"username": username, "password": password, "name": fullname}
            flash("Registration successful! Please log in.", "success")
            return redirect(url_for("login"))
        else:
            flash("Error saving registration. Please try again.", "error")
            return render_template("register.html", fullname=fullname, username=username, email=email)

    return render_template("register.html")

@app.route("/forgot-password", methods=["GET","POST"])
def forgot_password():
    if "user" in session:
        return redirect(url_for("home"))

    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        if email in users or any(u['email']==email for u in users_db):
            token = secrets.token_urlsafe(32)
            reset_tokens[token] = {"email": email, "timestamp": datetime.now()}
            print(f"Password reset link: http://localhost:5000/reset-password/{token}")
            flash("Password reset link sent to your email (check console in dev mode).", "success")
            return redirect(url_for("login"))
        return render_template("forgot_password.html", error="Email not found")
    return render_template("forgot_password.html")

@app.route("/reset-password/<token>", methods=["GET","POST"])
def reset_password(token):
    if token not in reset_tokens:
        return render_template("reset_password.html", error="Invalid or expired token")
    
    token_data = reset_tokens[token]
    # Check if token is expired (24 hours)
    if datetime.now() - token_data["timestamp"] > timedelta(hours=24):
        del reset_tokens[token]
        return render_template("reset_password.html", error="Token has expired")
    
    email = token_data['email']
    
    if request.method == "POST":
        new_password = request.form.get("new_password")
        confirm_password = request.form.get("confirm_password")
        
        if new_password != confirm_password:
            return render_template("reset_password.html", token=token, error="Passwords do not match")
        
        # Update password in both users dict and users_db
        if email in users:
            users[email]['password'] = new_password
        
        for user in users_db:
            if user['email'] == email:
                user['password'] = new_password
                break
        
        del reset_tokens[token]
        flash("Password reset successful. Please login.", "success")
        return redirect(url_for("login"))
    
    return render_template("reset_password.html", token=token)

@app.route("/home")
def home():
    if "user" not in session:
        return redirect(url_for("login"))
    
    students = student_manager.all_students()
    all_interests = []
    for s in students: 
        # Fix: Convert interest value to string before processing
        interest_value = s.get("interest", "")
        if interest_value is None:
            interest_value = ""
        else:
            interest_value = str(interest_value)  # Convert to string
        all_interests.extend(extract_interests(interest_value))
    
    chart_image = generate_chart()
    
    # Prepare students data for template (ensure it's properly formatted)
    students_for_template = []
    for student in students[:5]:
        students_for_template.append({
            'id': student.get('id'),
            'name': student.get('name', ''),
            'email': student.get('email', ''),
            'phone': student.get('phone', ''),
            'education': student.get('education', ''),
            'branch': student.get('branch', ''),
            'year': student.get('year', ''),
            'skills': student.get('skills', ''),
            'interest': str(student.get('interest', ''))  # Convert to string here too
        })
    
    return render_template("home.html", 
                         user=session.get("user_name", "User"),
                         students=students_for_template, 
                         chart=chart_image, 
                         total_students=len(students),
                         unique_interests=len(set(all_interests)))

@app.route("/interns")
def interns():
    if "user" not in session: 
        return redirect(url_for("login"))
    return render_template("interns.html", 
                         students=student_manager.all_students(), 
                         user=session.get("user_name"))


# Add this custom filter to your Flask app
@app.template_filter('string')
def string_filter(value):
    """Convert value to string in templates"""
    if value is None:
        return ""
    return str(value)

@app.route("/new", methods=["GET","POST"])
@app.route("/new-entry", methods=["GET","POST"])
def new_entry():
    if "user" not in session: 
        return redirect(url_for("login"))
    
    if request.method == "POST":
        student = {
            "name": request.form.get("name","").strip(),
            "email": request.form.get("email","").strip(),
            "phone": request.form.get("phone","").strip(),
            "education": request.form.get("education","").strip(),
            "branch": request.form.get("branch","").strip(),
            "year": request.form.get("year","").strip(),
            "skills": request.form.get("skills","").strip(),
            "interest": request.form.get("interest","").strip()
        }
        
        # Check required fields
        required_fields = ["name", "email", "phone", "education", "branch", "year"]
        missing_fields = [field for field in required_fields if not student[field]]
        
        if missing_fields:
            return render_template("new_entry.html", 
                                 error=f"Missing required fields: {', '.join(missing_fields)}")
        
        # Add student to manager (this now automatically saves to Excel)
        student_manager.add_student(student)
        
        flash("Student added successfully!", "success")
        return redirect(url_for("interns"))
    
    return render_template("new_entry.html")

# Change this route definition
@app.route("/edit/<int:student_id>", methods=["GET","POST"])
def edit_student(student_id):
    if "user" not in session:
        return redirect(url_for("login"))
    
    student = student_manager.get_student(student_id)
    if not student:
        flash("Student not found!", "error")
        return redirect(url_for("interns"))
    
    if request.method == "POST":
        updated_data = {
            "name": request.form.get("name","").strip(),
            "email": request.form.get("email","").strip(),
            "phone": request.form.get("phone","").strip(),
            "education": request.form.get("education","").strip(),
            "branch": request.form.get("branch","").strip(),
            "year": request.form.get("year","").strip(),
            "skills": request.form.get("skills","").strip(),
            "interest": request.form.get("interest","").strip()
        }
        
        required_fields = ["name", "email", "phone", "education", "branch", "year"]
        missing_fields = [field for field in required_fields if not updated_data[field]]
        
        if missing_fields:
            return render_template("edit_student.html", 
                                 student=student,
                                 error=f"Missing required fields: {', '.join(missing_fields)}")
        
        student_manager.update_student(student_id, updated_data)
        flash("Student updated successfully!", "success")
        return redirect(url_for("interns"))
    
    return render_template("edit_student.html", student=student)

@app.route("/delete/<int:student_id>", methods=["POST"])
def delete_student(student_id):
    if "user" not in session:
        return redirect(url_for("login"))
    
    student = student_manager.get_student(student_id)
    if student:
        student_manager.delete_student(student_id)
        flash("Student deleted successfully!", "success")
    else:
        flash("Student not found!", "error")
    
    return redirect(url_for("interns"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ----------------- ERROR HANDLERS -----------------
@app.errorhandler(404)
def page_not_found(e):
    return render_template("error.html", error_message="Page not found"), 404

@app.errorhandler(500)
def internal_error(e):
    return render_template("error.html", error_message="Internal server error"), 500

if __name__=="__main__":
    app.run(debug=True, use_debugger=True, use_reloader=True)