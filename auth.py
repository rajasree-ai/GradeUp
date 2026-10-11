from flask import (Blueprint, render_template, request, redirect,
                   url_for, flash)
from flask_login import (LoginManager, UserMixin, login_user,
                         logout_user, login_required, current_user)
from flask_bcrypt import Bcrypt
from database import create_user, get_user_by_email, get_user_by_id


auth_bp = Blueprint("auth", __name__)
bcrypt = Bcrypt()
login_manager = LoginManager()


class User(UserMixin):
    def __init__(self, user_dict):
        self.id = str(user_dict["id"])
        self.full_name = user_dict["full_name"]
        self.email = user_dict["email"]
        self.department = user_dict.get("department", "") or ""


@login_manager.user_loader
def load_user(user_id):
    data = get_user_by_id(int(user_id))
    return User(data) if data else None


def init_auth(app):
    bcrypt.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = "auth.login"
    login_manager.login_message = "Please log in to access this page."
    login_manager.login_message_category = "warning"
    app.register_blueprint(auth_bp)


@auth_bp.route("/register", methods=["GET", "POST"])
def register():
    if current_user.is_authenticated:
        return redirect(url_for("index"))

    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        email = request.form.get("email", "").strip().lower()
        department = request.form.get("department", "").strip()
        password = request.form.get("password", "")
        confirm = request.form.get("confirm_password", "")

        if not (full_name and email and password):
            flash("All required fields must be filled.", "danger")
            return redirect(url_for("auth.register"))
        if len(password) < 6:
            flash("Password must be at least 6 characters.", "danger")
            return redirect(url_for("auth.register"))
        if password != confirm:
            flash("Passwords do not match.", "danger")
            return redirect(url_for("auth.register"))
        if get_user_by_email(email):
            flash("Email already registered. Please log in.", "warning")
            return redirect(url_for("auth.login"))

        password_hash = bcrypt.generate_password_hash(password).decode("utf-8")
        user_id = create_user(full_name, email, department, password_hash)

        if user_id:
            flash("Registration successful! Please log in.", "success")
            return redirect(url_for("auth.login"))
        else:
            flash("Registration failed. Try again.", "danger")

    return render_template("register.html")


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("index"))

    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        remember = bool(request.form.get("remember"))

        user_data = get_user_by_email(email)
        if user_data and bcrypt.check_password_hash(user_data["password_hash"], password):
            login_user(User(user_data), remember=remember)
            flash(f"Welcome back, {user_data['full_name']}!", "success")
            next_page = request.args.get("next")
            if next_page and next_page.startswith("/"):
                return redirect(next_page)
            return redirect(url_for("index"))
        else:
            flash("Invalid email or password.", "danger")

    return render_template("login.html")


@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    flash("You have been logged out.", "success")
    return redirect(url_for("auth.login"))