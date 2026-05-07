from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, session
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from datetime import date, datetime, timedelta
import random
import os
import json
from io import BytesIO
import openpyxl
from apscheduler.schedulers.background import BackgroundScheduler
from PIL import Image

app = Flask(__name__)
app.config['SECRET_KEY'] = 'jkt2-altest-Edge2020!'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///alcoholtest.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = os.path.join('static', 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload
app.config['SITE_NAME'] = os.environ.get('SITE_NAME', 'AlcoCheck')

@app.context_processor
def inject_globals():
    """Inject site_name, active_site, and user's sites into all templates."""
    active_site = None
    user_sites = []

    if current_user.is_authenticated:
        # Admin sees all sites; others see only their assigned sites
        if current_user.role == 'admin':
            user_sites = Site.query.order_by(Site.name).all()
        else:
            user_sites = current_user.sites

        # Get active site from session
        active_site_id = session.get('active_site_id')
        if active_site_id:
            active_site = Site.query.get(active_site_id)
        # If no active site in session, default to first available
        if not active_site and user_sites:
            active_site = user_sites[0]
            session['active_site_id'] = active_site.id

    return dict(
        site_name=app.config['SITE_NAME'],
        active_site=active_site,
        user_sites=user_sites
    )

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
DATA_RETENTION_DAYS = 60

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'


# ─── MODELS ───────────────────────────────────────────────────────────────────

# Many-to-many: User <-> Site
user_sites = db.Table('user_sites',
    db.Column('user_id', db.Integer, db.ForeignKey('user.id'), primary_key=True),
    db.Column('site_id', db.Integer, db.ForeignKey('site.id'), primary_key=True)
)

# Many-to-many: Employee <-> Site
employee_sites = db.Table('employee_sites',
    db.Column('employee_id', db.Integer, db.ForeignKey('employee.id'), primary_key=True),
    db.Column('site_id', db.Integer, db.ForeignKey('site.id'), primary_key=True)
)


class Site(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)  # e.g. JKT1, JKT2
    description = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(20), nullable=False)  # admin / spv / viewer
    full_name = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    sites = db.relationship('Site', secondary=user_sites, backref='users')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def can_access_site(self, site_id):
        """Check if user can access a given site."""
        if self.role == 'admin':
            return True
        return any(s.id == site_id for s in self.sites)


class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    employee_id = db.Column(db.String(50), unique=True, nullable=False)
    discipline = db.Column(db.String(80), nullable=False)
    is_active = db.Column(db.Boolean, default=True)
    is_multisite = db.Column(db.Boolean, default=False)  # Excluded from auto random sampling
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    sites = db.relationship('Site', secondary=employee_sites, backref='employees')

    def last_tested_date(self):
        result = TestResult.query.filter_by(employee_id=self.id)\
            .order_by(TestResult.test_date.desc()).first()
        return result.test_date if result else None

    def tested_this_week(self, site_id=None):
        today = date.today()
        week_start = today - timedelta(days=today.weekday())
        query = TestResult.query.filter(
            TestResult.employee_id == self.id,
            TestResult.test_date >= week_start
        )
        if site_id:
            query = query.filter(TestResult.site_id == site_id)
        return query.count() > 0


class WeeklySchedule(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    week_start = db.Column(db.Date, nullable=False)  # Always a Monday
    site_id = db.Column(db.Integer, db.ForeignKey('site.id'), nullable=False)
    mon = db.Column(db.Boolean, default=False)
    tue = db.Column(db.Boolean, default=False)
    wed = db.Column(db.Boolean, default=False)
    thu = db.Column(db.Boolean, default=False)
    fri = db.Column(db.Boolean, default=False)
    sat = db.Column(db.Boolean, default=False)
    sun = db.Column(db.Boolean, default=False)
    created_by = db.Column(db.Integer, db.ForeignKey('user.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    employee = db.relationship('Employee', backref='schedules')
    site = db.relationship('Site', backref='weekly_schedules')


class SwapLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    original_employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    replacement_employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    swap_date = db.Column(db.Date, nullable=False)
    remark = db.Column(db.Text, nullable=False)
    swapped_by = db.Column(db.Integer, db.ForeignKey('user.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    original = db.relationship('Employee', foreign_keys=[original_employee_id], backref='swapped_out')
    replacement = db.relationship('Employee', foreign_keys=[replacement_employee_id], backref='swapped_in')
    swapper = db.relationship('User', backref='swaps_done')


class DailySelection(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    selection_date = db.Column(db.Date, nullable=False)
    site_id = db.Column(db.Integer, db.ForeignKey('site.id'), nullable=False)
    generated_by = db.Column(db.Integer, db.ForeignKey('user.id'))
    is_swapped = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    employee = db.relationship('Employee', backref='selections')
    site = db.relationship('Site', backref='daily_selections')


class TestResult(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    test_date = db.Column(db.Date, nullable=False)
    site_id = db.Column(db.Integer, db.ForeignKey('site.id'), nullable=False)
    result = db.Column(db.String(10), nullable=False)  # PASS / FAIL
    remark = db.Column(db.Text)
    evidence_path = db.Column(db.String(300))
    tested_by = db.Column(db.Integer, db.ForeignKey('user.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    employee = db.relationship('Employee', backref='test_results')
    tester = db.relationship('User', backref='conducted_tests')
    site = db.relationship('Site', backref='test_results')


# ─── HELPERS ──────────────────────────────────────────────────────────────────

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def compress_image(file, filename, save_path):
    """Compress and save uploaded image. Returns saved filename."""
    try:
        img = Image.open(file)
        # Convert RGBA to RGB if needed (e.g. PNG with transparency)
        if img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')
        # Resize if too large — max 1200px wide
        max_width = 1200
        if img.width > max_width:
            ratio = max_width / img.width
            new_height = int(img.height * ratio)
            img = img.resize((max_width, new_height), Image.LANCZOS)
        # Always save as JPEG for consistency and smaller size
        filename = filename.rsplit('.', 1)[0] + '.jpg'
        save_path = save_path.rsplit('.', 1)[0] + '.jpg'
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        img.save(save_path, 'JPEG', quality=75, optimize=True)
        return filename
    except Exception as e:
        print(f"Image compression error: {e}")
        # Fallback — save original if compression fails
        file.seek(0)
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        file.save(save_path)
        return filename


def get_disciplines(site_id=None):
    if site_id:
        emps = Employee.query.filter(
            Employee.is_active == True,
            Employee.sites.any(Site.id == site_id)
        ).all()
        return list(set(e.discipline for e in emps))
    disciplines = db.session.query(Employee.discipline).distinct().all()
    return [d[0] for d in disciplines]


def get_active_site():
    """Get the currently active site from session."""
    site_id = session.get('active_site_id')
    if site_id:
        return Site.query.get(site_id)
    return None


def cleanup_old_data():
    cutoff_date = date.today() - timedelta(days=DATA_RETENTION_DAYS)
    old_results = TestResult.query.filter(TestResult.test_date < cutoff_date).all()
    for result in old_results:
        if result.evidence_path:
            photo_path = os.path.join(app.config['UPLOAD_FOLDER'], result.evidence_path)
            if os.path.exists(photo_path):
                os.remove(photo_path)
    TestResult.query.filter(TestResult.test_date < cutoff_date).delete()
    DailySelection.query.filter(DailySelection.selection_date < cutoff_date).delete()
    WeeklySchedule.query.filter(WeeklySchedule.week_start < cutoff_date).delete()
    db.session.commit()
    print(f"Cleanup done: deleted data older than {cutoff_date}")


def get_todays_employees(site_id):
    """Get employees scheduled to work today at a given site."""
    today = date.today()
    week_start = today - timedelta(days=today.weekday())
    day_map = {0:'mon', 1:'tue', 2:'wed', 3:'thu', 4:'fri', 5:'sat', 6:'sun'}
    today_col = day_map[today.weekday()]

    schedules = WeeklySchedule.query.filter_by(week_start=week_start, site_id=site_id).all()
    present = []
    for s in schedules:
        if getattr(s, today_col):
            emp = Employee.query.get(s.employee_id)
            if emp and emp.is_active:
                present.append(emp)
    return present


def generate_weekly_selection(week_start, generated_by_id, site_id):
    """Generate weekly random selections for a specific site."""
    day_map = {0:'mon', 1:'tue', 2:'wed', 3:'thu', 4:'fri', 5:'sat', 6:'sun'}
    schedules = WeeklySchedule.query.filter_by(week_start=week_start, site_id=site_id).all()

    # Delete any existing selections for this week + site
    for i in range(7):
        day = week_start + timedelta(days=i)
        DailySelection.query.filter_by(selection_date=day, site_id=site_id).delete()
    db.session.commit()

    for i in range(7):
        day = week_start + timedelta(days=i)
        day_col = day_map[i]

        # Get employees working this day at this site (exclude multi-site employees from auto pool)
        present = []
        for s in schedules:
            if getattr(s, day_col):
                emp = Employee.query.get(s.employee_id)
                if emp and emp.is_active and not emp.is_multisite:
                    present.append(emp)

        if not present:
            continue

        disciplines = list(set(e.discipline for e in present))
        not_tested = [e for e in present if not e.tested_this_week(site_id)]
        tested = [e for e in present if e.tested_this_week(site_id)]

        selected = []
        selected_ids = set()

        # Ensure at least 1 per discipline
        for disc in disciplines:
            disc_emps = [e for e in not_tested if e.discipline == disc]
            if not disc_emps:
                disc_emps = [e for e in present if e.discipline == disc]
            if disc_emps:
                pick = random.choice(disc_emps)
                if pick.id not in selected_ids:
                    selected.append(pick)
                    selected_ids.add(pick.id)

        # Fill to 15 from not-yet-tested pool
        pool = [e for e in not_tested if e.id not in selected_ids]
        random.shuffle(pool)
        for e in pool:
            if len(selected) >= 15:
                break
            selected.append(e)
            selected_ids.add(e.id)

        # If still under 15, fill from already-tested pool
        if len(selected) < 15:
            pool2 = [e for e in tested if e.id not in selected_ids]
            random.shuffle(pool2)
            for e in pool2:
                if len(selected) >= 15:
                    break
                selected.append(e)
                selected_ids.add(e.id)

        # Save selections
        for emp in selected:
            sel = DailySelection(
                employee_id=emp.id,
                selection_date=day,
                site_id=site_id,
                generated_by=generated_by_id,
                is_swapped=False
            )
            db.session.add(sel)

    db.session.commit()


# ─── ROUTES: AUTH ─────────────────────────────────────────────────────────────

@app.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user, remember=True)
            # Set default active site in session
            if user.role == 'admin':
                first_site = Site.query.order_by(Site.name).first()
            else:
                first_site = user.sites[0] if user.sites else None
            if first_site:
                session['active_site_id'] = first_site.id
            return redirect(url_for('dashboard'))
        flash('Invalid username or password.', 'error')
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    session.pop('active_site_id', None)
    logout_user()
    return redirect(url_for('login'))


@app.route('/switch-site/<int:site_id>')
@login_required
def switch_site(site_id):
    """Switch the active site for the current session."""
    if not current_user.can_access_site(site_id):
        flash('You do not have access to that site.', 'error')
        return redirect(url_for('dashboard'))
    site = Site.query.get_or_404(site_id)
    session['active_site_id'] = site.id
    flash(f'Switched to site: {site.name}', 'success')
    # Redirect back to the page they came from
    return redirect(request.referrer or url_for('dashboard'))


# ─── ROUTES: DASHBOARD ────────────────────────────────────────────────────────

@app.route('/dashboard')
@login_required
def dashboard():
    active_site = get_active_site()
    if not active_site:
        flash('No site available. Please ask Admin to create a site and assign you to it.', 'error')
        return render_template('dashboard.html',
            today=date.today(), no_site=True,
            selections_today=[], results_today=[],
            tested_ids=[], pass_count=0, fail_count=0,
            pending_count=0, week_data=json.dumps([]), disc_summary={}
        )

    today = date.today()
    selections_today = DailySelection.query.filter_by(
        selection_date=today, site_id=active_site.id
    ).all()
    selected_ids = [s.employee_id for s in selections_today]

    results_today = TestResult.query.filter_by(
        test_date=today, site_id=active_site.id
    ).all()
    tested_ids = [r.employee_id for r in results_today]

    pass_count = sum(1 for r in results_today if r.result == 'PASS')
    fail_count = sum(1 for r in results_today if r.result == 'FAIL')
    pending_count = len(selected_ids) - len(tested_ids)

    # Last 7 days summary for this site
    week_data = []
    for i in range(6, -1, -1):
        d = today - timedelta(days=i)
        day_results = TestResult.query.filter_by(test_date=d, site_id=active_site.id).all()
        week_data.append({
            'date': d.strftime('%d %b'),
            'pass': sum(1 for r in day_results if r.result == 'PASS'),
            'fail': sum(1 for r in day_results if r.result == 'FAIL'),
        })

    # Discipline summary today for this site
    disc_summary = {}
    for r in results_today:
        disc = r.employee.discipline
        if disc not in disc_summary:
            disc_summary[disc] = {'pass': 0, 'fail': 0}
        if r.result == 'PASS':
            disc_summary[disc]['pass'] += 1
        else:
            disc_summary[disc]['fail'] += 1

    return render_template('dashboard.html',
        today=today,
        no_site=False,
        active_site=active_site,
        selections_today=selections_today,
        results_today=results_today,
        tested_ids=tested_ids,
        pass_count=pass_count,
        fail_count=fail_count,
        pending_count=pending_count,
        week_data=json.dumps(week_data),
        disc_summary=disc_summary
    )


# ─── ROUTES: SCHEDULE ─────────────────────────────────────────────────────────

@app.route('/schedule', methods=['GET'])
@login_required
def schedule():
    if current_user.role == 'viewer':
        flash('Access denied.', 'error')
        return redirect(url_for('dashboard'))

    active_site = get_active_site()
    if not active_site:
        flash('No site selected.', 'error')
        return redirect(url_for('dashboard'))

    today = date.today()
    week_start = today - timedelta(days=today.weekday())

    # Only employees assigned to this site
    employees = Employee.query.filter(
        Employee.is_active == True,
        Employee.sites.any(Site.id == active_site.id)
    ).order_by(Employee.discipline, Employee.name).all()

    schedules = WeeklySchedule.query.filter_by(
        week_start=week_start, site_id=active_site.id
    ).all()

    selection_exists = DailySelection.query.filter(
        DailySelection.selection_date >= week_start,
        DailySelection.selection_date <= week_start + timedelta(days=6),
        DailySelection.site_id == active_site.id
    ).first()

    schedule_dict = {s.employee_id: s for s in schedules}

    by_discipline = {}
    for emp in employees:
        if emp.discipline not in by_discipline:
            by_discipline[emp.discipline] = []
        by_discipline[emp.discipline].append(emp)

    days = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']
    day_labels = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    day_dates = [week_start + timedelta(days=i) for i in range(7)]

    return render_template('schedule.html',
        week_start=week_start,
        active_site=active_site,
        by_discipline=by_discipline,
        schedule_dict=schedule_dict,
        selection_exists=selection_exists,
        days=days,
        day_labels=day_labels,
        day_dates=day_dates,
        today=today
    )


@app.route('/schedule/upload', methods=['POST'])
@login_required
def upload_schedule():
    if current_user.role == 'viewer':
        flash('Access denied.', 'error')
        return redirect(url_for('schedule'))

    active_site = get_active_site()
    if not active_site:
        flash('No site selected.', 'error')
        return redirect(url_for('schedule'))

    today = date.today()
    week_start = today - timedelta(days=today.weekday())

    selection_exists = DailySelection.query.filter(
        DailySelection.selection_date >= week_start,
        DailySelection.selection_date <= week_start + timedelta(days=6),
        DailySelection.site_id == active_site.id
    ).first()

    if selection_exists:
        flash('Selection already generated this week. Cannot change schedule.', 'error')
        return redirect(url_for('schedule'))

    WeeklySchedule.query.filter_by(week_start=week_start, site_id=active_site.id).delete()

    employees = Employee.query.filter(
        Employee.is_active == True,
        Employee.sites.any(Site.id == active_site.id)
    ).all()

    days = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']
    for emp in employees:
        working_days = {d: (request.form.get(f'{emp.id}_{d}') == 'on') for d in days}
        if any(working_days.values()):
            sched = WeeklySchedule(
                employee_id=emp.id,
                week_start=week_start,
                site_id=active_site.id,
                created_by=current_user.id,
                **working_days
            )
            db.session.add(sched)

    db.session.commit()
    flash('Weekly schedule saved successfully!', 'success')
    return redirect(url_for('schedule'))


@app.route('/schedule/upload-excel', methods=['POST'])
@login_required
def upload_schedule_excel():
    if current_user.role == 'viewer':
        flash('Access denied.', 'error')
        return redirect(url_for('schedule'))

    active_site = get_active_site()
    if not active_site:
        flash('No site selected.', 'error')
        return redirect(url_for('schedule'))

    today = date.today()
    week_start = today - timedelta(days=today.weekday())

    selection_exists = DailySelection.query.filter(
        DailySelection.selection_date >= week_start,
        DailySelection.selection_date <= week_start + timedelta(days=6),
        DailySelection.site_id == active_site.id
    ).first()
    if selection_exists:
        flash('Selection already generated this week. Cannot change schedule.', 'error')
        return redirect(url_for('schedule'))

    if 'excel_file' not in request.files:
        flash('No file selected.', 'error')
        return redirect(url_for('schedule'))

    file = request.files['excel_file']
    if file.filename == '':
        flash('No file selected.', 'error')
        return redirect(url_for('schedule'))

    if not file.filename.endswith('.xlsx'):
        flash('Please upload an .xlsx file only.', 'error')
        return redirect(url_for('schedule'))

    try:
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        days = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']

        WeeklySchedule.query.filter_by(week_start=week_start, site_id=active_site.id).delete()

        count = 0
        for row in ws.iter_rows(min_row=3, values_only=True):
            emp_name = row[0]
            emp_id_str = str(row[1]) if row[1] else ''
            if not emp_name:
                continue

            emp = Employee.query.filter(
                (Employee.employee_id == emp_id_str) |
                (Employee.name == emp_name)
            ).filter_by(is_active=True).first()

            if not emp:
                continue

            # Only process if employee is assigned to this site
            if not any(s.id == active_site.id for s in emp.sites):
                continue

            working_days = {}
            for i, day in enumerate(days):
                val = row[3 + i]
                working_days[day] = (str(val).strip() == '1') if val is not None else False

            if any(working_days.values()):
                sched = WeeklySchedule(
                    employee_id=emp.id,
                    week_start=week_start,
                    site_id=active_site.id,
                    created_by=current_user.id,
                    **working_days
                )
                db.session.add(sched)
                count += 1

        db.session.commit()
        flash(f'Schedule uploaded! {count} employees scheduled for {active_site.name}.', 'success')

    except Exception as e:
        flash(f'Error reading Excel file: {str(e)}', 'error')

    return redirect(url_for('schedule'))


@app.route('/schedule/reset-week', methods=['POST'])
@login_required
def reset_week_schedule():
    if current_user.role != 'admin':
        flash('Access denied. Admin only.', 'error')
        return redirect(url_for('schedule'))

    active_site = get_active_site()
    if not active_site:
        flash('No site selected.', 'error')
        return redirect(url_for('schedule'))

    today = date.today()
    week_start = today - timedelta(days=today.weekday())

    DailySelection.query.filter(
        DailySelection.selection_date >= week_start,
        DailySelection.selection_date <= week_start + timedelta(days=6),
        DailySelection.site_id == active_site.id
    ).delete()
    WeeklySchedule.query.filter_by(week_start=week_start, site_id=active_site.id).delete()
    db.session.commit()
    flash(f'Week schedule and selections for {active_site.name} have been reset.', 'success')
    return redirect(url_for('schedule'))


@app.route('/schedule/download-template')
@login_required
def download_template():
    active_site = get_active_site()
    if not active_site:
        flash('No site selected.', 'error')
        return redirect(url_for('schedule'))

    employees = Employee.query.filter(
        Employee.is_active == True,
        Employee.sites.any(Site.id == active_site.id)
    ).order_by(Employee.discipline, Employee.name).all()

    today = date.today()
    week_start = today - timedelta(days=today.weekday())

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Weekly Schedule"

    ws.append(['Employee Name', 'Employee ID', 'Discipline', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'])
    ws['A1'].font = openpyxl.styles.Font(bold=True)

    days_row = ['', '', 'Week of:']
    for i in range(7):
        days_row.append((week_start + timedelta(days=i)).strftime('%d %b'))
    ws.append(days_row)

    for emp in employees:
        ws.append([emp.name, emp.employee_id, emp.discipline, '', '', '', '', '', '', ''])

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 15

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"schedule_template_{active_site.name}_{week_start}.xlsx"
    return send_file(output, download_name=filename, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ─── ROUTES: GENERATE SELECTION ───────────────────────────────────────────────

@app.route('/generate-selection', methods=['POST'])
@login_required
def generate_selection():
    if current_user.role == 'viewer':
        return jsonify({'error': 'Access denied'}), 403

    active_site = get_active_site()
    if not active_site:
        flash('No site selected.', 'error')
        return redirect(url_for('schedule'))

    today = date.today()
    week_start = today - timedelta(days=today.weekday())

    schedule_exists = WeeklySchedule.query.filter_by(
        week_start=week_start, site_id=active_site.id
    ).first()
    if not schedule_exists:
        flash('Please upload the weekly schedule first.', 'error')
        return redirect(url_for('schedule'))

    existing = DailySelection.query.filter(
        DailySelection.selection_date >= week_start,
        DailySelection.selection_date <= week_start + timedelta(days=6),
        DailySelection.site_id == active_site.id
    ).first()

    if existing:
        flash('Selection already generated for this week.', 'error')
        return redirect(url_for('schedule'))

    generate_weekly_selection(week_start, current_user.id, active_site.id)
    flash(f'Weekly selection generated successfully for {active_site.name}!', 'success')
    return redirect(url_for('schedule'))


# ─── ROUTES: TESTING ──────────────────────────────────────────────────────────

@app.route('/testing')
@login_required
def testing():
    active_site = get_active_site()
    if not active_site:
        flash('No site selected.', 'error')
        return redirect(url_for('dashboard'))

    today = date.today()
    selections = DailySelection.query.filter_by(
        selection_date=today, site_id=active_site.id
    ).all()
    results_today = TestResult.query.filter_by(
        test_date=today, site_id=active_site.id
    ).all()
    tested_ids = {r.employee_id: r for r in results_today}
    selection_exists = bool(selections)

    return render_template('testing.html',
        today=today,
        active_site=active_site,
        selections=selections,
        tested_ids=tested_ids,
        selection_exists=selection_exists
    )


@app.route('/submit-test/<int:employee_id>', methods=['GET', 'POST'])
@login_required
def submit_test(employee_id):
    if current_user.role == 'viewer':
        flash('Access denied.', 'error')
        return redirect(url_for('testing'))

    active_site = get_active_site()
    if not active_site:
        flash('No site selected.', 'error')
        return redirect(url_for('testing'))

    employee = Employee.query.get_or_404(employee_id)
    today = date.today()
    existing = TestResult.query.filter_by(
        employee_id=employee_id, test_date=today, site_id=active_site.id
    ).first()
    if existing:
        flash('Test already submitted for this employee today.', 'error')
        return redirect(url_for('testing'))

    if request.method == 'POST':
        result = request.form.get('result')
        remark = request.form.get('remark', '')
        evidence_path = None

        if 'evidence' in request.files:
            file = request.files['evidence']
            if file and file.filename and allowed_file(file.filename):
                filename = secure_filename(f"{today}_{employee_id}_{file.filename}")
                save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                evidence_path = compress_image(file, filename, save_path)

        test = TestResult(
            employee_id=employee_id,
            test_date=today,
            site_id=active_site.id,
            result=result,
            remark=remark,
            evidence_path=evidence_path,
            tested_by=current_user.id
        )
        db.session.add(test)
        db.session.commit()
        flash(f'Test result submitted for {employee.name}.', 'success')
        return redirect(url_for('testing'))

    return render_template('submit_test.html',
        employee=employee, today=today, active_site=active_site
    )


@app.route('/swap-employee/<int:selection_id>', methods=['GET', 'POST'])
@login_required
def swap_employee(selection_id):
    if current_user.role == 'viewer':
        flash('Access denied.', 'error')
        return redirect(url_for('testing'))

    active_site = get_active_site()
    if not active_site:
        flash('No site selected.', 'error')
        return redirect(url_for('testing'))

    selection = DailySelection.query.get_or_404(selection_id)
    today = date.today()

    today_employees = get_todays_employees(active_site.id)
    already_selected = [s.employee_id for s in DailySelection.query.filter_by(
        selection_date=today, site_id=active_site.id
    ).all()]
    available = [e for e in today_employees
                 if e.discipline == selection.employee.discipline
                 and e.id not in already_selected
                 and e.id != selection.employee_id]

    if request.method == 'POST':
        replacement_id = request.form.get('replacement_id')
        remark = request.form.get('remark', '').strip()

        if not replacement_id or not remark:
            flash('Please select a replacement and provide a reason.', 'error')
            return render_template('swap.html', selection=selection, available=available, today=today)

        swap = SwapLog(
            original_employee_id=selection.employee_id,
            replacement_employee_id=int(replacement_id),
            swap_date=today,
            remark=remark,
            swapped_by=current_user.id
        )
        db.session.add(swap)
        selection.employee_id = int(replacement_id)
        selection.is_swapped = True
        db.session.commit()

        flash('Employee swapped successfully.', 'success')
        return redirect(url_for('testing'))

    return render_template('swap.html', selection=selection, available=available, today=today)


@app.route('/edit-test/<int:result_id>', methods=['GET', 'POST'])
@login_required
def edit_test(result_id):
    if current_user.role not in ['admin', 'spv']:
        flash('Access denied.', 'error')
        return redirect(url_for('history'))
    result = TestResult.query.get_or_404(result_id)
    if request.method == 'POST':
        result.result = request.form.get('result', result.result)
        result.remark = request.form.get('remark', result.remark)
        if 'evidence' in request.files:
            file = request.files['evidence']
            if file and file.filename and allowed_file(file.filename):
                filename = secure_filename(f"{result.test_date}_{result.employee_id}_{file.filename}")
                save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                result.evidence_path = compress_image(file, filename, save_path)
        db.session.commit()
        flash('Test result updated successfully.', 'success')
        return redirect(url_for('history'))
    return render_template('edit_test.html', result=result)


# ─── ROUTES: HISTORY ──────────────────────────────────────────────────────────

@app.route('/history')
@login_required
def history():
    active_site = get_active_site()
    if not active_site:
        flash('No site selected.', 'error')
        return redirect(url_for('dashboard'))

    page = request.args.get('page', 1, type=int)
    date_filter = request.args.get('date', '')
    result_filter = request.args.get('result', '')
    disc_filter = request.args.get('discipline', '')

    query = TestResult.query.join(Employee).filter(
        TestResult.site_id == active_site.id
    )

    if date_filter:
        try:
            filter_date = datetime.strptime(date_filter, '%Y-%m-%d').date()
            query = query.filter(TestResult.test_date == filter_date)
        except:
            pass
    if result_filter:
        query = query.filter(TestResult.result == result_filter)
    if disc_filter:
        query = query.filter(Employee.discipline == disc_filter)

    results = query.order_by(
        TestResult.test_date.desc(), TestResult.created_at.desc()
    ).paginate(page=page, per_page=20)

    disciplines = get_disciplines(active_site.id)

    return render_template('history.html',
        results=results,
        active_site=active_site,
        disciplines=disciplines,
        date_filter=date_filter,
        result_filter=result_filter,
        disc_filter=disc_filter
    )


@app.route('/export-results')
@login_required
def export_results():
    active_site = get_active_site()
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    query = TestResult.query.join(Employee)
    if active_site:
        query = query.filter(TestResult.site_id == active_site.id)

    if date_from:
        try:
            query = query.filter(TestResult.test_date >= datetime.strptime(date_from, '%Y-%m-%d').date())
        except:
            pass
    if date_to:
        try:
            query = query.filter(TestResult.test_date <= datetime.strptime(date_to, '%Y-%m-%d').date())
        except:
            pass

    results = query.order_by(TestResult.test_date.desc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Test Results"

    headers = ['Date', 'Site', 'Employee Name', 'Employee ID', 'Discipline', 'Result', 'Remark', 'Tested By']
    ws.append(headers)

    for r in results:
        ws.append([
            r.test_date.strftime('%d/%m/%Y'),
            r.site.name if r.site else '',
            r.employee.name,
            r.employee.employee_id,
            r.employee.discipline,
            r.result,
            r.remark or '',
            r.tester.full_name if r.tester else ''
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    site_label = active_site.name if active_site else 'all'
    filename = f"alcocheck_results_{site_label}_{date_from}_to_{date_to}.xlsx"
    return send_file(output, download_name=filename, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/run-cleanup', methods=['POST'])
@login_required
def run_cleanup():
    if current_user.role != 'admin':
        return jsonify({'error': 'Access denied'}), 403
    with app.app_context():
        cleanup_old_data()
    flash('Cleanup completed. Data older than 60 days has been deleted.', 'success')
    return redirect(url_for('dashboard'))


# ─── ROUTES: SITES (ADMIN) ────────────────────────────────────────────────────

@app.route('/sites')
@login_required
def sites():
    if current_user.role != 'admin':
        flash('Access denied. Admin only.', 'error')
        return redirect(url_for('dashboard'))
    all_sites = Site.query.order_by(Site.name).all()
    return render_template('sites.html', sites=all_sites)


@app.route('/sites/add', methods=['POST'])
@login_required
def add_site():
    if current_user.role != 'admin':
        return jsonify({'error': 'Access denied'}), 403

    name = request.form.get('name', '').strip().upper()
    description = request.form.get('description', '').strip()

    if not name:
        flash('Site name is required.', 'error')
        return redirect(url_for('sites'))

    if Site.query.filter_by(name=name).first():
        flash(f'Site {name} already exists.', 'error')
        return redirect(url_for('sites'))

    site = Site(name=name, description=description)
    db.session.add(site)
    db.session.commit()
    flash(f'Site {name} created successfully.', 'success')
    return redirect(url_for('sites'))


@app.route('/sites/delete/<int:site_id>', methods=['POST'])
@login_required
def delete_site(site_id):
    if current_user.role != 'admin':
        return jsonify({'error': 'Access denied'}), 403

    site = Site.query.get_or_404(site_id)

    # Check if site has any data
    has_data = (
        WeeklySchedule.query.filter_by(site_id=site_id).first() or
        DailySelection.query.filter_by(site_id=site_id).first() or
        TestResult.query.filter_by(site_id=site_id).first()
    )
    if has_data:
        flash(f'Cannot delete {site.name} — it has existing schedule or test data.', 'error')
        return redirect(url_for('sites'))

    db.session.delete(site)
    db.session.commit()
    flash(f'Site {site.name} deleted.', 'success')
    return redirect(url_for('sites'))


# ─── ROUTES: EMPLOYEES (ADMIN) ────────────────────────────────────────────────

@app.route('/employees')
@login_required
def employees():
    if current_user.role != 'admin':
        flash('Access denied. Admin only.', 'error')
        return redirect(url_for('dashboard'))
    emps = Employee.query.order_by(Employee.discipline, Employee.name).all()
    disciplines = get_disciplines()
    all_sites = Site.query.order_by(Site.name).all()
    return render_template('employees.html',
        employees=emps, disciplines=disciplines, all_sites=all_sites
    )


@app.route('/employees/add', methods=['POST'])
@login_required
def add_employee():
    if current_user.role != 'admin':
        return jsonify({'error': 'Access denied'}), 403

    name = request.form.get('name', '').strip()
    emp_id = request.form.get('employee_id', '').strip()
    discipline = request.form.get('discipline', '').strip()
    site_ids = request.form.getlist('site_ids')  # Multiple site checkboxes

    if not name or not emp_id or not discipline:
        flash('Name, Employee ID, and Discipline are required.', 'error')
        return redirect(url_for('employees'))

    if Employee.query.filter_by(employee_id=emp_id).first():
        flash(f'Employee ID {emp_id} already exists.', 'error')
        return redirect(url_for('employees'))

    emp = Employee(name=name, employee_id=emp_id, discipline=discipline)

    # Assign sites
    selected_sites = Site.query.filter(Site.id.in_(site_ids)).all()
    emp.sites = selected_sites
    emp.is_multisite = len(selected_sites) > 1

    db.session.add(emp)
    db.session.commit()
    flash(f'{name} added successfully.', 'success')
    return redirect(url_for('employees'))


@app.route('/employees/edit/<int:emp_id>', methods=['POST'])
@login_required
def edit_employee(emp_id):
    if current_user.role != 'admin':
        return jsonify({'error': 'Access denied'}), 403

    emp = Employee.query.get_or_404(emp_id)
    emp.name = request.form.get('name', emp.name).strip()
    emp.employee_id = request.form.get('employee_id', emp.employee_id).strip()
    emp.discipline = request.form.get('discipline', emp.discipline).strip()

    site_ids = request.form.getlist('site_ids')
    selected_sites = Site.query.filter(Site.id.in_(site_ids)).all()
    emp.sites = selected_sites
    emp.is_multisite = len(selected_sites) > 1

    db.session.commit()
    flash(f'{emp.name} has been updated successfully.', 'success')
    return redirect(url_for('employees'))


@app.route('/employees/delete/<int:emp_id>', methods=['POST'])
@login_required
def delete_employee(emp_id):
    if current_user.role != 'admin':
        return jsonify({'error': 'Access denied'}), 403
    emp = Employee.query.get_or_404(emp_id)
    name = emp.name
    DailySelection.query.filter_by(employee_id=emp_id).delete()
    TestResult.query.filter_by(employee_id=emp_id).delete()
    db.session.delete(emp)
    db.session.commit()
    flash(f'{name} has been permanently deleted.', 'success')
    return redirect(url_for('employees'))


@app.route('/employees/toggle/<int:emp_id>', methods=['POST'])
@login_required
def toggle_employee(emp_id):
    if current_user.role != 'admin':
        return jsonify({'error': 'Access denied'}), 403
    emp = Employee.query.get_or_404(emp_id)
    emp.is_active = not emp.is_active
    db.session.commit()
    status = 'activated' if emp.is_active else 'deactivated'
    flash(f'{emp.name} has been {status}.', 'success')
    return redirect(url_for('employees'))


@app.route('/employees/download-template')
@login_required
def download_employee_template():
    if current_user.role != 'admin':
        return redirect(url_for('dashboard'))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Employees"

    headers = ['Name', 'Employee ID', 'Discipline']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True)

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 20

    ws.append(['Ahmad Yusuf', 'EMP001', 'Mechanical'])
    ws.append(['Budi Santoso', 'EMP002', 'Electrical'])
    ws.append(['Dewi Rahayu', 'EMP003', 'Operator'])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, download_name='employee_template.xlsx',
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/employees/import', methods=['POST'])
@login_required
def import_employees():
    if current_user.role != 'admin':
        flash('Access denied.', 'error')
        return redirect(url_for('employees'))

    if 'employee_file' not in request.files:
        flash('No file selected.', 'error')
        return redirect(url_for('employees'))

    file = request.files['employee_file']
    if file.filename == '':
        flash('No file selected.', 'error')
        return redirect(url_for('employees'))

    if not (file.filename.endswith('.xlsx') or file.filename.endswith('.csv')):
        flash('Please upload .xlsx or .csv file only.', 'error')
        return redirect(url_for('employees'))

    # Get site to assign imported employees to (optional)
    import_site_id = request.form.get('import_site_id')
    import_site = Site.query.get(import_site_id) if import_site_id else None

    try:
        added = 0
        skipped = 0

        if file.filename.endswith('.csv'):
            import csv, io
            stream = io.StringIO(file.stream.read().decode('utf-8'))
            reader = csv.DictReader(stream)
            rows = [(r.get('Name','').strip(), r.get('Employee ID','').strip(), r.get('Discipline','').strip()) for r in reader]
        else:
            wb = openpyxl.load_workbook(file)
            ws = wb.active
            rows = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                name = str(row[0]).strip() if row[0] else ''
                emp_id = str(row[1]).strip() if row[1] else ''
                discipline = str(row[2]).strip() if row[2] else ''
                rows.append((name, emp_id, discipline))

        for name, emp_id, discipline in rows:
            if not name or not emp_id:
                continue
            existing = Employee.query.filter_by(employee_id=emp_id).first()
            if existing:
                skipped += 1
                continue
            emp = Employee(
                name=name,
                employee_id=emp_id,
                discipline=discipline or 'General',
                is_active=True
            )
            if import_site:
                emp.sites = [import_site]
            db.session.add(emp)
            added += 1

        db.session.commit()
        flash(f'Import complete! {added} employees added, {skipped} skipped (already exist).', 'success')

    except Exception as e:
        flash(f'Error reading file: {str(e)}', 'error')

    return redirect(url_for('employees'))


# ─── ROUTES: USER MANAGEMENT (ADMIN) ─────────────────────────────────────────

@app.route('/users')
@login_required
def users():
    if current_user.role != 'admin':
        flash('Access denied.', 'error')
        return redirect(url_for('dashboard'))
    all_users = User.query.order_by(User.role, User.username).all()
    all_sites = Site.query.order_by(Site.name).all()
    return render_template('users.html', users=all_users, all_sites=all_sites)


@app.route('/users/add', methods=['POST'])
@login_required
def add_user():
    if current_user.role != 'admin':
        return jsonify({'error': 'Access denied'}), 403

    username = request.form.get('username', '').strip()
    password = request.form.get('password', '')
    role = request.form.get('role', '')
    full_name = request.form.get('full_name', '').strip()
    site_ids = request.form.getlist('site_ids')

    if not username or not password or not role:
        flash('Username, password, and role are required.', 'error')
        return redirect(url_for('users'))

    if User.query.filter_by(username=username).first():
        flash(f'Username {username} already exists.', 'error')
        return redirect(url_for('users'))

    u = User(username=username, role=role, full_name=full_name)
    u.set_password(password)

    # Admin automatically gets all sites; others get selected sites
    if role == 'admin':
        u.sites = Site.query.all()
    else:
        u.sites = Site.query.filter(Site.id.in_(site_ids)).all()

    db.session.add(u)
    db.session.commit()
    flash(f'User {username} created successfully.', 'success')
    return redirect(url_for('users'))


@app.route('/users/edit/<int:user_id>', methods=['POST'])
@login_required
def edit_user(user_id):
    if current_user.role != 'admin':
        return redirect(url_for('dashboard'))

    user = User.query.get_or_404(user_id)
    user.full_name = request.form.get('full_name', user.full_name)
    user.role = request.form.get('role', user.role)

    new_password = request.form.get('password', '').strip()
    if new_password:
        user.set_password(new_password)

    site_ids = request.form.getlist('site_ids')
    if user.role == 'admin':
        user.sites = Site.query.all()
    else:
        user.sites = Site.query.filter(Site.id.in_(site_ids)).all()

    db.session.commit()
    flash('User updated successfully.', 'success')
    return redirect(url_for('users'))


@app.route('/users/delete/<int:user_id>', methods=['POST'])
@login_required
def delete_user(user_id):
    if current_user.role != 'admin':
        return jsonify({'error': 'Access denied'}), 403
    if user_id == current_user.id:
        flash('Cannot delete your own account.', 'error')
        return redirect(url_for('users'))
    u = User.query.get_or_404(user_id)
    db.session.delete(u)
    db.session.commit()
    flash(f'User {u.username} deleted.', 'success')
    return redirect(url_for('users'))


# ─── INIT DB ──────────────────────────────────────────────────────────────────

def init_db():
    with app.app_context():
        db.create_all()

        # Create default admin if not exists
        if not User.query.filter_by(username='admin').first():
            admin = User(username='admin', role='admin', full_name='Administrator')
            admin.set_password('Admin@1234')
            db.session.add(admin)
            db.session.commit()
            print("Default admin created. Username: admin / Password: Admin@1234")

        db.session.commit()
        print("Database initialized.")


if __name__ == '__main__':
    init_db()
    scheduler = BackgroundScheduler()
    scheduler.add_job(func=cleanup_old_data, trigger='cron', hour=1, minute=0)
    scheduler.start()
    app.run(host='0.0.0.0', port=5000, debug=True)