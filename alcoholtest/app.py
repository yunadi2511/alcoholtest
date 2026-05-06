from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
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

app = Flask(__name__)
app.config['SECRET_KEY'] = 'jkt2-altest-Edge2020!'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///alcoholtest.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = os.path.join('static', 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
DATA_RETENTION_DAYS = 60

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

# ─── MODELS ───────────────────────────────────────────────────────────────────

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(20), nullable=False)  # admin / spv / viewer
    full_name = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    employee_id = db.Column(db.String(50), unique=True, nullable=False)
    discipline = db.Column(db.String(80), nullable=False)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def last_tested_date(self):
        result = TestResult.query.filter_by(employee_id=self.id)\
            .order_by(TestResult.test_date.desc()).first()
        return result.test_date if result else None

    def tested_this_week(self):
        today = date.today()
        week_start = today - timedelta(days=today.weekday())
        count = TestResult.query.filter(
            TestResult.employee_id == self.id,
            TestResult.test_date >= week_start
        ).count()
        return count > 0


class WeeklySchedule(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    week_start = db.Column(db.Date, nullable=False)  # Always a Monday
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
    generated_by = db.Column(db.Integer, db.ForeignKey('user.id'))
    is_swapped = db.Column(db.Boolean, default=False)  # ← ADD THIS
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    employee = db.relationship('Employee', backref='selections')


class TestResult(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    test_date = db.Column(db.Date, nullable=False)
    result = db.Column(db.String(10), nullable=False)  # PASS / FAIL
    remark = db.Column(db.Text)
    evidence_path = db.Column(db.String(300))
    tested_by = db.Column(db.Integer, db.ForeignKey('user.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    employee = db.relationship('Employee', backref='test_results')
    tester = db.relationship('User', backref='conducted_tests')


# ─── HELPERS ──────────────────────────────────────────────────────────────────

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_disciplines():
    disciplines = db.session.query(Employee.discipline).distinct().all()
    return [d[0] for d in disciplines]


def cleanup_old_data():
    cutoff_date = date.today() - timedelta(days=DATA_RETENTION_DAYS)
    
    # Find old test results
    old_results = TestResult.query.filter(TestResult.test_date < cutoff_date).all()
    
    # Delete their photos first
    for result in old_results:
        if result.evidence_path:
            photo_path = os.path.join(app.config['UPLOAD_FOLDER'], result.evidence_path)
            if os.path.exists(photo_path):
                os.remove(photo_path)
    
    # Delete old records
    TestResult.query.filter(TestResult.test_date < cutoff_date).delete()
    DailySelection.query.filter(DailySelection.selection_date < cutoff_date).delete()
    WeeklySchedule.query.filter(WeeklySchedule.week_start < cutoff_date).delete()
    
    db.session.commit()
    print(f"Cleanup done: deleted data older than {cutoff_date}")


def get_todays_employees():
    today = date.today()
    week_start = today - timedelta(days=today.weekday())
    day_map = {0:'mon', 1:'tue', 2:'wed', 3:'thu', 4:'fri', 5:'sat', 6:'sun'}
    today_col = day_map[today.weekday()]

    schedules = WeeklySchedule.query.filter_by(week_start=week_start).all()
    present = []
    for s in schedules:
        if getattr(s, today_col):
            emp = Employee.query.get(s.employee_id)
            if emp and emp.is_active:
                present.append(emp)
    return present


def generate_weekly_selection(week_start, generated_by_id):
    day_map = {0:'mon', 1:'tue', 2:'wed', 3:'thu', 4:'fri', 5:'sat', 6:'sun'}
    schedules = WeeklySchedule.query.filter_by(week_start=week_start).all()

    # Delete any existing selections for this week
    for i in range(7):
        day = week_start + timedelta(days=i)
        DailySelection.query.filter_by(selection_date=day).delete()
    db.session.commit()

    # Generate for each day of the week
    for i in range(7):
        day = week_start + timedelta(days=i)
        day_col = day_map[i]

        # Get employees working this day
        present = []
        for s in schedules:
            if getattr(s, day_col):
                emp = Employee.query.get(s.employee_id)
                if emp and emp.is_active:
                    present.append(emp)

        if not present:
            continue

        disciplines = list(set(e.discipline for e in present))
        not_tested = [e for e in present if not e.tested_this_week()]
        tested = [e for e in present if e.tested_this_week()]

        selected = []
        selected_ids = set()

        # Ensure 1 per discipline
        for disc in disciplines:
            disc_emps = [e for e in not_tested if e.discipline == disc]
            if not disc_emps:
                disc_emps = [e for e in present if e.discipline == disc]
            if disc_emps:
                pick = random.choice(disc_emps)
                if pick.id not in selected_ids:
                    selected.append(pick)
                    selected_ids.add(pick.id)

        # Fill to 15
        pool = [e for e in not_tested if e.id not in selected_ids]
        random.shuffle(pool)
        for e in pool:
            if len(selected) >= 15:
                break
            selected.append(e)
            selected_ids.add(e.id)

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
            return redirect(url_for('dashboard'))
        flash('Invalid username or password.', 'error')
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


# ─── ROUTES: DASHBOARD ────────────────────────────────────────────────────────

@app.route('/dashboard')
@login_required
def dashboard():
    today = date.today()
    # Today's selections
    selections_today = DailySelection.query.filter_by(selection_date=today).all()
    selected_ids = [s.employee_id for s in selections_today]

    # Today's results
    results_today = TestResult.query.filter_by(test_date=today).all()
    tested_ids = [r.employee_id for r in results_today]

    pass_count = sum(1 for r in results_today if r.result == 'PASS')
    fail_count = sum(1 for r in results_today if r.result == 'FAIL')
    pending_count = len(selected_ids) - len(tested_ids)

    # Last 7 days summary
    week_data = []
    for i in range(6, -1, -1):
        d = today - timedelta(days=i)
        day_results = TestResult.query.filter_by(test_date=d).all()
        week_data.append({
            'date': d.strftime('%d %b'),
            'pass': sum(1 for r in day_results if r.result == 'PASS'),
            'fail': sum(1 for r in day_results if r.result == 'FAIL'),
        })

    # Discipline summary today
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
    today = date.today()
    week_start = today - timedelta(days=today.weekday())
    schedules = WeeklySchedule.query.filter_by(week_start=week_start).all()
    employees = Employee.query.filter_by(is_active=True).order_by(Employee.discipline, Employee.name).all()
    selection_exists = DailySelection.query.filter(
        DailySelection.selection_date >= week_start,
        DailySelection.selection_date <= week_start + timedelta(days=6)
    ).first()

    # Build schedule dict for easy lookup in template
    schedule_dict = {}
    for s in schedules:
        schedule_dict[s.employee_id] = s

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

    today = date.today()
    week_start = today - timedelta(days=today.weekday())

    # Check if selection already generated this week
    selection_exists = DailySelection.query.filter(
        DailySelection.selection_date >= week_start,
        DailySelection.selection_date <= week_start + timedelta(days=6)
    ).first()

    if selection_exists:
        flash('Selection already generated this week. Cannot change schedule.', 'error')
        return redirect(url_for('schedule'))

    # Delete existing schedule for this week
    WeeklySchedule.query.filter_by(week_start=week_start).delete()

    employees = Employee.query.filter_by(is_active=True).all()
    days = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']

    for emp in employees:
        working_days = {d: (request.form.get(f'{emp.id}_{d}') == 'on') for d in days}
        if any(working_days.values()):
            sched = WeeklySchedule(
                employee_id=emp.id,
                week_start=week_start,
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

    today = date.today()
    week_start = today - timedelta(days=today.weekday())

    selection_exists = DailySelection.query.filter(
        DailySelection.selection_date >= week_start,
        DailySelection.selection_date <= week_start + timedelta(days=6)
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

        WeeklySchedule.query.filter_by(week_start=week_start).delete()

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

            working_days = {}
            for i, day in enumerate(days):
                val = row[3 + i]
                working_days[day] = (str(val).strip() == '1') if val is not None else False

            if any(working_days.values()):
                sched = WeeklySchedule(
                    employee_id=emp.id,
                    week_start=week_start,
                    created_by=current_user.id,
                    **working_days
                )
                db.session.add(sched)
                count += 1

        db.session.commit()
        flash(f'Schedule uploaded! {count} employees scheduled.', 'success')

    except Exception as e:
        flash(f'Error reading Excel file: {str(e)}', 'error')

    return redirect(url_for('schedule'))



@app.route('/schedule/reset-week', methods=['POST'])
@login_required
def reset_week_schedule():
    if current_user.role != 'admin':
        flash('Access denied. Admin only.', 'error')
        return redirect(url_for('schedule'))
    today = date.today()
    week_start = today - timedelta(days=today.weekday())
    DailySelection.query.filter(
        DailySelection.selection_date >= week_start,
        DailySelection.selection_date <= week_start + timedelta(days=6)
    ).delete()
    WeeklySchedule.query.filter_by(week_start=week_start).delete()
    db.session.commit()
    flash('Week schedule and selections have been reset successfully.', 'success')
    return redirect(url_for('schedule'))


@app.route('/schedule/download-template')
@login_required
def download_template():
    employees = Employee.query.filter_by(is_active=True).order_by(Employee.discipline, Employee.name).all()
    today = date.today()
    week_start = today - timedelta(days=today.weekday())

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Weekly Schedule"

    # Headers
    ws.append(['Employee Name', 'Employee ID', 'Discipline', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'])
    ws['A1'].font = openpyxl.styles.Font(bold=True)

    # Add week dates in row 2
    days_row = ['', '', 'Week of:']
    for i in range(7):
        days_row.append((week_start + timedelta(days=i)).strftime('%d %b'))
    ws.append(days_row)

    # Employee rows
    for emp in employees:
        ws.append([emp.name, emp.employee_id, emp.discipline, '', '', '', '', '', '', ''])

    # Style columns
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 15

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"schedule_template_{week_start}.xlsx"
    return send_file(output, download_name=filename, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ─── ROUTES: GENERATE SELECTION ───────────────────────────────────────────────

@app.route('/generate-selection', methods=['POST'])
@login_required
def generate_selection():
    if current_user.role == 'viewer':
        return jsonify({'error': 'Access denied'}), 403

    today = date.today()
    week_start = today - timedelta(days=today.weekday())

    schedule_exists = WeeklySchedule.query.filter_by(week_start=week_start).first()
    if not schedule_exists:
        flash('Please upload the weekly schedule first.', 'error')
        return redirect(url_for('schedule'))

    existing = DailySelection.query.filter(
        DailySelection.selection_date >= week_start,
        DailySelection.selection_date <= week_start + timedelta(days=6)
    ).first()

    if existing:
        flash('Selection already generated for this week.', 'error')
        return redirect(url_for('schedule'))

    generate_weekly_selection(week_start, current_user.id)
    flash('Weekly selection generated successfully for all 7 days!', 'success')
    return redirect(url_for('schedule'))

# ─── ROUTES: TESTING ──────────────────────────────────────────────────────────

@app.route('/testing')
@login_required
def testing():
    today = date.today()
    selections = DailySelection.query.filter_by(selection_date=today).all()
    results_today = TestResult.query.filter_by(test_date=today).all()
    tested_ids = {r.employee_id: r for r in results_today}
    selection_exists = bool(selections)

    return render_template('testing.html',
        today=today,
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

    employee = Employee.query.get_or_404(employee_id)
    today = date.today()
    existing = TestResult.query.filter_by(employee_id=employee_id, test_date=today).first()
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
                os.makedirs(os.path.dirname(save_path), exist_ok=True)
                file.save(save_path)
                evidence_path = filename

        test = TestResult(
            employee_id=employee_id,
            test_date=today,
            result=result,
            remark=remark,
            evidence_path=evidence_path,
            tested_by=current_user.id
        )
        db.session.add(test)
        db.session.commit()
        flash(f'Test result submitted for {employee.name}.', 'success')
        return redirect(url_for('testing'))

    return render_template('submit_test.html', employee=employee, today=today)


@app.route('/swap-employee/<int:selection_id>', methods=['GET', 'POST'])
@login_required
def swap_employee(selection_id):
    if current_user.role == 'viewer':
        flash('Access denied.', 'error')
        return redirect(url_for('testing'))

    selection = DailySelection.query.get_or_404(selection_id)
    today = date.today()

    # Get available replacements — same discipline, working today, not already selected
    today_employees = get_todays_employees()
    already_selected = [s.employee_id for s in DailySelection.query.filter_by(selection_date=today).all()]
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

        # Log the swap
        swap = SwapLog(
            original_employee_id=selection.employee_id,
            replacement_employee_id=int(replacement_id),
            swap_date=today,
            remark=remark,
            swapped_by=current_user.id
        )
        db.session.add(swap)

        # Update the selection
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
                os.makedirs(os.path.dirname(save_path), exist_ok=True)
                file.save(save_path)
                result.evidence_path = filename
        db.session.commit()
        flash('Test result updated successfully.', 'success')
        return redirect(url_for('history'))
    return render_template('edit_test.html', result=result)


# ─── ROUTES: HISTORY ──────────────────────────────────────────────────────────

@app.route('/history')
@login_required
def history():
    page = request.args.get('page', 1, type=int)
    date_filter = request.args.get('date', '')
    result_filter = request.args.get('result', '')
    disc_filter = request.args.get('discipline', '')

    query = TestResult.query.join(Employee)

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

    results = query.order_by(TestResult.test_date.desc(), TestResult.created_at.desc()).paginate(page=page, per_page=20)
    disciplines = get_disciplines()

    return render_template('history.html',
        results=results,
        disciplines=disciplines,
        date_filter=date_filter,
        result_filter=result_filter,
        disc_filter=disc_filter
    )


@app.route('/export-results')
@login_required
def export_results():
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    query = TestResult.query.join(Employee)

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

    # Create Excel file
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Test Results"

    # Header row
    headers = ['Date', 'Employee Name', 'Employee ID', 'Discipline', 'Result', 'Remark', 'Tested By']
    ws.append(headers)

    # Data rows
    for r in results:
        ws.append([
            r.test_date.strftime('%d/%m/%Y'),
            r.employee.name,
            r.employee.employee_id,
            r.employee.discipline,
            r.result,
            r.remark or '',
            r.tester.full_name if r.tester else ''
        ])

    # Save to memory
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"alcocheck_results_{date_from}_to_{date_to}.xlsx"
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


# ─── ROUTES: EMPLOYEES (ADMIN) ────────────────────────────────────────────────

@app.route('/employees')
@login_required
def employees():
    if current_user.role not in ['admin']:
        flash('Access denied. Admin only.', 'error')
        return redirect(url_for('dashboard'))
    emps = Employee.query.order_by(Employee.discipline, Employee.name).all()
    disciplines = get_disciplines()
    return render_template('employees.html', employees=emps, disciplines=disciplines)


@app.route('/employees/add', methods=['POST'])
@login_required
def add_employee():
    if current_user.role != 'admin':
        return jsonify({'error': 'Access denied'}), 403
    name = request.form.get('name', '').strip()
    emp_id = request.form.get('employee_id', '').strip()
    discipline = request.form.get('discipline', '').strip()

    if not name or not emp_id or not discipline:
        flash('All fields are required.', 'error')
        return redirect(url_for('employees'))

    existing = Employee.query.filter_by(employee_id=emp_id).first()
    if existing:
        flash(f'Employee ID {emp_id} already exists.', 'error')
        return redirect(url_for('employees'))

    emp = Employee(name=name, employee_id=emp_id, discipline=discipline)
    db.session.add(emp)
    db.session.commit()
    flash(f'{name} added successfully.', 'success')
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

    # Example rows so user knows the format
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
            db.session.add(emp)
            added += 1

        db.session.commit()
        flash(f'Import complete! {added} employees added, {skipped} skipped (already exist).', 'success')

    except Exception as e:
        flash(f'Error reading file: {str(e)}', 'error')

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
    db.session.commit()
    flash(f'{emp.name} has been updated successfully.', 'success')
    return redirect(url_for('employees'))


# ─── ROUTES: USER MANAGEMENT (ADMIN) ─────────────────────────────────────────

@app.route('/users')
@login_required
def users():
    if current_user.role != 'admin':
        flash('Access denied.', 'error')
        return redirect(url_for('dashboard'))
    all_users = User.query.order_by(User.role, User.username).all()
    return render_template('users.html', users=all_users)


@app.route('/users/add', methods=['POST'])
@login_required
def add_user():
    if current_user.role != 'admin':
        return jsonify({'error': 'Access denied'}), 403
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '')
    role = request.form.get('role', '')
    full_name = request.form.get('full_name', '').strip()

    if not username or not password or not role:
        flash('All fields are required.', 'error')
        return redirect(url_for('users'))

    if User.query.filter_by(username=username).first():
        flash(f'Username {username} already exists.', 'error')
        return redirect(url_for('users'))

    u = User(username=username, role=role, full_name=full_name)
    u.set_password(password)
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
        user.password = generate_password_hash(new_password)
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

        # No placeholder employees - admin adds real employees via the UI

        db.session.commit()
        print("Database initialized.")


if __name__ == '__main__':
    init_db()
    scheduler = BackgroundScheduler()
    scheduler.add_job(func=cleanup_old_data, trigger='cron', hour=1, minute=0)
    scheduler.start()
    app.run(host='0.0.0.0', port=5000, debug=True)