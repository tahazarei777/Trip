import sqlite3
import io
import os
from flask import Flask, render_template_string, request, redirect, url_for, flash, send_file
import xlsxwriter

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-me-1234'

DATABASE = os.path.join(os.path.dirname(__file__), 'expenses.db')

# ================== DATABASE ==================
def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    with get_db() as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS families (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                leader_id INTEGER,
                FOREIGN KEY (leader_id) REFERENCES persons(id) ON DELETE SET NULL
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS persons (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                family_id INTEGER NOT NULL,
                FOREIGN KEY (family_id) REFERENCES families(id) ON DELETE CASCADE
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                family_id INTEGER NOT NULL,
                person_id INTEGER NOT NULL,
                description TEXT NOT NULL,
                price INTEGER NOT NULL,
                date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (family_id) REFERENCES families(id) ON DELETE CASCADE,
                FOREIGN KEY (person_id) REFERENCES persons(id) ON DELETE CASCADE
            )
        ''')
        conn.commit()

# بلافاصله جداول را بساز (برای WSGI ضروری است)
init_db()

# ================== UI/UX CSS ==================
BASE_CSS = '''
@import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;500;700&display=swap');

* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}

body {
    font-family: 'Vazirmatn', Tahoma, sans-serif;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    padding: 20px;
    direction: rtl;
    color: #333;
}

.container {
    max-width: 1000px;
    margin: 0 auto;
}

/* ----- TOP BAR (hamburger on mobile) ----- */
.top-bar {
    display: none;
    justify-content: space-between;
    align-items: center;
    background: rgba(255,255,255,0.9);
    backdrop-filter: blur(10px);
    border-radius: 12px;
    padding: 10px 20px;
    margin-bottom: 15px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
}

.top-bar .brand {
    font-size: 18px;
    font-weight: 700;
    color: #4a3f6b;
}

.hamburger {
    background: none;
    border: none;
    font-size: 26px;
    cursor: pointer;
    color: #4a3f6b;
    padding: 5px;
}

/* ----- NAVBAR ----- */
.navbar {
    background: rgba(255, 255, 255, 0.2);
    backdrop-filter: blur(15px);
    border-radius: 16px;
    padding: 12px 20px;
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    justify-content: center;
    margin-bottom: 25px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.1);
}

.navbar a {
    color: white;
    text-decoration: none;
    padding: 10px 18px;
    border-radius: 30px;
    font-weight: 500;
    transition: 0.3s;
    background: rgba(255,255,255,0.15);
    display: flex;
    align-items: center;
    gap: 6px;
}

.navbar a:hover {
    background: rgba(255,255,255,0.35);
    transform: translateY(-2px);
}

.navbar .btn-excel {
    background: #27ae60;
    margin-right: auto;
}

.navbar .btn-excel:hover {
    background: #2ecc71;
}

/* Mobile menu toggle */
@media (max-width: 768px) {
    .top-bar {
        display: flex;
    }

    .navbar {
        display: none;
        flex-direction: column;
        align-items: stretch;
        background: rgba(255,255,255,0.95);
        backdrop-filter: blur(20px);
        box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        border-radius: 20px;
        padding: 20px;
        margin-top: 0;
    }

    .navbar.show {
        display: flex;
    }

    .navbar a {
        color: #333;
        background: #f0f0f0;
        justify-content: center;
    }

    .navbar a:hover {
        background: #e0e0e0;
    }

    .navbar .btn-excel {
        margin-right: 0;
        background: #27ae60;
        color: white;
    }
}

/* ----- CARD ----- */
.card {
    background: white;
    border-radius: 20px;
    padding: 30px;
    box-shadow: 0 20px 50px rgba(0,0,0,0.15);
}

.card h2 {
    color: #4a3f6b;
    margin-bottom: 20px;
    font-size: 24px;
    display: flex;
    align-items: center;
    gap: 8px;
}

/* ----- ALERTS ----- */
.alert {
    padding: 14px 18px;
    border-radius: 12px;
    margin-bottom: 20px;
    font-size: 15px;
    animation: slideDown 0.4s ease-out;
    transition: opacity 0.5s;
}

.alert-error {
    background: #ffe0e0;
    color: #c0392b;
    border-left: 5px solid #e74c3c;
}

.alert-success {
    background: #e0ffe0;
    color: #27ae60;
    border-left: 5px solid #2ecc71;
}

@keyframes slideDown {
    from { opacity: 0; transform: translateY(-20px); }
    to { opacity: 1; transform: translateY(0); }
}

/* ----- FORMS ----- */
.form-group {
    margin-bottom: 20px;
}

label {
    display: block;
    margin-bottom: 6px;
    font-weight: 500;
    color: #555;
}

input, select {
    width: 100%;
    padding: 12px 15px;
    border: 2px solid #e0e0e0;
    border-radius: 12px;
    font-family: inherit;
    font-size: 15px;
    transition: 0.3s;
    background: #fafafa;
}

input:focus, select:focus {
    outline: none;
    border-color: #667eea;
    background: white;
    box-shadow: 0 0 0 3px rgba(102,126,234,0.2);
}

.inline-fields {
    display: flex;
    gap: 15px;
    flex-wrap: wrap;
}

.inline-fields > * {
    flex: 1;
}

button, .btn {
    background: #667eea;
    color: white;
    border: none;
    padding: 12px 24px;
    border-radius: 30px;
    font-family: inherit;
    font-size: 15px;
    font-weight: 500;
    cursor: pointer;
    transition: 0.3s;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    text-decoration: none;
    justify-content: center;
}

button:hover, .btn:hover {
    background: #5a6fd6;
    transform: translateY(-2px);
    box-shadow: 0 5px 15px rgba(102,126,234,0.4);
}

.btn-danger {
    background: #e74c3c;
}

.btn-danger:hover {
    background: #c0392b;
    box-shadow: 0 5px 15px rgba(231,76,60,0.4);
}

.btn-success {
    background: #2ecc71;
}

.btn-success:hover {
    background: #27ae60;
    box-shadow: 0 5px 15px rgba(46,204,113,0.4);
}

.btn-warning {
    background: #f39c12;
}

.btn-warning:hover {
    background: #e67e22;
    box-shadow: 0 5px 15px rgba(243,156,18,0.4);
}

.btn-outline-danger {
    background: transparent;
    color: #e74c3c;
    border: 2px solid #e74c3c;
}

.btn-outline-danger:hover {
    background: #e74c3c;
    color: white;
}

/* ----- RESPONSIVE TABLES ----- */
.table-wrapper {
    margin: 20px 0;
    border-radius: 12px;
}

table {
    width: 100%;
    border-collapse: collapse;
    box-shadow: 0 5px 15px rgba(0,0,0,0.05);
}

th {
    background: #667eea;
    color: white;
    font-weight: 500;
    padding: 14px;
    text-align: center;
}

td {
    padding: 12px;
    text-align: center;
    border-bottom: 1px solid #f0f0f0;
}

tr:nth-child(even) {
    background: #f9f9ff;
}

tr:hover td {
    background: #eef0ff;
    transition: 0.2s;
}

@media (max-width: 600px) {
    .table-wrapper {
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
    }

    table {
        min-width: 700px;
    }

    .card {
        padding: 20px;
    }
}

/* ----- COLLAPSIBLE FAMILY CARDS ----- */
.family-details {
    background: #f4f6ff;
    border-radius: 16px;
    margin-bottom: 20px;
    overflow: hidden;
    border: 1px solid #e0e0f0;
}

.family-details summary {
    list-style: none;
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 18px 20px;
    background: white;
    cursor: pointer;
    font-weight: 500;
    font-size: 18px;
    color: #4a3f6b;
    border-radius: 16px;
    transition: 0.2s;
}

.family-details summary::-webkit-details-marker {
    display: none;
}

.family-details summary::after {
    content: '▼';
    font-size: 14px;
    transition: 0.3s;
    color: #667eea;
}

.family-details[open] summary::after {
    transform: rotate(180deg);
}

.family-details summary:hover {
    background: #f9f9ff;
}

.family-content {
    padding: 20px;
}

.members-list {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin: 10px 0;
}

.member-badge {
    background: white;
    padding: 6px 15px;
    border-radius: 20px;
    font-size: 14px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    display: flex;
    align-items: center;
    gap: 8px;
}

.inline-form {
    display: flex;
    gap: 10px;
    margin-top: 15px;
    align-items: center;
}

.inline-form input {
    flex: 1;
}

.stats-bar {
    display: flex;
    gap: 20px;
    margin-bottom: 20px;
    background: #f0f0f8;
    padding: 15px;
    border-radius: 12px;
    justify-content: center;
    font-weight: 500;
}

/* ----- REPORT CARDS ----- */
.report-cards {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 20px;
    margin: 25px 0;
}

.stat-card {
    background: white;
    padding: 20px;
    border-radius: 16px;
    text-align: center;
    box-shadow: 0 5px 15px rgba(0,0,0,0.05);
    border-top: 4px solid #667eea;
}

.stat-card .value {
    font-size: 28px;
    font-weight: 700;
    color: #4a3f6b;
}

.stat-card .label {
    color: #888;
    font-size: 14px;
    margin-top: 5px;
}

.report-columns {
    display: flex;
    gap: 20px;
    flex-wrap: wrap;
}

.report-columns > div {
    flex: 1;
    min-width: 250px;
}
'''

# ================== BASE TEMPLATE ==================
BASE_HTML = f'''
<!DOCTYPE html>
<html lang="fa">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>مدیریت هزینه سفر</title>
    <style>{BASE_CSS}</style>
</head>
<body>
    <div class="top-bar">
        <span class="brand">💰 سفر</span>
        <button class="hamburger" onclick="toggleMenu()" aria-label="منو">☰</button>
    </div>
    <div class="container">
        <nav class="navbar" id="navbar">
            <a href="/">🏠 خریدها</a>
            <a href="/add">➕ ثبت خرید</a>
            <a href="/families">👥 خانواده‌ها</a>
            <a href="/summary">📊 گزارش</a>
            <a href="/summary/export" class="btn-excel">📥 خروجی اکسل</a>
        </nav>
        {{% with messages = get_flashed_messages(with_categories=true) %}}
          {{% if messages %}}
            {{% for category, message in messages %}}
              <div class="alert alert-{{{{ category }}}}">{{{{ message }}}}</div>
            {{% endfor %}}
          {{% endif %}}
        {{% endwith %}}
        <div class="card">
            {{% block content %}}{{% endblock %}}
        </div>
    </div>
    <script>
        function toggleMenu() {{
            document.getElementById('navbar').classList.toggle('show');
        }}
        document.querySelectorAll('#navbar a').forEach(link => {{
            link.addEventListener('click', () => {{
                document.getElementById('navbar').classList.remove('show');
            }});
        }});
        setTimeout(function() {{
            var alerts = document.querySelectorAll('.alert');
            alerts.forEach(function(alert) {{
                alert.style.opacity = '0';
                setTimeout(function() {{ alert.remove(); }}, 500);
            }});
        }}, 4000);
    </script>
</body>
</html>
'''

# ================== ROUTES ==================
@app.route('/')
def index():
    return redirect(url_for('list_expenses'))

@app.route('/expenses')
def list_expenses():
    db = get_db()
    expenses = db.execute('''
        SELECT expenses.id, families.name AS family, persons.name AS person,
               expenses.description, expenses.price, expenses.date
        FROM expenses
        JOIN families ON expenses.family_id = families.id
        JOIN persons ON expenses.person_id = persons.id
        ORDER BY expenses.date DESC
    ''').fetchall()
    return render_template_string(
        BASE_HTML.replace('{% block content %}{% endblock %}', '''
            {% block content %}
            <h2>📋 تمام خریدها</h2>
            {% if expenses %}
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>خانواده</th><th>خریدار</th><th>شرح</th><th>مبلغ (تومان)</th><th>تاریخ</th><th>عملیات</th>
                        </tr>
                    </thead>
                    <tbody>
                    {% for e in expenses %}
                    <tr>
                        <td>{{ e['family'] }}</td>
                        <td>{{ e['person'] }}</td>
                        <td>{{ e['description'] }}</td>
                        <td>{{ "{:,}".format(e['price']) }}</td>
                        <td>{{ e['date'] }}</td>
                        <td><a href="/delete/{{ e['id'] }}" class="btn btn-danger" onclick="return confirm('حذف شود؟')">حذف</a></td>
                    </tr>
                    {% endfor %}
                    </tbody>
                </table>
            </div>
            {% else %}
            <p>هنوز خریدی ثبت نشده. از منوی بالا یک خرید اضافه کنید.</p>
            {% endif %}
            {% endblock %}
        '''), expenses=expenses)

@app.route('/delete/<int:expense_id>')
def delete_expense(expense_id):
    db = get_db()
    db.execute('DELETE FROM expenses WHERE id = ?', (expense_id,))
    db.commit()
    flash('هزینه با موفقیت حذف شد.', 'success')
    return redirect(url_for('list_expenses'))

@app.route('/add', methods=['GET', 'POST'])
def add_expense():
    db = get_db()
    families = db.execute('SELECT * FROM families ORDER BY name').fetchall()
    if not families:
        flash('ابتدا حداقل یک خانواده بسازید.', 'error')
        return redirect(url_for('manage_families'))

    persons_by_family = {}
    for fam in families:
        members = db.execute('SELECT * FROM persons WHERE family_id = ?', (fam['id'],)).fetchall()
        persons_by_family[fam['id']] = members

    if request.method == 'POST':
        family_id = request.form.get('family_id')
        person_id = request.form.get('person_id')
        description = request.form.get('description', '').strip()
        price_str = request.form.get('price', '').strip()

        errors = []
        if not family_id or not person_id:
            errors.append('خانواده و خریدار را انتخاب کنید.')
        if not description:
            errors.append('شرح خرید را وارد کنید.')
        try:
            price = int(price_str)
            if price <= 0:
                errors.append('مبلغ باید بزرگتر از صفر باشد.')
        except (ValueError, TypeError):
            errors.append('مبلغ را به عدد وارد کنید (تومان).')

        if errors:
            for e in errors:
                flash(e, 'error')
            return render_template_string(ADD_HTML, families=families, selected_family=family_id if family_id else None, persons_by_family=persons_by_family)

        db.execute('INSERT INTO expenses (family_id, person_id, description, price) VALUES (?,?,?,?)',
                   (int(family_id), int(person_id), description, price))
        db.commit()
        flash('✅ خرید با موفقیت ثبت شد.', 'success')
        return redirect(url_for('list_expenses'))

    return render_template_string(ADD_HTML, families=families, selected_family=None, persons_by_family=persons_by_family)

ADD_HTML = BASE_HTML.replace('{% block content %}{% endblock %}', '''
{% block content %}
<h2>➕ ثبت خرید جدید</h2>
<form method="post">
    <div class="inline-fields">
        <div class="form-group">
            <label>👨‍👩‍👧‍👦 خانواده</label>
            <select name="family_id" id="family_select" required onchange="updatePersons()">
                <option value="">-- انتخاب کنید --</option>
                {% for f in families %}
                <option value="{{ f['id'] }}" {% if selected_family and f['id']|string == selected_family %}selected{% endif %}>
                    {{ f['name'] }}
                </option>
                {% endfor %}
            </select>
        </div>
        <div class="form-group">
            <label>🧑 خریدار</label>
            <select name="person_id" id="person_select" required>
                <option value="">-- ابتدا خانواده را انتخاب کنید --</option>
            </select>
        </div>
    </div>
    <div class="form-group">
        <label>🛒 شرح خرید</label>
        <input type="text" name="description" placeholder="مثلاً شام، بنزین، بستنی..." required>
    </div>
    <div class="form-group">
        <label>💰 مبلغ (تومان)</label>
        <input type="number" name="price" placeholder="فقط عدد" min="1" required>
    </div>
    <button type="submit" class="btn-success">ثبت خرید</button>
</form>

<script>
    const allPersons = [
        {% for f in families %}
            {% for p in persons_by_family[f['id']] %}
                {id: {{ p['id'] }}, name: "{{ p['name'] }}", family_id: {{ f['id'] }} },
            {% endfor %}
        {% endfor %}
    ];

    function updatePersons() {
        const familyId = document.getElementById('family_select').value;
        const personSelect = document.getElementById('person_select');
        personSelect.innerHTML = '<option value="">-- انتخاب کنید --</option>';
        allPersons.forEach(p => {
            if (p.family_id == familyId) {
                const opt = document.createElement('option');
                opt.value = p.id;
                opt.textContent = p.name;
                personSelect.appendChild(opt);
            }
        });
    }

    window.onload = updatePersons;
</script>
{% endblock %}
''')

@app.route('/families', methods=['GET', 'POST'])
def manage_families():
    db = get_db()
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add_family':
            name = request.form.get('family_name', '').strip()
            if not name:
                flash('نام خانواده را وارد کنید.', 'error')
            else:
                try:
                    db.execute('INSERT INTO families (name) VALUES (?)', (name,))
                    db.commit()
                    flash(f'خانواده "{name}" اضافه شد.', 'success')
                except sqlite3.IntegrityError:
                    flash('این نام قبلاً ثبت شده.', 'error')
        elif action == 'add_member':
            family_id = request.form.get('family_id')
            person_name = request.form.get('person_name', '').strip()
            if not family_id or not person_name:
                flash('خانواده و نام عضو را وارد کنید.', 'error')
            else:
                db.execute('INSERT INTO persons (name, family_id) VALUES (?,?)', (person_name, int(family_id)))
                db.commit()
                flash('عضو جدید اضافه شد.', 'success')
        elif action == 'set_leader':
            family_id = request.form.get('family_id')
            leader_id = request.form.get('leader_id')
            if family_id and leader_id:
                db.execute('UPDATE families SET leader_id = ? WHERE id = ?', (int(leader_id), int(family_id)))
                db.commit()
                flash('لیدر خانواده بروزرسانی شد.', 'success')
        elif action == 'delete_family':
            family_id = request.form.get('family_id')
            if family_id:
                db.execute('DELETE FROM families WHERE id = ?', (int(family_id),))
                db.commit()
                flash('خانواده و تمام اعضای آن حذف شدند.', 'success')
        elif action == 'delete_member':
            member_id = request.form.get('member_id')
            if member_id:
                db.execute('DELETE FROM persons WHERE id = ?', (int(member_id),))
                db.commit()
                flash('عضو حذف شد.', 'success')
        return redirect(url_for('manage_families'))

    families = db.execute('SELECT * FROM families ORDER BY name').fetchall()
    family_data = []
    total_members = 0
    for fam in families:
        members = db.execute('SELECT * FROM persons WHERE family_id = ?', (fam['id'],)).fetchall()
        total_members += len(members)
        leader = db.execute('SELECT name FROM persons WHERE id = ?', (fam['leader_id'],)).fetchone() if fam['leader_id'] else None
        family_data.append({
            'family': fam,
            'members': members,
            'leader_name': leader['name'] if leader else 'تعیین نشده'
        })
    return render_template_string(FAMILIES_HTML, family_data=family_data, total_families=len(families), total_members=total_members)

FAMILIES_HTML = BASE_HTML.replace('{% block content %}{% endblock %}', '''
{% block content %}
<h2>👥 مدیریت خانواده‌ها</h2>

<div class="stats-bar">
    <span>👨‍👩‍👧‍👦 تعداد خانواده‌ها: <strong>{{ total_families }}</strong></span>
    <span>🧑 تعداد کل اعضا: <strong>{{ total_members }}</strong></span>
</div>

<form method="post" class="form-group" style="display: flex; gap: 10px; align-items: center;">
    <input type="hidden" name="action" value="add_family">
    <input type="text" name="family_name" placeholder="نام خانواده جدید" required style="flex:1;">
    <button type="submit" class="btn-success">➕ ساختن</button>
</form>

{% for data in family_data %}
<details class="family-details">
    <summary>
        <span>🏠 {{ data['family']['name'] }}</span>
        <span style="font-size:14px; color:#888;">{{ data['members']|length }} عضو | لیدر: {{ data['leader_name'] }}</span>
    </summary>
    <div class="family-content">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
            <span><strong>اعضا:</strong></span>
            <form method="post" onsubmit="return confirm('کل این خانواده و همه خریدهایش حذف شود؟')">
                <input type="hidden" name="action" value="delete_family">
                <input type="hidden" name="family_id" value="{{ data['family']['id'] }}">
                <button type="submit" class="btn btn-outline-danger" style="padding:6px 15px;">🗑 حذف خانواده</button>
            </form>
        </div>
        {% if data['members'] %}
        <div class="members-list">
            {% for m in data['members'] %}
            <div class="member-badge">
                {{ m['name'] }}
                <form method="post" style="display:inline;" onsubmit="return confirm('این عضو حذف شود؟')">
                    <input type="hidden" name="action" value="delete_member">
                    <input type="hidden" name="member_id" value="{{ m['id'] }}">
                    <button type="submit" class="btn btn-danger" style="padding:2px 8px; font-size:12px;">✕</button>
                </form>
                <form method="post" style="display:inline;">
                    <input type="hidden" name="action" value="set_leader">
                    <input type="hidden" name="family_id" value="{{ data['family']['id'] }}">
                    <input type="hidden" name="leader_id" value="{{ m['id'] }}">
                    <button type="submit" class="btn btn-warning" style="padding:4px 10px; font-size:12px;">👑 لیدر</button>
                </form>
            </div>
            {% endfor %}
        </div>
        {% else %}
        <p>هنوز عضوی ندارد.</p>
        {% endif %}
        <form method="post" class="inline-form">
            <input type="hidden" name="action" value="add_member">
            <input type="hidden" name="family_id" value="{{ data['family']['id'] }}">
            <input type="text" name="person_name" placeholder="نام عضو جدید" required>
            <button type="submit" class="btn-success">➕ افزودن</button>
        </form>
    </div>
</details>
{% endfor %}
{% endblock %}
''')

@app.route('/summary')
def summary():
    db = get_db()

    # هزینه هر نفر (واقعی)
    per_person = db.execute('''
        SELECT persons.id, persons.name, families.name as family, families.id as family_id,
               COALESCE(SUM(expenses.price), 0) as total
        FROM persons
        LEFT JOIN expenses ON expenses.person_id = persons.id
        JOIN families ON persons.family_id = families.id
        GROUP BY persons.id
        ORDER BY total DESC
    ''').fetchall()

    # هزینه هر خانواده (واقعی)
    per_family = db.execute('''
        SELECT families.id, families.name,
               COALESCE(SUM(expenses.price), 0) as total
        FROM families
        LEFT JOIN expenses ON expenses.family_id = families.id
        GROUP BY families.id
        ORDER BY total DESC
    ''').fetchall()

    # تعداد اعضای هر خانواده
    family_members = {}
    for fam in per_family:
        cnt = db.execute('SELECT COUNT(*) as cnt FROM persons WHERE family_id = ?', (fam['id'],)).fetchone()['cnt']
        family_members[fam['id']] = cnt

    total_all = sum(row['total'] for row in per_family)
    total_persons = sum(family_members.values())
    per_person_share = total_all / total_persons if total_persons > 0 else 0

    # تراز مالی هر خانواده
    family_balance = []
    for fam in per_family:
        paid = fam['total']
        members = family_members.get(fam['id'], 1)
        fair_share = per_person_share * members
        balance = paid - fair_share
        family_balance.append({
            'name': fam['name'],
            'members': members,
            'paid': paid,
            'fair_share': round(fair_share),
            'balance': round(balance)
        })

    return render_template_string(SUMMARY_HTML,
                                  per_person=per_person,
                                  family_balance=family_balance,
                                  total_all=total_all,
                                  total_persons=total_persons,
                                  per_person_share=round(per_person_share))

SUMMARY_HTML = BASE_HTML.replace('{% block content %}{% endblock %}', '''
{% block content %}
<h2>📊 گزارش نهایی و تسویه حساب</h2>

<div class="report-cards">
    <div class="stat-card">
        <div class="value">{{ "{:,}".format(total_all) }}</div>
        <div class="label">مجموع هزینه‌ها (تومان)</div>
    </div>
    <div class="stat-card">
        <div class="value">{{ total_persons }}</div>
        <div class="label">تعداد کل افراد</div>
    </div>
    <div class="stat-card">
        <div class="value">{{ "{:,}".format(per_person_share) }}</div>
        <div class="label">سرانه هر نفر (تومان)</div>
    </div>
</div>

<h3>💰 هزینه هر نفر (پرداختی واقعی)</h3>
<div class="table-wrapper">
    <table>
        <thead><tr><th>نام</th><th>خانواده</th><th>پرداخت واقعی</th></tr></thead>
        <tbody>
        {% for p in per_person %}
        <tr>
            <td>{{ p['name'] }}</td>
            <td>{{ p['family'] }}</td>
            <td>{{ "{:,}".format(p['total']) }}</td>
        </tr>
        {% endfor %}
        </tbody>
    </table>
</div>

<h3>🏠 تراز مالی خانواده‌ها</h3>
<div class="table-wrapper">
    <table>
        <thead>
            <tr>
                <th>خانواده</th>
                <th>تعداد اعضا</th>
                <th>پرداخت واقعی</th>
                <th>سهم عادلانه</th>
                <th>تراز</th>
                <th>وضعیت</th>
            </tr>
        </thead>
        <tbody>
        {% for fb in family_balance %}
        <tr>
            <td>{{ fb['name'] }}</td>
            <td>{{ fb['members'] }}</td>
            <td>{{ "{:,}".format(fb['paid']) }}</td>
            <td>{{ "{:,}".format(fb['fair_share']) }}</td>
            <td>{{ "{:,}".format(fb['balance']) }}</td>
            <td>
                {% if fb['balance'] > 0 %}
                    <span style="color:green;">طلبکار</span>
                {% elif fb['balance'] < 0 %}
                    <span style="color:red;">بدهکار</span>
                {% else %}
                    <span style="color:gray;">تسویه</span>
                {% endif %}
            </td>
        </tr>
        {% endfor %}
        </tbody>
    </table>
</div>

<div style="margin-top:20px; background:#f8f9fa; padding:15px; border-radius:12px;">
    <strong>📝 راهنمای تسویه:</strong><br>
    سرانه هر نفر <strong>{{ "{:,}".format(per_person_share) }}</strong> تومان است.<br>
    خانواده‌های <span style="color:red;">بدهکار</span> باید به خانواده‌های <span style="color:green;">طلبکار</span> پرداخت کنند تا حساب‌ها صاف شود.<br>
    (تراز مثبت یعنی طلبکار، تراز منفی یعنی بدهکار)
</div>

<a href="/summary/export" class="btn btn-success" style="margin-top:20px;">📥 دانلود اکسل کامل</a>
{% endblock %}
''')

# ================== EXCEL EXPORT (xlsxwriter) ==================
@app.route('/summary/export')
def export_excel():
    db = get_db()
    expenses = db.execute('''
        SELECT families.name AS family, persons.name AS person,
               expenses.description, expenses.price, expenses.date
        FROM expenses
        JOIN families ON expenses.family_id = families.id
        JOIN persons ON expenses.person_id = persons.id
        ORDER BY expenses.date DESC
    ''').fetchall()

    per_person = db.execute('''
        SELECT persons.name, families.name as family, COALESCE(SUM(expenses.price), 0) as total
        FROM persons
        LEFT JOIN expenses ON expenses.person_id = persons.id
        JOIN families ON persons.family_id = families.id
        GROUP BY persons.id
        ORDER BY total DESC
    ''').fetchall()

    per_family = db.execute('''
        SELECT families.id, families.name, COALESCE(SUM(expenses.price), 0) as total
        FROM families
        LEFT JOIN expenses ON expenses.family_id = families.id
        GROUP BY families.id
        ORDER BY total DESC
    ''').fetchall()

    family_members = {}
    for f in per_family:
        cnt = db.execute('SELECT COUNT(*) as cnt FROM persons WHERE family_id = ?', (f['id'],)).fetchone()['cnt']
        family_members[f['id']] = cnt

    total_all = sum(row['total'] for row in per_family) if per_family else 0
    total_persons = sum(family_members.values())
    per_person_share = total_all / total_persons if total_persons > 0 else 0

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)

    header_format = workbook.add_format({'bold': True, 'bg_color': '#667eea', 'font_color': 'white', 'border': 1})
    cell_format = workbook.add_format({'border': 1})
    number_format = workbook.add_format({'num_format': '#,##0', 'border': 1})
    rtl_format = workbook.add_format({'border': 1, 'reading_order': 2})

    # شیت ۱: لیست خریدها
    ws1 = workbook.add_worksheet('لیست خریدها')
    ws1.set_right_to_left()
    headers1 = ['خانواده', 'خریدار', 'شرح', 'مبلغ (تومان)', 'تاریخ']
    for col, h in enumerate(headers1):
        ws1.write(0, col, h, header_format)
    for i, e in enumerate(expenses, start=1):
        ws1.write(i, 0, e['family'], rtl_format)
        ws1.write(i, 1, e['person'], rtl_format)
        ws1.write(i, 2, e['description'], rtl_format)
        ws1.write(i, 3, e['price'], number_format)
        ws1.write(i, 4, e['date'], cell_format)
    ws1.set_column(0, 4, 15)

    # شیت ۲: هزینه هر نفر
    ws2 = workbook.add_worksheet('هزینه هر نفر')
    ws2.set_right_to_left()
    headers2 = ['نام', 'خانواده', 'جمع هزینه']
    for col, h in enumerate(headers2):
        ws2.write(0, col, h, header_format)
    for i, p in enumerate(per_person, start=1):
        ws2.write(i, 0, p['name'], rtl_format)
        ws2.write(i, 1, p['family'], rtl_format)
        ws2.write(i, 2, p['total'], number_format)
    ws2.set_column(0, 2, 20)

    # شیت ۳: هزینه هر خانواده
    ws3 = workbook.add_worksheet('هزینه هر خانواده')
    ws3.set_right_to_left()
    headers3 = ['خانواده', 'جمع هزینه']
    for col, h in enumerate(headers3):
        ws3.write(0, col, h, header_format)
    for i, f in enumerate(per_family, start=1):
        ws3.write(i, 0, f['name'], rtl_format)
        ws3.write(i, 1, f['total'], number_format)
    ws3.write(len(per_family)+1, 0, 'مجموع کل', workbook.add_format({'bold': True, 'border': 1, 'reading_order': 2}))
    ws3.write(len(per_family)+1, 1, total_all, workbook.add_format({'bold': True, 'num_format': '#,##0', 'border': 1}))
    ws3.set_column(0, 1, 20)

    # شیت ۴: تراز مالی خانواده‌ها
    ws4 = workbook.add_worksheet('تراز مالی خانواده‌ها')
    ws4.set_right_to_left()
    headers4 = ['خانواده', 'تعداد اعضا', 'پرداخت واقعی', 'سهم عادلانه', 'تراز', 'وضعیت']
    for col, h in enumerate(headers4):
        ws4.write(0, col, h, header_format)
    row = 1
    for f in per_family:
        paid = f['total']
        members = family_members.get(f['id'], 0)
        fair = per_person_share * members
        balance = paid - fair
        status = 'تسویه'
        if balance > 0:
            status = 'طلبکار'
        elif balance < 0:
            status = 'بدهکار'
        ws4.write(row, 0, f['name'], rtl_format)
        ws4.write(row, 1, members, cell_format)
        ws4.write(row, 2, paid, number_format)
        ws4.write(row, 3, round(fair), number_format)
        ws4.write(row, 4, round(balance), number_format)
        ws4.write(row, 5, status, rtl_format)
        row += 1
    ws4.set_column(0, 5, 15)

    workbook.close()
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='Trip_Expenses.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
