from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
from werkzeug.utils import secure_filename
import os
import json
from datetime import datetime
from processor import process_docx

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'imr-dev-key-change-in-production')

SUBMISSION_PASSWORD = os.environ.get('SUBMISSION_PASSWORD', 'imr2026')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'imradmin2026')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# On Railway, set DATA_DIR to a mounted Volume path (e.g. /data) so
# uploads/output/submissions.json survive redeploys. Without a Volume,
# Railway's filesystem is wiped on every redeploy.
DATA_DIR = os.environ.get('DATA_DIR', BASE_DIR)
UPLOAD_FOLDER = os.path.join(DATA_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(DATA_DIR, 'output')
SUBMISSIONS_LOG = os.path.join(DATA_DIR, 'submissions.json')
ALLOWED_EXTENSIONS = {'docx'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def load_submissions():
    if not os.path.exists(SUBMISSIONS_LOG):
        return []
    with open(SUBMISSIONS_LOG, 'r') as f:
        return json.load(f)


def save_submission(entry):
    submissions = load_submissions()
    submissions.append(entry)
    with open(SUBMISSIONS_LOG, 'w') as f:
        json.dump(submissions, f, indent=2)


# ─── Contributor login ───

@app.route('/', methods=['GET', 'POST'])
def login():
    if session.get('authenticated'):
        return redirect(url_for('submit'))
    if request.method == 'POST':
        if request.form.get('password') == SUBMISSION_PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('submit'))
        flash('Incorrect password.')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ─── Contributor submission ───

@app.route('/submit', methods=['GET', 'POST'])
def submit():
    if not session.get('authenticated'):
        return redirect(url_for('login'))

    if request.method == 'POST':
        title = request.form.get('title', '').strip()
        author = request.form.get('author', '').strip()
        standfirst = request.form.get('standfirst', '').strip()
        article_type = request.form.get('article_type', 'article')
        email = request.form.get('email', '').strip()

        if not title or not author:
            flash('Title and author are required.')
            return render_template('submit.html')

        if 'docx_file' not in request.files or request.files['docx_file'].filename == '':
            flash('Please upload a Word document.')
            return render_template('submit.html')

        file = request.files['docx_file']
        if not allowed_file(file.filename):
            flash('Please upload a valid .docx file.')
            return render_template('submit.html')

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_name = secure_filename(file.filename)
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{timestamp}_{safe_name}')
        file.save(upload_path)

        try:
            output_path = process_docx(
                filepath=upload_path,
                title=title,
                author=author,
                standfirst=standfirst,
                article_type=article_type,
                output_folder=app.config['OUTPUT_FOLDER'],
                timestamp=timestamp,
            )
        except Exception as e:
            flash(f'Something went wrong processing that file: {e}')
            return render_template('submit.html')

        save_submission({
            'timestamp': timestamp,
            'title': title,
            'author': author,
            'email': email,
            'standfirst': standfirst,
            'article_type': article_type,
            'original_filename': file.filename,
            'output_file': os.path.basename(output_path),
            'received_at': datetime.now().isoformat(timespec='seconds'),
        })

        return render_template('submit.html', success=True)

    return render_template('submit.html')


# ─── Admin area ───

@app.route('/admin', methods=['GET', 'POST'])
def admin_login():
    if session.get('is_admin'):
        return redirect(url_for('admin_dashboard'))
    if request.method == 'POST':
        if request.form.get('password') == ADMIN_PASSWORD:
            session['is_admin'] = True
            return redirect(url_for('admin_dashboard'))
        flash('Incorrect admin password.')
    return render_template('admin_login.html')


@app.route('/admin/dashboard')
def admin_dashboard():
    if not session.get('is_admin'):
        return redirect(url_for('admin_login'))
    submissions = list(reversed(load_submissions()))
    return render_template('admin.html', submissions=submissions)


@app.route('/admin/download/<filename>')
def admin_download(filename):
    if not session.get('is_admin'):
        return redirect(url_for('admin_login'))
    safe_name = secure_filename(filename)
    path = os.path.join(app.config['OUTPUT_FOLDER'], safe_name)
    if not os.path.exists(path):
        flash('File not found.')
        return redirect(url_for('admin_dashboard'))
    return send_file(path, as_attachment=True)


@app.route('/admin/reprocess_all', methods=['POST'])
def admin_reprocess_all():
    """Re-run every existing submission's original upload through the
    current processor.py - useful after a style/layout change, so
    already-submitted articles pick up the new look without contributors
    having to resubmit. Reuses each submission's original timestamp, so
    the regenerated file overwrites the old output in place rather than
    creating a duplicate entry."""
    if not session.get('is_admin'):
        return redirect(url_for('admin_login'))

    submissions = load_submissions()
    success_count = 0
    failed = []
    skipped_old = []

    for sub in submissions:
        title = sub.get('title', 'Untitled')

        # Submissions made before standfirst tracking was added to the
        # log can't be safely reprocessed - reprocessing would silently
        # drop the standfirst rather than reproduce it.
        if 'standfirst' not in sub:
            skipped_old.append(title)
            continue

        safe_name = secure_filename(sub.get('original_filename', ''))
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{sub['timestamp']}_{safe_name}")
        if not os.path.exists(upload_path):
            failed.append(title)
            continue

        try:
            process_docx(
                filepath=upload_path,
                title=sub['title'],
                author=sub['author'],
                standfirst=sub['standfirst'],
                article_type=sub.get('article_type', 'article'),
                output_folder=app.config['OUTPUT_FOLDER'],
                timestamp=sub['timestamp'],
            )
            success_count += 1
        except Exception:
            failed.append(title)

    message = f'Reprocessed {success_count} submission(s) with the current style.'
    if failed:
        message += f' Could not reprocess (missing original upload or error): {", ".join(failed)}.'
    if skipped_old:
        message += f' Skipped (submitted before reprocessing support was added - ask them to resubmit): {", ".join(skipped_old)}.'
    flash(message)
    return redirect(url_for('admin_dashboard'))


@app.route('/admin/logout')
def admin_logout():
    session.pop('is_admin', None)
    return redirect(url_for('admin_login'))


if __name__ == '__main__':
    app.run(debug=True)
