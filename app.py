import os
import uuid
import zipfile
import io
import time
from datetime import datetime, date
from functools import wraps
from flask import Flask, request, send_file, jsonify, render_template, session
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from pdf2docx import Converter
from docx import Document
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from docx2pdf import convert as docx2pdf_convert
from PIL import Image
import fitz  # PyMuPDF
from apscheduler.schedulers.background import BackgroundScheduler

app = Flask(__name__, static_folder='static', template_folder='templates')
app.secret_key = 'docutools_production_secret_key_change_me'

# Set maximum upload size limit to 20 MB
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024

# SQLite Database Configuration
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///docutools.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
bcrypt = Bcrypt(app)

UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
OUTPUT_FOLDER = os.path.join(os.getcwd(), 'outputs')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

GUEST_DAILY_LIMIT = 15
LOGGED_IN_DAILY_LIMIT = 20


# ------------------------------------------------------------------
# DATABASE MODEL
# ------------------------------------------------------------------
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(128), nullable=False)
    usage_count = db.Column(db.Integer, default=0)
    last_usage_date = db.Column(db.Date, default=date.today)

    def reset_usage_if_new_day(self):
        if self.last_usage_date != date.today():
            self.usage_count = 0
            self.last_usage_date = date.today()
            db.session.commit()


with app.app_context():
    db.create_all()


# ------------------------------------------------------------------
# ERROR HANDLERS
# ------------------------------------------------------------------
@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'error': 'File size exceeds the 20 MB limit. Please upload a smaller file.'}), 413


# ------------------------------------------------------------------
# BACKGROUND CLEANUP
# ------------------------------------------------------------------
def cleanup_old_files():
    now = time.time()
    cutoff = now - (15 * 60)
    for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
        if os.path.exists(folder):
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                if os.path.isfile(file_path):
                    if os.path.getmtime(file_path) < cutoff:
                        try:
                            os.remove(file_path)
                        except Exception as e:
                            print(f"Error removing {file_path}: {e}")


scheduler = BackgroundScheduler()
scheduler.add_job(func=cleanup_old_files, trigger="interval", minutes=10)
scheduler.start()


# ------------------------------------------------------------------
# HYBRID QUOTA DECORATOR
# ------------------------------------------------------------------
def check_hybrid_quota(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        user_id = session.get('user_id')
        today_str = date.today().isoformat()

        if user_id:
            user = User.query.get(user_id)
            if user:
                user.reset_usage_if_new_day()
                if user.usage_count >= LOGGED_IN_DAILY_LIMIT:
                    return jsonify({
                        'error': f'Daily limit of {LOGGED_IN_DAILY_LIMIT} conversions reached. Please return tomorrow!'
                    }), 429

                response = f(*args, **kwargs)
                if isinstance(response, tuple) and response[1] != 200:
                    return response
                user.usage_count += 1
                db.session.commit()
                return response

        if session.get('last_usage_date') != today_str:
            session['last_usage_date'] = today_str
            session['guest_usage_count'] = 0

        current_guest_usage = session.get('guest_usage_count', 0)
        if current_guest_usage >= GUEST_DAILY_LIMIT:
            return jsonify({
                'error': f'Guest daily limit reached ({GUEST_DAILY_LIMIT}/{GUEST_DAILY_LIMIT}). Log in or Sign Up to unlock 20 daily conversions!'
            }), 429

        response = f(*args, **kwargs)
        if isinstance(response, tuple) and response[1] != 200:
            return response
        session['guest_usage_count'] = current_guest_usage + 1
        return response

    return decorated_function


# ------------------------------------------------------------------
# AUTH & STATUS ROUTES
# ------------------------------------------------------------------
@app.route('/')
def home():
    return render_template('index.html')


@app.route('/api/auth/status', methods=['GET'])
def auth_status():
    user_id = session.get('user_id')
    today_str = date.today().isoformat()

    if user_id:
        user = User.query.get(user_id)
        if user:
            user.reset_usage_if_new_day()
            return jsonify({
                'logged_in': True,
                'username': user.username,
                'used': user.usage_count,
                'limit': LOGGED_IN_DAILY_LIMIT
            })

    if session.get('last_usage_date') != today_str:
        session['last_usage_date'] = today_str
        session['guest_usage_count'] = 0

    return jsonify({
        'logged_in': False,
        'used': session.get('guest_usage_count', 0),
        'limit': GUEST_DAILY_LIMIT,
        'unlocked_limit': LOGGED_IN_DAILY_LIMIT
    })


@app.route('/api/auth/register', methods=['POST'])
def register():
    data = request.get_json() or {}
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()

    if not username or not password:
        return jsonify({'error': 'Username and password are required.'}), 400

    if User.query.filter_by(username=username).first():
        return jsonify({'error': 'Username is already taken.'}), 400

    hashed_pw = bcrypt.generate_password_hash(password).decode('utf-8')
    new_user = User(username=username, password_hash=hashed_pw)
    db.session.add(new_user)
    db.session.commit()

    session['user_id'] = new_user.id
    return jsonify(
        {'message': 'Registration successful! Unlocked 20 daily conversions.', 'username': new_user.username})


@app.route('/api/auth/login', methods=['POST'])
def login():
    data = request.get_json() or {}
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()

    user = User.query.filter_by(username=username).first()
    if user and bcrypt.check_password_hash(user.password_hash, password):
        session['user_id'] = user.id
        user.reset_usage_if_new_day()
        return jsonify({'message': 'Logged in! You now have 20 daily conversions.', 'username': user.username})

    return jsonify({'error': 'Invalid username or password.'}), 401


@app.route('/api/auth/logout', methods=['POST'])
def logout():
    session.pop('user_id', None)
    return jsonify({'message': 'Logged out successfully.'})


# ------------------------------------------------------------------
# CONVERSION ROUTES
# ------------------------------------------------------------------
@app.route('/api/convert/pdf-to-word', methods=['POST'])
@check_hybrid_quota
def pdf_to_word():
    file = request.files.get('file')
    if not file:
        return jsonify({'error': 'No file uploaded'}), 400

    uid = str(uuid.uuid4())
    pdf_path = os.path.join(UPLOAD_FOLDER, f"{uid}.pdf")
    docx_path = os.path.join(OUTPUT_FOLDER, f"{uid}.docx")
    file.save(pdf_path)

    try:
        try:
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()
        except Exception as primary_err:
            print(f"[pdf2docx failed, switching to PyMuPDF fallback]: {primary_err}")
            doc = fitz.open(pdf_path)
            word_doc = Document()
            for page in doc:
                text = page.get_text("text")
                if text.strip():
                    word_doc.add_paragraph(text)
            doc.close()
            word_doc.save(docx_path)

        out_name = f"{file.filename.rsplit('.', 1)[0]}_converted.docx"
        return send_file(
            docx_path,
            as_attachment=True,
            download_name=out_name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        print(f"[PDF TO WORD ERROR]: {str(e)}")
        return jsonify({'error': f'PDF conversion failed: {str(e)}'}), 500
    finally:
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except:
                pass


@app.route('/api/convert/word-to-pdf', methods=['POST'])
@check_hybrid_quota
def word_to_pdf():
    file = request.files.get('file')
    if not file: return jsonify({'error': 'No file uploaded'}), 400
    uid = str(uuid.uuid4())
    docx_path = os.path.join(UPLOAD_FOLDER, f"{uid}.docx")
    pdf_path = os.path.join(OUTPUT_FOLDER, f"{uid}.pdf")
    file.save(docx_path)
    try:
        docx2pdf_convert(docx_path, pdf_path)
        out_name = f"{file.filename.rsplit('.', 1)[0]}_converted.pdf"
        return send_file(pdf_path, as_attachment=True, download_name=out_name, mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': f'Word to PDF failed: Ensure MS Word is installed locally ({str(e)})'}), 500
    finally:
        if os.path.exists(docx_path): os.remove(docx_path)


@app.route('/api/convert/merge-pdf', methods=['POST'])
@check_hybrid_quota
def merge_pdf():
    files = request.files.getlist('files')
    if not files or len(files) < 2: return jsonify({'error': 'Select at least 2 PDFs to merge.'}), 400
    uid = str(uuid.uuid4())
    merged_path = os.path.join(OUTPUT_FOLDER, f"{uid}_merged.pdf")
    merger = PdfMerger()
    saved_paths = []
    try:
        for f in files:
            path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}.pdf")
            f.save(path)
            saved_paths.append(path)
            merger.append(path)
        merger.write(merged_path)
        merger.close()
        return send_file(merged_path, as_attachment=True, download_name="Merged_Document.pdf",
                         mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': f'Merge failed: {str(e)}'}), 500
    finally:
        for p in saved_paths:
            if os.path.exists(p): os.remove(p)


@app.route('/api/convert/image-to-pdf', methods=['POST'])
@check_hybrid_quota
def image_to_pdf():
    files = request.files.getlist('files')
    if not files: return jsonify({'error': 'No images uploaded'}), 400
    uid = str(uuid.uuid4())
    pdf_path = os.path.join(OUTPUT_FOLDER, f"{uid}_images.pdf")
    image_list = []
    try:
        for f in files:
            img = Image.open(f.stream).convert('RGB')
            image_list.append(img)
        image_list[0].save(pdf_path, save_all=True, append_images=image_list[1:])
        return send_file(pdf_path, as_attachment=True, download_name="Converted_Images.pdf", mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': f'Image conversion failed: {str(e)}'}), 500


@app.route('/api/convert/compress-pdf', methods=['POST'])
@check_hybrid_quota
def compress_pdf():
    file = request.files.get('file')
    if not file: return jsonify({'error': 'No file uploaded'}), 400
    uid = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_FOLDER, f"{uid}.pdf")
    output_path = os.path.join(OUTPUT_FOLDER, f"{uid}_compressed.pdf")
    file.save(input_path)
    try:
        doc = fitz.open(input_path)
        doc.save(output_path, deflate=True, garbage=4, clean=True)
        doc.close()
        out_name = f"{file.filename.rsplit('.', 1)[0]}_compressed.pdf"
        return send_file(output_path, as_attachment=True, download_name=out_name, mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': f'Compression failed: {str(e)}'}), 500
    finally:
        if os.path.exists(input_path): os.remove(input_path)


@app.route('/api/convert/split-pdf', methods=['POST'])
@check_hybrid_quota
def split_pdf():
    file = request.files.get('file')
    pages_param = request.form.get('pages', '').strip()
    if not file: return jsonify({'error': 'No file uploaded'}), 400
    uid = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_FOLDER, f"{uid}.pdf")
    file.save(input_path)
    try:
        reader = PdfReader(input_path)
        total_pages = len(reader.pages)
        if pages_param:
            selected_pages = []
            for part in pages_param.split(','):
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    selected_pages.extend(range(start - 1, min(end, total_pages)))
                else:
                    pg = int(part) - 1
                    if 0 <= pg < total_pages: selected_pages.append(pg)
        else:
            selected_pages = list(range(total_pages))

        download_urls = []
        for idx in selected_pages:
            writer = PdfWriter()
            writer.add_page(reader.pages[idx])
            out_filename = f"{uid}_page_{idx + 1}.pdf"
            out_path = os.path.join(OUTPUT_FOLDER, out_filename)
            with open(out_path, 'wb') as f_out:
                writer.write(f_out)
            download_urls.append(f"/api/download/{out_filename}")
        return jsonify({'files': download_urls})
    except Exception as e:
        return jsonify({'error': f'Split failed: {str(e)}'}), 500
    finally:
        if os.path.exists(input_path): os.remove(input_path)


@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    file_path = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(file_path):
        page_num = filename.rsplit('_page_', 1)[-1]
        return send_file(file_path, as_attachment=True, download_name=f"Split_Page_{page_num}")
    return jsonify({'error': 'File not found'}), 404


@app.route('/api/convert/pdf-to-image', methods=['POST'])
@check_hybrid_quota
def pdf_to_image():
    file = request.files.get('file')
    if not file: return jsonify({'error': 'No file uploaded'}), 400
    uid = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_FOLDER, f"{uid}.pdf")
    file.save(input_path)
    try:
        doc = fitz.open(input_path)
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w') as zf:
            for i, page in enumerate(doc):
                pix = page.get_pixmap(dpi=150)
                zf.writestr(f"Page_{i + 1}.png", pix.tobytes("png"))
        doc.close()
        memory_file.seek(0)
        out_name = f"{file.filename.rsplit('.', 1)[0]}_images.zip"
        return send_file(memory_file, mimetype='application/zip', as_attachment=True, download_name=out_name)
    except Exception as e:
        return jsonify({'error': f'PDF to Image failed: {str(e)}'}), 500
    finally:
        if os.path.exists(input_path): os.remove(input_path)


@app.route('/api/convert/rotate-pdf', methods=['POST'])
@check_hybrid_quota
def rotate_pdf():
    file = request.files.get('file')
    angle = int(request.form.get('angle', 90))
    if not file: return jsonify({'error': 'No file uploaded'}), 400
    uid = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_FOLDER, f"{uid}.pdf")
    output_path = os.path.join(OUTPUT_FOLDER, f"{uid}_rotated.pdf")
    file.save(input_path)
    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()
        for page in reader.pages:
            page.rotate(angle)
            writer.add_page(page)
        with open(output_path, 'wb') as f_out:
            writer.write(f_out)
        out_name = f"{file.filename.rsplit('.', 1)[0]}_rotated.pdf"
        return send_file(output_path, as_attachment=True, download_name=out_name, mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': f'Rotation failed: {str(e)}'}), 500
    finally:
        if os.path.exists(input_path): os.remove(input_path)


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)