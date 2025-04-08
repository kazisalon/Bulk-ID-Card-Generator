from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
import pandas as pd
import os
from werkzeug.utils import secure_filename
import tempfile
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import Image
from reportlab.lib.units import inch
from PIL import Image as PILImage
import io
import uuid

app = Flask(__name__)
CORS(app)  # Enable CORS for development

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Add templates configuration
TEMPLATES = {
    'template1': {
        'name': 'Basic Template',
        'background_color': '#FFFFFF',
        'text_color': '#000000',
        'font': 'Helvetica',
        'font_size': 12,
        'type': 'basic'
    },
    'template2': {
        'name': 'Modern Template',
        'background_color': '#dc3545',
        'text_color': '#FFFFFF',
        'font': 'Helvetica-Bold',
        'font_size': 14,
        'type': 'modern'
    }
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html', templates=TEMPLATES)

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        try:
            # Read Excel file and validate required columns
            df = pd.read_excel(file_path)
            required_columns = {'student_name', 'phone_number', 'id_number', 'photo_path'}
            missing_columns = required_columns - set(df.columns)
            
            if missing_columns:
                return jsonify({'error': f'Missing required columns: {", ".join(missing_columns)}'}), 400
            
            return jsonify({
                'message': 'File uploaded successfully',
                'columns': list(df.columns),
                'preview': df.head(5).to_dict('records'),
                'filename': filename
            })
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    
    return jsonify({'error': 'File type not allowed'}), 400

@app.route('/api/generate-pdf', methods=['POST'])
def generate_pdf():
    data = request.json
    
    if not data or 'filePath' not in data or 'templateId' not in data:
        return jsonify({'error': 'Missing required data'}), 400
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(data['filePath']))
    template = TEMPLATES.get(data['templateId'])
    
    try:
        df = pd.read_excel(file_path)
        if 'id_number' in df.columns:
            df['id_number'] = df['id_number'].astype(str)
        if 'phone_number' in df.columns:
            df['phone_number'] = df['phone_number'].astype(str)
        
        # Create temporary file with .pdf extension
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        pdf_path = temp_file.name
        temp_file.close()
        
        # Create canvas with custom size
        c = canvas.Canvas(pdf_path, pagesize=(612, 396))
        width, height = 612, 396
        
        for i, row in enumerate(df.to_dict('records')):
            if i > 0:
                c.showPage()
            
            if template['type'] == 'modern':
                # Red background
                c.setFillColor(colors.HexColor('#dc3545'))
                c.rect(0, 0, width, height, fill=True)
                
                # Yellow curved wave
                c.setFillColor(colors.HexColor('#FFC107'))
                p = c.beginPath()
                p.moveTo(0, height/2)
                p.curveTo(width/3, height/2 + 50, width*2/3, height/2 - 50, width, height/2)
                p.lineTo(width, 0)
                p.lineTo(0, 0)
                p.close()
                c.drawPath(p, fill=True)
                
                # White curved wave
                c.setFillColor(colors.white)
                p = c.beginPath()
                p.moveTo(0, height/2 - 30)
                p.curveTo(width/3, height/2 + 20, width*2/3, height/2 - 80, width, height/2 - 30)
                p.lineTo(width, 0)
                p.lineTo(0, 0)
                p.close()
                c.drawPath(p, fill=True)
                
                # Photo with white circular border
                if 'photo_path' in row and str(row['photo_path']) != 'nan':
                    try:
                        # White circle background
                        c.setFillColor(colors.white)
                        c.circle(width/2, height - 120, 60, fill=True)
                        
                        # Red circle border
                        c.setStrokeColor(colors.HexColor('#dc3545'))
                        c.setLineWidth(5)
                        c.circle(width/2, height - 120, 55, stroke=True)
                        
                        # Photo
                        img_path = str(row['photo_path'])
                        c.drawImage(img_path, 
                                  width/2 - 50, height - 170,
                                  width=100, height=100,
                                  mask='auto')
                    except Exception as e:
                        print(f"Error loading image: {e}")
                
                # Name and role
                c.setFillColor(colors.black)
                c.setFont('Helvetica-Bold', 24)
                c.drawCentredString(width/2, height - 220, str(row.get('student_name', '')))
                c.setFont('Helvetica', 14)
                c.drawCentredString(width/2, height - 245, "STUDENT")
                
                # Contact information with icons
                icon_x = width/2 - 100
                base_y = 120
                spacing = 30
                
                c.setFont('Helvetica', 12)
                # Email/ID
                c.drawString(icon_x + 30, base_y, str(row.get('id_number', '')))
                # Phone
                c.drawString(icon_x + 30, base_y - spacing, str(row.get('phone_number', '')))
                
            c.setFillColor(colors.black)  # Reset color
        
        c.save()
        
        return jsonify({
            'message': 'PDF generated successfully',
            'downloadUrl': f'/api/download/{os.path.basename(pdf_path)}',
            'filename': f"student_id_cards_{uuid.uuid4().hex[:8]}.pdf"
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    if '..' in filename or '/' in filename:
        return jsonify({'error': 'Invalid filename'}), 400
    
    if filename.endswith('.pdf'):
        return send_file(os.path.join(tempfile.gettempdir(), filename),
                        download_name="student_id_cards.pdf",
                        as_attachment=True)
    
    return jsonify({'error': 'File not found'}), 404

if __name__ == '__main__':
    app.run(debug=True)