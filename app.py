# backend/app.py
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
import os
import base64
import io
from PIL import Image
import weasyprint
import tempfile
import uuid
import json

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    if file and file.filename.endswith(('.xlsx', '.xls')):
        # Save the file temporarily
        temp_path = os.path.join(UPLOAD_FOLDER, str(uuid.uuid4()) + '.xlsx')
        file.save(temp_path)
        
        try:
            # Read the Excel file
            df = pd.read_excel(temp_path)
            # Convert DataFrame to dict for JSON response
            data = df.to_dict(orient='records')
            
            # Process image paths if they exist
            for item in data:
                if 'Photo' in item and item['Photo'] and isinstance(item['Photo'], str):
                    # If it's a file path and the file exists
                    photo_path = item['Photo']
                    if os.path.isfile(photo_path):
                        with open(photo_path, 'rb') as img_file:
                            img_data = base64.b64encode(img_file.read()).decode('utf-8')
                            item['Photo'] = f"data:image/jpeg;base64,{img_data}"
            
            # Remove the temporary file
            os.remove(temp_path)
            
            return jsonify({
                "message": "File processed successfully",
                "data": data,
                "total": len(data)
            })
            
        except Exception as e:
            # Remove the temporary file in case of error
            if os.path.exists(temp_path):
                os.remove(temp_path)
            return jsonify({"error": str(e)}), 500
    
    return jsonify({"error": "File type not allowed"}), 400

@app.route('/api/upload-logo', methods=['POST'])
def upload_logo():
    if 'logo' not in request.files:
        return jsonify({"error": "No logo file part"}), 400
    
    file = request.files['logo']
    
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    allowed_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.svg'}
    file_ext = os.path.splitext(file.filename)[1].lower()
    
    if file_ext not in allowed_extensions:
        return jsonify({"error": "File type not allowed"}), 400
    
    try:
        # Process the image
        image_data = file.read()
        
        # For non-SVG files, use PIL to process
        if file_ext != '.svg':
            img = Image.open(io.BytesIO(image_data))
            
            # Resize if too large
            max_size = (300, 300)
            if img.width > max_size[0] or img.height > max_size[1]:
                img.thumbnail(max_size)
            
            # Convert back to bytes
            buffer = io.BytesIO()
            img.save(buffer, format="PNG")
            image_data = buffer.getvalue()
        
        # Encode as base64
        encoded_image = base64.b64encode(image_data).decode('utf-8')
        mime_type = f"image/{file_ext[1:]}" if file_ext != '.svg' else "image/svg+xml"
        data_url = f"data:{mime_type};base64,{encoded_image}"
        
        return jsonify({
            "message": "Logo uploaded successfully",
            "logo": data_url
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/save-image', methods=['POST'])
def save_image():
    try:
        data = request.json
        if not data or 'image' not in data:
            return jsonify({"error": "No image data provided"}), 400
        
        # The image data is already in base64 format from the client
        return jsonify({
            "message": "Image saved successfully",
            "image": data['image']
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/generate-pdf', methods=['POST'])
def generate_pdf():
    try:
        # Get data and template settings from request
        data = request.json.get('data', [])
        template = request.json.get('template', {})
        selected_ids = request.json.get('selectedIds', [])
        
        # If specific IDs are provided, filter the data
        if selected_ids:
            data = [item for item in data if str(item.get('id', '')) in selected_ids]
        
        # Process card size
        card_width = template.get('cardWidth', '3.375in')
        card_height = template.get('cardHeight', '2.125in')
        orientation = template.get('orientation', 'landscape')
        
        # If portrait, swap width and height
        if orientation == 'portrait':
            card_width, card_height = card_height, card_width
        
        # Create HTML for each ID card based on template type
        html_content = "<html><body style='margin: 0; padding: 0;'>"
        
        template_type = template.get('templateType', 'standard')
        
        for item in data:
            # Create different HTML based on template type
            if template_type == 'corporate':
                card_html = create_corporate_template(item, template, card_width, card_height)
            elif template_type == 'casual':
                card_html = create_casual_template(item, template, card_width, card_height)
            else:  # standard template
                card_html = create_standard_template(item, template, card_width, card_height)
                
            html_content += card_html
        
        html_content += "</body></html>"
        
        # Generate PDF from HTML
        pdf = weasyprint.HTML(string=html_content).write_pdf()
        
        # Create a temporary file to store the PDF
        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_pdf.write(pdf)
        temp_pdf.close()
        
        # Send the PDF file as a response
        return send_file(temp_pdf.name, mimetype='application/pdf', 
                         download_name='id_cards.pdf', as_attachment=True)
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def create_standard_template(item, template, width, height):
    return f"""
    <div class="id-card" style="
        width: {width}; 
        height: {height}; 
        border: 1px solid #000; 
        margin: 0.1in; 
        display: inline-block;
        box-sizing: border-box;
        padding: 0.1in;
        background-color: {template.get('backgroundColor', '#ffffff')};
        color: {template.get('textColor', '#000000')};
        font-family: {template.get('font', 'Arial, sans-serif')};
        page-break-after: always;
    ">
        <div class="header" style="text-align: center; font-weight: bold; font-size: 16px;">
            {'' if not template.get('logo') else f'<img src="{template.get("logo")}" style="max-height: 40px; max-width: 100px; margin-bottom: 5px;" />'}
            <div>{template.get('headerText', 'IDENTIFICATION CARD')}</div>
        </div>
        
        <div class="content" style="display: flex; margin-top: 10px;">
            <div class="photo" style="width: 1in; height: 1.25in; border: 1px solid #ccc; overflow: hidden;">
                {f'<img src="{item.get("Photo", "")}" style="width: 100%; height: 100%; object-fit: cover;">' if item.get("Photo") else ''}
            </div>
            
            <div class="info" style="margin-left: 0.2in; flex: 1;">
                <p style="margin: 2px 0;"><strong>Name:</strong> {item.get("Name", "")}</p>
                <p style="margin: 2px 0;"><strong>ID:</strong> {item.get("ID", "")}</p>
                {"".join([f'<p style="margin: 2px 0;"><strong>{key}:</strong> {value}</p>' for key, value in item.items() 
                          if key not in ["Name", "ID", "Photo", "id"]])}
            </div>
        </div>
        
        <div class="footer" style="text-align: center; font-size: 12px; margin-top: 10px;">
            {template.get('footerText', '')}
        </div>
    </div>
    """

def create_corporate_template(item, template, width, height):
    return f"""
    <div class="id-card" style="
        width: {width}; 
        height: {height}; 
        border: none;
        border-radius: 10px;
        margin: 0.1in; 
        display: inline-block;
        box-sizing: border-box;
        padding: 0.15in;
        background: linear-gradient(135deg, {template.get('backgroundColor', '#1a365d')}, #0f172a);
        color: {template.get('textColor', '#ffffff')};
        font-family: {template.get('font', 'Arial, sans-serif')};
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        page-break-after: always;
    ">
        <div class="header" style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px;">
            {'' if not template.get('logo') else f'<img src="{template.get("logo")}" style="max-height: 40px; max-width: 100px;" />'}
            <div style="font-weight: bold; font-size: 14px; text-transform: uppercase; letter-spacing: 1px;">{template.get('headerText', 'CORPORATE ID')}</div>
        </div>
        
        <div class="content" style="display: flex; align-items: center;">
            <div class="photo" style="width: 1.1in; height: 1.35in; border-radius: 5px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.2);">
                {f'<img src="{item.get("Photo", "")}" style="width: 100%; height: 100%; object-fit: cover;">' if item.get("Photo") else ''}
            </div>
            
            <div class="info" style="margin-left: 0.2in; flex: 1;">
                <p style="margin: 3px 0; font-size: 16px; font-weight: 600;">{item.get("Name", "")}</p>
                <p style="margin: 3px 0; font-size: 14px; opacity: 0.9;">{item.get("Position", "")}</p>
                <p style="margin: 3px 0; font-size: 12px; opacity: 0.8;">{item.get("Department", "")}</p>
                <p style="margin: 3px 0; font-size: 12px; opacity: 0.8;">ID: {item.get("ID", "")}</p>
                {"".join([f'<p style="margin: 3px 0; font-size: 12px; opacity: 0.8;">{key}: {value}</p>' for key, value in item.items() 
                          if key not in ["Name", "ID", "Photo", "Position", "Department", "id"]])}
            </div>
        </div>
        
        <div class="footer" style="text-align: center; font-size: 10px; margin-top: 12px; opacity: 0.7; border-top: 1px solid rgba(255,255,255,0.2); padding-top: 8px;">
            {template.get('footerText', '')}
        </div>
    </div>
    """

def create_casual_template(item, template, width, height):
    return f"""
    <div class="id-card" style="
        width: {width}; 
        height: {height}; 
        border: 2px dashed {template.get('accentColor', '#3b82f6')};
        border-radius: 8px;
        margin: 0.1in; 
        display: inline-block;
        box-sizing: border-box;
        padding: 0.12in;
        background-color: {template.get('backgroundColor', '#ffffff')};
        color: {template.get('textColor', '#374151')};
        font-family: {template.get('font', 'Arial, sans-serif')};
        page-break-after: always;
    ">
        <div class="header" style="text-align: center; font-weight: bold; font-size: 16px; color: {template.get('accentColor', '#3b82f6')};">
            {'' if not template.get('logo') else f'<img src="{template.get("logo")}" style="max-height: 45px; max-width: 100px; margin-bottom: 5px;" />'}
            <div>{template.get('headerText', 'MEMBER CARD')}</div>
        </div>
        
        <div class="content" style="display: flex; flex-direction: column; align-items: center; margin-top: 5px;">
            <div class="photo" style="width: 1.2in; height: 1.2in; border-radius: 50%; border: 3px solid {template.get('accentColor', '#3b82f6')}; overflow: hidden; margin-bottom: 8px;">
                {f'<img src="{item.get("Photo", "")}" style="width: 100%; height: 100%; object-fit: cover;">' if item.get("Photo") else ''}
            </div>
            
            <div class="info" style="text-align: center; width: 100%;">
                <p style="margin: 2px 0; font-size: 16px; font-weight: bold;">{item.get("Name", "")}</p>
                <p style="margin: 2px 0;"><strong>ID:</strong> {item.get("ID", "")}</p>
                {"".join([f'<p style="margin: 2px 0;"><strong>{key}:</strong> {value}</p>' for key, value in item.items() 
                          if key not in ["Name", "ID", "Photo", "id"]])}
            </div>
        </div>
        
        <div class="footer" style="text-align: center; font-size: 11px; margin-top: 8px; font-style: italic;">
            {template.get('footerText', '')}
        </div>
    </div>
    """

if __name__ == '__main__':
    app.run(debug=True, port=5000)