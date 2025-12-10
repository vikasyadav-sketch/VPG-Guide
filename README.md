# Vehicle Platform Guide Generator - Web App

A web application that converts formatted Word documents into beautiful HTML vehicle platform guides. Upload a `.docx` file and 4 car images to instantly generate a professional HTML guide.

## 🚀 Features

- **Simple Upload Interface** - Drag and drop or select files
- **Automatic Processing** - Parses Word documents and extracts:
  - Vehicle specifications
  - Common issues by category
  - Parts and brands information
  - Hyperlinks and fault codes
- **Embedded Images** - Car images are embedded as base64 for standalone HTML
- **Instant Download** - Get your generated HTML file immediately

## 📋 Requirements

- Python 3.11+
- Flask web framework
- python-docx for Word document parsing
- Jinja2 for HTML templating

## 🛠️ Local Development

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the Application

```bash
python app.py
```

The application will start at `http://localhost:5000`

### 3. Use the Application

1. Open your browser to `http://localhost:5000`
2. Upload a formatted `.docx` file
3. Upload 4 car images (front, side, rear, quarter views)
4. Click "Generate HTML Guide"
5. Download your generated HTML file

## 🌐 Deploy to Render

### Quick Deploy (Recommended)

1. **Create a GitHub Repository**
   - Push this project to GitHub
   - Make sure all files are committed

2. **Sign Up for Render**
   - Go to [render.com](https://render.com)
   - Sign up with your GitHub account (free)

3. **Create a New Web Service**
   - Click "New +" → "Web Service"
   - Connect your GitHub repository
   - Render will auto-detect the configuration from `render.yaml`

4. **Deploy**
   - Click "Create Web Service"
   - Wait 2-3 minutes for the build to complete
   - Your app will be live at `https://your-app-name.onrender.com`

### Manual Configuration (Alternative)

If auto-detection doesn't work:

1. **New Web Service Settings:**
   - **Name:** `vpg-generator` (or your choice)
   - **Environment:** `Python 3`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
   - **Instance Type:** `Free`

2. **Environment Variables:**
   - `PYTHON_VERSION` = `3.11.0`

3. Click **"Create Web Service"**

## 📁 Project Structure

```
.
├── app.py                  # Flask web application
├── generate_html.py        # Document parsing logic
├── template.html           # HTML template for output
├── requirements.txt        # Python dependencies
├── render.yaml            # Render deployment config
├── templates/
│   └── upload.html        # Upload form interface
└── README.md              # This file
```

## 🎨 How It Works

1. **Upload**: User uploads a `.docx` file and car images
2. **Parse**: `generate_html.py` extracts structured data from the document
3. **Process**: Images are converted to base64 and embedded
4. **Template**: Data is rendered using `template.html`
5. **Download**: User receives a standalone HTML file

## 🔒 File Size Limits

- Maximum file size: 16MB per upload
- Supported document format: `.docx`
- Supported image formats: `.jpg`, `.jpeg`, `.png`

## 💡 Tips

- **Document Format**: Ensure your Word document follows the expected structure
- **Image Quality**: Use clear, high-resolution images for best results
- **Free Tier**: Render's free tier may spin down after inactivity (15 min startup time)
- **Upgrade**: For production use, consider Render's paid tiers for better performance

## 🐛 Troubleshooting

### Build Fails on Render
- Check that `requirements.txt` is in the root directory
- Verify Python version compatibility
- Check build logs in Render dashboard

### Upload Errors
- Ensure file sizes are under 16MB
- Use only `.docx` format (not `.doc`)
- Provide at least one car image

### Template Not Found
- Verify `template.html` exists in root directory
- Check that `templates/upload.html` exists

## 📝 Environment Variables (Optional)

You can set these in Render dashboard:

- `FLASK_ENV` - Set to `production` for production deployment
- `MAX_CONTENT_LENGTH` - Override max file size (in bytes)

## 🆓 Cost

**100% FREE** with Render's free tier:
- 750 hours/month of runtime
- Auto-sleep after 15 min of inactivity
- Perfect for personal projects and demos

## 📧 Support

For issues or questions about deployment, check:
- [Render Documentation](https://render.com/docs)
- [Flask Documentation](https://flask.palletsprojects.com/)

## 📄 License

This project is provided as-is for personal and commercial use.

---

**Made with ❤️ for Vehicle Platform Guide generation**
