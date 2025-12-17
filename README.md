# PowerPoint Template Populator

A Python project to populate PowerPoint templates with data programmatically.

**Two ways to use this project:**
1. **Local Scripts** - Run Python scripts directly on your machine
2. **Web Service** - Deploy as an API service on Railway and call it from any application

---

## Option 1: Web Service (Recommended for Production)

Deploy this as a REST API service that accepts templates and data, returns populated PowerPoint files.

### Deploy to Railway from GitHub

**Step 1: Push to GitHub**

```bash
# Initialize git repository (if not already done)
git init

# Add all files
git add .

# Create initial commit
git commit -m "Initial commit: PowerPoint population service"

# Add your GitHub repository as remote
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git

# Push to GitHub
git push -u origin main
```

**Step 2: Deploy on Railway**

1. Go to [Railway.app](https://railway.app) and sign in
2. Click **"New Project"**
3. Select **"Deploy from GitHub repo"**
4. Choose your repository: `YOUR_USERNAME/YOUR_REPO`
5. Railway will automatically:
   - Detect the `Dockerfile`
   - Build the Docker image
   - Deploy the service
6. Click **"Settings"** → **"Generate Domain"** to get your public URL

Your service will be available at `https://your-app.up.railway.app`

**Environment Variables** (optional):
- Railway automatically sets the `PORT` environment variable
- No additional configuration needed!

**Automatic Deployments**:
- Every push to `main` branch will automatically redeploy
- Railway monitors your GitHub repo for changes

---

### Test the Service

#### Web UI (For Testing)

A simple web interface is included to test the service visually.

**To use the test UI:**

1. **Start the service locally:**
   ```bash
   python service.py
   ```

2. **Open `index.html` in your browser:**
   - Double-click `index.html`, or
   - Open it with: `open index.html` (macOS) / `start index.html` (Windows)

3. **Use the interface:**
   - Enter service URL (e.g., `http://localhost:8000` or your Railway URL)
   - Upload your PowerPoint template
   - Enter JSON data
   - Click "Generate PowerPoint"
   - File will download automatically

**To deploy the UI:**
- Host `index.html` on any static hosting (Netlify, Vercel, GitHub Pages)
- Update the service URL to your Railway URL
- Share with your team for testing

---

### API Endpoint

**POST /populate-pptx**

Upload a PowerPoint template and data to generate a populated presentation.

**Request:**
- Method: `POST`
- Content-Type: `multipart/form-data`
- Body:
  - `template` (file): Your .pptx template file
  - `data` (string): JSON string with field names and values
  - `slide_index` (optional, number): Slide to populate (default: 0)
  - `output_filename` (optional, string): Name for output file (default: "output.pptx")

**Response:**
- Content-Type: `application/vnd.openxmlformats-officedocument.presentationml.presentation`
- Body: The populated .pptx file

**Example Data Format:**
```json
{
  "slide_title": "VALUE ACTION PLAN",
  "role_name": "Chief Marketing Officer",
  "talent_name": "John Smith",
  "risk_action_table": [
    ["1", "Risk description", "Action to take"],
    ["2", "Another risk", "Another action"]
  ]
}
```

### Testing the Service Locally

1. **Install service dependencies**:
```bash
pip install -r requirements-service.txt
```

2. **Run the service**:
```bash
python service.py
```

Service will start at `http://localhost:8000`

3. **Test with the client script**:
```bash
python test_client.py
```

Or test with a deployed service:
```bash
python test_client.py https://your-app.up.railway.app
```

### Calling from Your Application

**Python Example:**
```python
import requests
import json

def populate_ppt(template_path, data_dict, service_url):
    with open(template_path, 'rb') as f:
        files = {'template': f}
        form_data = {'data': json.dumps(data_dict)}

        response = requests.post(
            f"{service_url}/populate-pptx",
            files=files,
            data=form_data
        )

        if response.status_code == 200:
            with open('output.pptx', 'wb') as out:
                out.write(response.content)
            return True
    return False

# Usage
data = {
    'slide_title': 'My Title',
    'role_name': 'Engineer',
    'talent_name': 'Jane Doe',
    'risk_action_table': [['Risk', 'Action']]
}
populate_ppt('template.pptx', data, 'https://your-app.up.railway.app')
```

**JavaScript/Node.js Example:**
```javascript
const FormData = require('form-data');
const fs = require('fs');
const axios = require('axios');

async function populatePPT(templatePath, data, serviceUrl) {
  const form = new FormData();
  form.append('template', fs.createReadStream(templatePath));
  form.append('data', JSON.stringify(data));

  const response = await axios.post(`${serviceUrl}/populate-pptx`, form, {
    headers: form.getHeaders(),
    responseType: 'arraybuffer'
  });

  fs.writeFileSync('output.pptx', response.data);
}

// Usage
const data = {
  slide_title: 'My Title',
  role_name: 'Engineer',
  talent_name: 'Jane Doe',
  risk_action_table: [['Risk', 'Action']]
};
populatePPT('template.pptx', data, 'https://your-app.up.railway.app');
```

**cURL Example:**
```bash
curl -X POST https://your-app.up.railway.app/populate-pptx \
  -F "template=@valueactionplan.pptx" \
  -F 'data={"slide_title":"My Title","role_name":"Engineer"}' \
  -o output.pptx
```

### Complete API Documentation

**📖 For comprehensive integration guide, see [API_DOCUMENTATION.md](API_DOCUMENTATION.md)**

The complete API documentation includes:
- ✅ Detailed endpoint specifications
- ✅ Request/response formats
- ✅ Error handling patterns
- ✅ Code examples in **Python, JavaScript, PHP, Go, cURL**
- ✅ Integration patterns (sync, async, queue-based, webhooks)
- ✅ Best practices and troubleshooting
- ✅ Ready to give to Claude Code or developers for integration

**Quick start for developers:**
1. Read [API_DOCUMENTATION.md](API_DOCUMENTATION.md)
2. Copy the code example for your language
3. Replace the service URL with your Railway URL
4. Start integrating!

---

## Option 2: Local Scripts

Run the scripts directly on your machine for testing or one-off usage.

### Setup

1. Create and activate a virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

### Usage

### 1. Debug Script - Inspect Template Shapes

First, use the debug script to identify all shape names in your PowerPoint template:

```bash
python debug_shapes.py valueactionplan.pptx
```

This will print detailed information about all shapes in each slide, including:
- Shape names
- Shape types
- Whether they are placeholders
- Whether they are tables
- Text content (if any)

### 2. Main Script - Populate Template

Once you've identified the shape names, use the main script to populate the template:

```bash
python populate_ppt.py
```

This will:
- Read the `valueactionplan.pptx` template
- Populate the named placeholders:
  - `slide_title` - The main slide title
  - `role_name` - The role/position name
  - `talent_name` - The person's name
  - `risk_action_table` - A table with risk and action items
- Save the result as `valueactionplan_populated.pptx`

### Customizing Data

Edit the `data` dictionary in `populate_ppt.py` to customize the content:

```python
data = {
    'slide_title': 'Your Title Here',
    'role_name': 'Your Role',
    'talent_name': 'Your Name',
    'risk_action_table': [
        ['Header 1', 'Header 2', 'Header 3'],
        ['Row 1 Col 1', 'Row 1 Col 2', 'Row 1 Col 3'],
        ['Row 2 Col 1', 'Row 2 Col 2', 'Row 2 Col 3'],
    ]
}
```

## Template Requirements

Your PowerPoint template should have shapes named:
- `slide_title` - A text placeholder for the slide title
- `role_name` - A text placeholder for the role name
- `talent_name` - A text placeholder for the talent name
- `risk_action_table` - A table shape for risk and action items

To set shape names in PowerPoint:
1. Select the shape
2. Right-click and choose "Edit Alt Text" or use the Selection Pane
3. The name appears in the Selection Pane (View > Selection Pane)

## Project Structure

```
GeneratePPT/
├── venv/                          # Virtual environment (not in git)
│
├── Web Service Files:
├── service.py                     # FastAPI web service
├── requirements-service.txt       # Service dependencies
├── Dockerfile                     # Docker container config for Railway
├── railway.json                   # Railway deployment configuration
├── test_client.py                 # Example Python client for testing service
├── index.html                     # 🌐 Web UI for testing the service
│
├── Documentation:
├── README.md                      # This file - Quick start guide
├── API_DOCUMENTATION.md           # 📖 Complete API integration guide
│
├── Local Script Files:
├── debug_shapes.py                # Debug script to inspect templates
├── populate_ppt.py                # Main script to populate templates locally
├── create_sample_template.py      # Creates a sample template with named shapes
├── requirements.txt               # Local script dependencies
│
├── Template Files:
├── valueactionplan_template.pptx  # Sample template with named shapes
│
└── .gitignore                     # Git ignore configuration
```

**Total:** 14 essential files ready for deployment

## Notes

- The scripts currently work with the first slide of the presentation
- You can modify the scripts to work with multiple slides
- The table population expects data as a list of lists (rows and columns)
- Make sure shape names in your template match the names used in the script
