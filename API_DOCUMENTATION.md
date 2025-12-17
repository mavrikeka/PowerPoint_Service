# PowerPoint Population Service - API Documentation

Complete API documentation for integrating the PowerPoint Template Population Service into your application.

---

## Table of Contents

1. [Overview](#overview)
2. [Base URL](#base-url)
3. [Authentication](#authentication)
4. [Endpoints](#endpoints)
5. [Data Formats](#data-formats)
6. [Code Examples](#code-examples)
7. [Error Handling](#error-handling)
8. [Best Practices](#best-practices)
9. [Integration Patterns](#integration-patterns)

---

## Overview

The PowerPoint Population Service is a REST API that accepts PowerPoint template files and JSON data, then returns populated PowerPoint presentations.

**Use Cases:**
- Automated report generation
- Personalized presentation creation
- Batch document processing
- Template-based content management

**Key Features:**
- No authentication required (add if needed)
- Simple REST API
- Accepts multipart/form-data
- Returns binary PowerPoint files
- Supports named placeholders and tables

---

## Base URL

### Local Development
```
http://localhost:8000
```

### Production (Railway)
```
https://your-app.up.railway.app
```

Replace `your-app.up.railway.app` with your actual Railway domain.

---

## Authentication

Currently, the API does not require authentication.

**To add authentication (recommended for production):**
- Add API key header: `X-API-Key: your-api-key`
- Or use Bearer token: `Authorization: Bearer your-token`

---

## Endpoints

### 1. Health Check

Check if the service is running.

**Endpoint:** `GET /health`

**Response:**
```json
{
  "status": "healthy"
}
```

**Example:**
```bash
curl https://your-app.up.railway.app/health
```

---

### 2. Populate PowerPoint Template

Upload a PowerPoint template and data to generate a populated presentation.

**Endpoint:** `POST /populate-pptx`

**Content-Type:** `multipart/form-data`

#### Request Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `template` | File | Yes | PowerPoint template file (.pptx) |
| `data` | String (JSON) | Yes | JSON string containing field names and values |
| `slide_index` | Integer | No | Slide index to populate (default: 0) |
| `output_filename` | String | No | Name for output file (default: "output.pptx") |

#### Request Example

```bash
curl -X POST https://your-app.up.railway.app/populate-pptx \
  -F "template=@template.pptx" \
  -F 'data={"slide_title":"My Title","role_name":"Engineer"}' \
  -o output.pptx
```

#### Response

**Success (200 OK):**
- Content-Type: `application/vnd.openxmlformats-officedocument.presentationml.presentation`
- Body: Binary PowerPoint file
- Headers:
  - `Content-Disposition: attachment; filename="output.pptx"`
  - `X-Population-Message: Successfully populated N fields: field1, field2, ...`

**Error (400 Bad Request):**
```json
{
  "detail": "Error message describing what went wrong"
}
```

**Error (500 Internal Server Error):**
```json
{
  "detail": "Internal server error message"
}
```

---

## Data Formats

### JSON Data Structure

The `data` parameter must be a JSON string with field names matching your template's shape names.

#### Text Fields

For simple text placeholders:

```json
{
  "slide_title": "My Presentation Title",
  "role_name": "Senior Software Engineer",
  "talent_name": "John Doe"
}
```

#### Table Fields

For tables, use a key ending with `_table` and provide data as a 2D array:

```json
{
  "risk_action_table": [
    ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
    ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"],
    ["Row 3 Col 1", "Row 3 Col 2", "Row 3 Col 3"]
  ]
}
```

**Note:** By default, table data skips the header row (row 0) and starts populating from row 1.

#### Complete Example

```json
{
  "slide_title": "Q4 2024 Performance Review",
  "role_name": "Chief Technology Officer",
  "talent_name": "Jane Smith",
  "risk_action_table": [
    ["1", "Technical debt accumulation", "Implement tech debt sprints", "Q1 2025"],
    ["2", "Team skill gaps", "Training program", "Q2 2025"],
    ["3", "Legacy systems", "Migration plan", "Q3 2025"]
  ]
}
```

### PowerPoint Template Requirements

Your PowerPoint template must have shapes with specific names:

1. **Text placeholders:** Name them exactly as they appear in your JSON (e.g., `slide_title`, `role_name`)
2. **Tables:** Name them with `_table` suffix (e.g., `risk_action_table`)

**To name shapes in PowerPoint:**
1. Select the shape
2. Open Selection Pane: View → Selection Pane
3. Double-click the shape name and rename it

---

## Code Examples

### Python

#### Using `requests` library

```python
import requests
import json

def populate_powerpoint(template_path, data, service_url, output_path="output.pptx"):
    """
    Populate a PowerPoint template using the API.

    Args:
        template_path: Path to the template .pptx file
        data: Dictionary with field names and values
        service_url: URL of the service (e.g., "https://your-app.up.railway.app")
        output_path: Where to save the populated file

    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Prepare the request
        with open(template_path, 'rb') as template_file:
            files = {
                'template': (template_path, template_file,
                           'application/vnd.openxmlformats-officedocument.presentationml.presentation')
            }
            form_data = {
                'data': json.dumps(data),
                'output_filename': output_path
            }

            # Make the request
            response = requests.post(
                f"{service_url}/populate-pptx",
                files=files,
                data=form_data,
                timeout=30
            )

            # Check response
            if response.status_code == 200:
                # Save the file
                with open(output_path, 'wb') as output_file:
                    output_file.write(response.content)

                # Print success message
                message = response.headers.get('X-Population-Message', 'Success')
                print(f"✓ {message}")
                print(f"✓ Saved to: {output_path}")
                return True
            else:
                print(f"✗ Error {response.status_code}: {response.text}")
                return False

    except Exception as e:
        print(f"✗ Error: {str(e)}")
        return False


# Example usage
if __name__ == "__main__":
    service_url = "https://your-app.up.railway.app"
    template_path = "template.pptx"

    data = {
        'slide_title': 'VALUE ACTION PLAN',
        'role_name': 'CTO',
        'talent_name': 'Jane Doe',
        'risk_action_table': [
            ['1', 'Technical debt', 'Implement sprints'],
            ['2', 'Skill gaps', 'Training program']
        ]
    }

    populate_powerpoint(template_path, data, service_url, "output.pptx")
```

#### Using `httpx` (async)

```python
import httpx
import json
import asyncio

async def populate_powerpoint_async(template_path, data, service_url):
    """Async version using httpx."""
    async with httpx.AsyncClient(timeout=30.0) as client:
        with open(template_path, 'rb') as f:
            files = {'template': f}
            form_data = {'data': json.dumps(data)}

            response = await client.post(
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
# asyncio.run(populate_powerpoint_async("template.pptx", data, service_url))
```

---

### JavaScript / Node.js

#### Using `axios` and `form-data`

```javascript
const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');

async function populatePowerPoint(templatePath, data, serviceUrl, outputPath = 'output.pptx') {
  try {
    // Create form data
    const form = new FormData();
    form.append('template', fs.createReadStream(templatePath));
    form.append('data', JSON.stringify(data));

    // Make request
    const response = await axios.post(`${serviceUrl}/populate-pptx`, form, {
      headers: {
        ...form.getHeaders()
      },
      responseType: 'arraybuffer'
    });

    // Save file
    fs.writeFileSync(outputPath, response.data);

    // Get message from headers
    const message = response.headers['x-population-message'] || 'Success';
    console.log(`✓ ${message}`);
    console.log(`✓ Saved to: ${outputPath}`);

    return true;
  } catch (error) {
    console.error(`✗ Error: ${error.message}`);
    if (error.response) {
      console.error(`Status: ${error.response.status}`);
      console.error(`Data: ${error.response.data}`);
    }
    return false;
  }
}

// Example usage
const serviceUrl = 'https://your-app.up.railway.app';
const templatePath = './template.pptx';

const data = {
  slide_title: 'VALUE ACTION PLAN',
  role_name: 'CTO',
  talent_name: 'Jane Doe',
  risk_action_table: [
    ['1', 'Technical debt', 'Implement sprints'],
    ['2', 'Skill gaps', 'Training program']
  ]
};

populatePowerPoint(templatePath, data, serviceUrl, 'output.pptx');
```

#### Using `fetch` (Browser/Node.js 18+)

```javascript
async function populatePowerPoint(templateFile, data, serviceUrl) {
  const formData = new FormData();
  formData.append('template', templateFile);
  formData.append('data', JSON.stringify(data));

  try {
    const response = await fetch(`${serviceUrl}/populate-pptx`, {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const blob = await response.blob();

    // Download in browser
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'output.pptx';
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);

    console.log('✓ File downloaded successfully');
    return true;
  } catch (error) {
    console.error('✗ Error:', error);
    return false;
  }
}

// Example usage in browser
// const fileInput = document.getElementById('fileInput');
// const file = fileInput.files[0];
// populatePowerPoint(file, data, 'https://your-app.up.railway.app');
```

---

### cURL

```bash
# Basic usage
curl -X POST https://your-app.up.railway.app/populate-pptx \
  -F "template=@template.pptx" \
  -F 'data={"slide_title":"My Title","role_name":"Engineer"}' \
  -o output.pptx

# With all parameters
curl -X POST https://your-app.up.railway.app/populate-pptx \
  -F "template=@template.pptx" \
  -F 'data={"slide_title":"My Title","role_name":"Engineer","talent_name":"John Doe"}' \
  -F "slide_index=0" \
  -F "output_filename=my_presentation.pptx" \
  -o my_presentation.pptx \
  -w "\nStatus: %{http_code}\n"

# Check headers
curl -X POST https://your-app.up.railway.app/populate-pptx \
  -F "template=@template.pptx" \
  -F 'data={"slide_title":"Test"}' \
  -o output.pptx \
  -D headers.txt
```

---

### PHP

```php
<?php

function populatePowerPoint($templatePath, $data, $serviceUrl, $outputPath = 'output.pptx') {
    $ch = curl_init();

    // Prepare the file
    $cfile = new CURLFile($templatePath, 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'template.pptx');

    // Prepare form data
    $postData = [
        'template' => $cfile,
        'data' => json_encode($data)
    ];

    // Set cURL options
    curl_setopt_array($ch, [
        CURLOPT_URL => $serviceUrl . '/populate-pptx',
        CURLOPT_POST => true,
        CURLOPT_POSTFIELDS => $postData,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HEADER => true
    ]);

    // Execute request
    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $headerSize = curl_getinfo($ch, CURLINFO_HEADER_SIZE);

    curl_close($ch);

    if ($httpCode === 200) {
        // Extract body (file content)
        $body = substr($response, $headerSize);

        // Save file
        file_put_contents($outputPath, $body);

        echo "✓ File saved to: {$outputPath}\n";
        return true;
    } else {
        echo "✗ Error {$httpCode}\n";
        return false;
    }
}

// Example usage
$serviceUrl = 'https://your-app.up.railway.app';
$templatePath = 'template.pptx';

$data = [
    'slide_title' => 'VALUE ACTION PLAN',
    'role_name' => 'CTO',
    'talent_name' => 'Jane Doe',
    'risk_action_table' => [
        ['1', 'Technical debt', 'Implement sprints'],
        ['2', 'Skill gaps', 'Training program']
    ]
];

populatePowerPoint($templatePath, $data, $serviceUrl, 'output.pptx');
?>
```

---

### Go

```go
package main

import (
    "bytes"
    "encoding/json"
    "fmt"
    "io"
    "mime/multipart"
    "net/http"
    "os"
)

func populatePowerPoint(templatePath string, data map[string]interface{}, serviceURL string, outputPath string) error {
    // Open template file
    file, err := os.Open(templatePath)
    if err != nil {
        return err
    }
    defer file.Close()

    // Create multipart form
    body := &bytes.Buffer{}
    writer := multipart.NewWriter(body)

    // Add template file
    part, err := writer.CreateFormFile("template", templatePath)
    if err != nil {
        return err
    }
    _, err = io.Copy(part, file)
    if err != nil {
        return err
    }

    // Add data field
    dataJSON, err := json.Marshal(data)
    if err != nil {
        return err
    }
    writer.WriteField("data", string(dataJSON))

    // Close writer
    err = writer.Close()
    if err != nil {
        return err
    }

    // Create request
    req, err := http.NewRequest("POST", serviceURL+"/populate-pptx", body)
    if err != nil {
        return err
    }
    req.Header.Set("Content-Type", writer.FormDataContentType())

    // Send request
    client := &http.Client{}
    resp, err := client.Do(req)
    if err != nil {
        return err
    }
    defer resp.Body.Close()

    // Check response
    if resp.StatusCode != 200 {
        return fmt.Errorf("server returned status %d", resp.StatusCode)
    }

    // Save output file
    outFile, err := os.Create(outputPath)
    if err != nil {
        return err
    }
    defer outFile.Close()

    _, err = io.Copy(outFile, resp.Body)
    if err != nil {
        return err
    }

    fmt.Printf("✓ File saved to: %s\n", outputPath)
    return nil
}

func main() {
    serviceURL := "https://your-app.up.railway.app"
    templatePath := "template.pptx"

    data := map[string]interface{}{
        "slide_title": "VALUE ACTION PLAN",
        "role_name":   "CTO",
        "talent_name": "Jane Doe",
        "risk_action_table": [][]string{
            {"1", "Technical debt", "Implement sprints"},
            {"2", "Skill gaps", "Training program"},
        },
    }

    err := populatePowerPoint(templatePath, data, serviceURL, "output.pptx")
    if err != nil {
        fmt.Printf("✗ Error: %v\n", err)
    }
}
```

---

## Error Handling

### Common Errors

#### 400 Bad Request

**Causes:**
- Invalid JSON in `data` parameter
- Template file is not a valid .pptx file
- Required shape names not found in template
- Template has no slides

**Example Error:**
```json
{
  "detail": "Invalid JSON data: Expecting value: line 1 column 1 (char 0)"
}
```

**Solution:** Validate your JSON before sending.

---

#### 500 Internal Server Error

**Causes:**
- Server error processing the template
- File I/O errors
- Corrupted template file

**Example Error:**
```json
{
  "detail": "Error processing template: 'NoneType' object has no attribute 'text'"
}
```

**Solution:** Check template file integrity and shape names.

---

#### Network Errors

**Causes:**
- Service is down
- Network connectivity issues
- Timeout

**Solution:**
- Implement retry logic
- Add timeout handling
- Check service health endpoint first

---

### Error Handling Pattern

```python
import requests
import time

def populate_with_retry(template_path, data, service_url, max_retries=3):
    """Populate PowerPoint with retry logic."""
    for attempt in range(max_retries):
        try:
            # Check if service is healthy
            health_response = requests.get(f"{service_url}/health", timeout=5)
            if health_response.status_code != 200:
                raise Exception("Service is not healthy")

            # Make request
            with open(template_path, 'rb') as f:
                files = {'template': f}
                form_data = {'data': json.dumps(data)}

                response = requests.post(
                    f"{service_url}/populate-pptx",
                    files=files,
                    data=form_data,
                    timeout=30
                )

                if response.status_code == 200:
                    return response.content
                elif response.status_code == 400:
                    # Client error - don't retry
                    raise Exception(f"Client error: {response.text}")
                else:
                    # Server error - retry
                    raise Exception(f"Server error: {response.status_code}")

        except Exception as e:
            if attempt == max_retries - 1:
                raise
            print(f"Attempt {attempt + 1} failed: {e}. Retrying...")
            time.sleep(2 ** attempt)  # Exponential backoff

    raise Exception("Max retries exceeded")
```

---

## Best Practices

### 1. Validate JSON Before Sending

```python
import json

def validate_data(data):
    """Validate data structure."""
    try:
        # Ensure it's valid JSON
        json.dumps(data)

        # Check required fields
        required_fields = ['slide_title', 'role_name', 'talent_name']
        for field in required_fields:
            if field not in data:
                raise ValueError(f"Missing required field: {field}")

        return True
    except Exception as e:
        print(f"Validation error: {e}")
        return False
```

### 2. Use Timeouts

Always set reasonable timeouts to prevent hanging requests:

```python
response = requests.post(url, data=data, timeout=30)  # 30 second timeout
```

### 3. Handle Large Files

For large templates or batch processing:

```python
# Use streaming for large files
with requests.post(url, data=data, stream=True) as response:
    with open('output.pptx', 'wb') as f:
        for chunk in response.iter_content(chunk_size=8192):
            f.write(chunk)
```

### 4. Template Naming Convention

Use consistent naming in your templates:
- `{field_name}` for text fields
- `{field_name}_table` for tables
- Lowercase with underscores

### 5. Security

For production:
- Add API key authentication
- Use HTTPS only
- Implement rate limiting
- Validate file sizes
- Scan uploaded files for malware

---

## Integration Patterns

### Pattern 1: Simple Synchronous

Best for: Single document generation, low volume

```python
def generate_report(template_id, user_data):
    """Simple synchronous generation."""
    template = get_template(template_id)
    data = prepare_data(user_data)

    result = populate_powerpoint(template, data, SERVICE_URL)
    return result
```

### Pattern 2: Queue-Based Async

Best for: High volume, batch processing

```python
from celery import Celery

app = Celery('tasks', broker='redis://localhost:6379')

@app.task
def generate_presentation_async(template_path, data):
    """Generate presentation asynchronously."""
    result = populate_powerpoint(template_path, data, SERVICE_URL)
    # Store result in S3, send email, etc.
    return result

# Usage
task = generate_presentation_async.delay(template, data)
```

### Pattern 3: Webhook Callback

Best for: Long-running processes, multiple steps

```python
def generate_with_callback(template, data, callback_url):
    """Generate and notify via webhook."""
    try:
        result = populate_powerpoint(template, data, SERVICE_URL)

        # Save to cloud storage
        url = upload_to_s3(result)

        # Notify via webhook
        requests.post(callback_url, json={
            'status': 'success',
            'url': url
        })
    except Exception as e:
        requests.post(callback_url, json={
            'status': 'error',
            'message': str(e)
        })
```

### Pattern 4: Caching

Best for: Same template, different data

```python
import hashlib
from functools import lru_cache

@lru_cache(maxsize=100)
def get_cached_template(template_hash):
    """Cache template processing."""
    return load_template(template_hash)

def generate_with_cache(template_path, data):
    # Hash template for caching
    with open(template_path, 'rb') as f:
        template_hash = hashlib.md5(f.read()).hexdigest()

    template = get_cached_template(template_hash)
    return populate_powerpoint(template, data, SERVICE_URL)
```

---

## Support & Troubleshooting

### Debug Mode

Enable verbose logging to troubleshoot:

```python
import logging

logging.basicConfig(level=logging.DEBUG)

# Your code here
```

### Common Issues

1. **"Shape not found"** - Check shape names in PowerPoint Selection Pane
2. **"Invalid JSON"** - Validate JSON with jsonlint.com
3. **"Timeout"** - Increase timeout or check network
4. **"File corrupt"** - Verify template file integrity

### Testing

Use the included test UI (`index.html`) to test your service before integration.

---

## Changelog & Versioning

**Version 1.0.0**
- Initial release
- Basic text and table population
- Health check endpoint

---

## License

This API documentation is provided as-is for integration purposes.

---

**Need help?** Open an issue on GitHub or contact support.
