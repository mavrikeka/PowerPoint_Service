# PowerPoint Population Service - API Documentation

Complete API documentation for integrating the PowerPoint Template Population Service into your application.

---

## Quick Start (5 Minutes)

Get started with the PowerPoint Population Service in minutes:

### Step 1: Extract Template Structure (30 seconds)
```bash
# Extract data from your PowerPoint template to see its structure
curl -X POST https://createpowerpoint-development.up.railway.app/extract-data \
  -F "presentation=@your-template.pptx" \
  -F "slide_index=0" \
  > template-structure.json
```

### Step 2: Modify the Data (2 minutes)
Edit `template-structure.json` with your actual data:
```json
{
  "slide_title": "Q4 2024 Report",
  "presenter_name": "Jane Smith",
  "summary_table": [
    ["Metric", "Value"],
    ["Revenue", "$1.2M"],
    ["Growth", "23%"]
  ]
}
```

### Step 3: Generate PowerPoint (30 seconds)
```bash
# Populate the template with your data
curl -X POST https://createpowerpoint-development.up.railway.app/populate-pptx \
  -F "template=@your-template.pptx" \
  -F "data=$(cat template-structure.json)" \
  -o populated-presentation.pptx
```

**Done!** You now have a populated PowerPoint presentation.

### Copy-Paste Integration (Python)
```python
import requests
import json

# 1. Extract template structure
with open('template.pptx', 'rb') as f:
    response = requests.post(
        'https://createpowerpoint-development.up.railway.app/extract-data',
        files={'presentation': f},
        data={'slide_index': 0}
    )
    template_data = response.json()

# 2. Modify the data
template_data['slide_title'] = 'My Custom Title'
template_data['presenter_name'] = 'John Doe'

# 3. Generate populated PowerPoint
with open('template.pptx', 'rb') as f:
    response = requests.post(
        'https://createpowerpoint-development.up.railway.app/populate-pptx',
        files={'template': f},
        data={'data': json.dumps(template_data)}
    )

    with open('output.pptx', 'wb') as out:
        out.write(response.content)

print("Done! Check output.pptx")
```

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
10. [Troubleshooting](#troubleshooting)

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
https://createpowerpoint-development.up.railway.app
```

This is the live production API endpoint.

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
curl https://createpowerpoint-development.up.railway.app/health
```

---

### 2. Extract Data from PowerPoint

Extract text and table data from a PowerPoint presentation to JSON format.

**Endpoint:** `POST /extract-data`

**Content-Type:** `multipart/form-data`

#### Request Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `presentation` | File | Yes | PowerPoint file (.pptx) to extract data from |
| `slide_index` | Integer | No | Specific slide index to extract (0-based). If not provided, extracts all slides |
| `extract_all` | Boolean | No | Set to true to extract all slides (default: false) |

#### Request Example

```bash
# Extract specific slide (default: slide 0)
curl -X POST https://createpowerpoint-development.up.railway.app/extract-data \
  -F "presentation=@template.pptx" \
  -F "slide_index=0"

# Extract all slides
curl -X POST https://createpowerpoint-development.up.railway.app/extract-data \
  -F "presentation=@template.pptx" \
  -F "extract_all=true"
```

#### Response

**Success (200 OK) - Single Slide:**
```json
{
  "slide_title": "My Presentation Title",
  "role_name": "Senior Software Engineer",
  "talent_name": "John Doe",
  "risk_action_table": [
    ["Risk", "Action", "Owner", "Date"],
    ["Risk 1", "Action 1", "Owner 1", "Date 1"],
    ["Risk 2", "Action 2", "Owner 2", "Date 2"]
  ]
}
```

**Success (200 OK) - All Slides:**
```json
{
  "slides": [
    {
      "slide_index": 0,
      "data": {
        "slide_title": "Title 1",
        "content": "Content 1"
      }
    },
    {
      "slide_index": 1,
      "data": {
        "slide_title": "Title 2",
        "content": "Content 2"
      }
    }
  ]
}
```

**Error (400 Bad Request):**
```json
{
  "detail": "File must be a .pptx file"
}
```

**Use Cases:**
- Convert existing presentations to data
- Create JSON templates from real presentations
- Extract data for analysis or migration
- Generate data format for populate endpoint

---

### 3. Populate PowerPoint Template

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
curl -X POST https://createpowerpoint-development.up.railway.app/populate-pptx \
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

The `data` parameter must be a JSON string with field names matching your template's shape names. The service supports two formats:

#### Format 1: Single-Slide (Simple)

For populating a single slide:

```json
{
  "slide_title": "My Presentation Title",
  "role_name": "Senior Software Engineer",
  "talent_name": "John Doe",
  "risk_action_table": [
    ["Risk", "Action", "Owner", "Date"],
    ["Risk 1", "Action 1", "Owner 1", "Date 1"],
    ["Risk 2", "Action 2", "Owner 2", "Date 2"]
  ]
}
```

#### Format 2: Multi-Slide (Advanced)

For populating multiple slides from the same template:

```json
{
  "slides": [
    {
      "slide_index": 0,
      "data": {
        "slide_title": "First Presentation",
        "role_name": "CTO",
        "talent_name": "John Doe"
      }
    },
    {
      "slide_index": 0,
      "data": {
        "slide_title": "Second Presentation",
        "role_name": "CEO",
        "talent_name": "Jane Smith"
      }
    }
  ]
}
```

**Note:** Multi-slide format creates multiple copies of the template, each populated with different data. This is useful for generating multiple personalized presentations from one template.

#### Text Fields

For simple text placeholders, use the shape name as the key:

```json
{
  "slide_title": "My Presentation Title",
  "role_name": "Senior Software Engineer",
  "talent_name": "John Doe"
}
```

#### Table Fields

For tables, use the table shape name and provide data as a 2D array (including headers):

```json
{
  "risk_action_table": [
    ["Risk", "Action", "Owner", "Date"],
    ["Risk 1", "Action 1", "Owner 1", "Date 1"],
    ["Risk 2", "Action 2", "Owner 2", "Date 2"]
  ]
}
```

**Important Notes:**
- Table data should include the header row
- Tables no longer require a `_table` suffix (any shape name works)
- Extra rows in the template are automatically cleared
- If you provide fewer rows than the template has, empty rows will be cleared

### PowerPoint Template Requirements

Your PowerPoint template must have shapes with specific names that match your JSON keys:

1. **Text placeholders:** Name them exactly as they appear in your JSON (e.g., `slide_title`, `role_name`)
2. **Tables:** Name them to match your JSON keys (e.g., `risk_action_table`)

**To name shapes in PowerPoint:**
1. Select the shape
2. Open Selection Pane: View → Selection Pane
3. Double-click the shape name and rename it

**Best Practice - Use Extract Endpoint:**
Instead of manually inspecting your template, use the `/extract-data` endpoint to:
1. Extract shape names and structure from your template
2. Get JSON with the exact format needed for population
3. Modify the extracted values
4. Use the modified JSON with `/populate-pptx`

This ensures your JSON keys always match your template's shape names.

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
        service_url: URL of the service (e.g., "https://createpowerpoint-development.up.railway.app")
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
    service_url = "https://createpowerpoint-development.up.railway.app"
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
const serviceUrl = 'https://createpowerpoint-development.up.railway.app';
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
// populatePowerPoint(file, data, 'https://createpowerpoint-development.up.railway.app');
```

---

### cURL

```bash
# Basic usage
curl -X POST https://createpowerpoint-development.up.railway.app/populate-pptx \
  -F "template=@template.pptx" \
  -F 'data={"slide_title":"My Title","role_name":"Engineer"}' \
  -o output.pptx

# With all parameters
curl -X POST https://createpowerpoint-development.up.railway.app/populate-pptx \
  -F "template=@template.pptx" \
  -F 'data={"slide_title":"My Title","role_name":"Engineer","talent_name":"John Doe"}' \
  -F "slide_index=0" \
  -F "output_filename=my_presentation.pptx" \
  -o my_presentation.pptx \
  -w "\nStatus: %{http_code}\n"

# Check headers
curl -X POST https://createpowerpoint-development.up.railway.app/populate-pptx \
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
$serviceUrl = 'https://createpowerpoint-development.up.railway.app';
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
    serviceURL := "https://createpowerpoint-development.up.railway.app"
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

## Troubleshooting

### Common Issues and Solutions

#### 1. "Shape not found" or Fields Not Populating

**Problem:** The API returns success but some fields aren't populated.

**Cause:** Shape names in your JSON don't match the actual shape names in the PowerPoint template.

**Solution:**
```bash
# Step 1: Extract data from your template to see the exact shape names
curl -X POST https://createpowerpoint-development.up.railway.app/extract-data \
  -F "presentation=@your-template.pptx" \
  -F "slide_index=0" \
  > template-structure.json

# Step 2: Use the extracted JSON as your data template
cat template-structure.json
# This shows you the EXACT shape names to use
```

**Prevention:** Always use the `/extract-data` endpoint first to get the correct shape names.

---

#### 2. "Invalid JSON" Error

**Problem:** Getting `400 Bad Request` with "Invalid JSON" message.

**Cause:** Malformed JSON in the `data` parameter.

**Solution:**
```python
# Validate your JSON before sending
import json

data = {
    "slide_title": "My Title",
    "role_name": "Engineer"
}

# This will raise an error if JSON is invalid
json_string = json.dumps(data)
print(json_string)  # Verify it's valid
```

**Common JSON mistakes:**
- Missing quotes around keys or values
- Trailing commas
- Single quotes instead of double quotes
- Unescaped special characters

**Online validator:** Use [jsonlint.com](https://jsonlint.com) to validate your JSON.

---

#### 3. Table Data Not Showing or Overflowing

**Problem:** Table data doesn't appear or overflows cells.

**Causes:**
- Wrong table shape name
- Table cells too small for content
- Missing table data format

**Solutions:**

**a) Verify table name:**
```bash
# Extract to see table shape names
curl -X POST https://createpowerpoint-development.up.railway.app/extract-data \
  -F "presentation=@template.pptx" \
  -F "slide_index=0"
```

**b) Format table data correctly:**
```json
{
  "my_table": [
    ["Header 1", "Header 2", "Header 3"],
    ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
    ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"]
  ]
}
```

**c) Design template with appropriate cell sizes:**
- PowerPoint doesn't auto-fit table cells
- Make cells large enough for expected content
- Enable word wrap in template cells
- Keep content concise

---

#### 4. Connection Refused / Service Unavailable

**Problem:** Cannot connect to the API.

**Causes:**
- Service is down
- Wrong URL
- Network issues
- Firewall blocking

**Solutions:**

**a) Check service health:**
```bash
curl https://createpowerpoint-development.up.railway.app/health
# Should return: {"status":"healthy"}
```

**b) Verify URL:**
```python
# Production URL
SERVICE_URL = "https://createpowerpoint-development.up.railway.app"

# NOT localhost (unless testing locally)
# SERVICE_URL = "http://localhost:8000"  # Only for local testing
```

**c) Check Railway status:**
- Visit [Railway Dashboard](https://railway.app)
- Check if deployment is active
- View deployment logs for errors

**d) Test with timeout:**
```python
response = requests.post(url, data=data, timeout=30)
```

---

#### 5. File Downloaded But Won't Open

**Problem:** PowerPoint file downloads but is corrupted or won't open.

**Causes:**
- Incomplete download
- Binary data corruption
- Response not saved correctly

**Solutions:**

**a) Verify binary mode:**
```python
# CORRECT - Binary mode
with open('output.pptx', 'wb') as f:
    f.write(response.content)

# WRONG - Text mode
# with open('output.pptx', 'w') as f:
#     f.write(response.content)
```

**b) Check response status:**
```python
if response.status_code == 200:
    with open('output.pptx', 'wb') as f:
        f.write(response.content)
else:
    print(f"Error: {response.status_code}")
    print(response.text)
```

**c) Verify content type:**
```python
content_type = response.headers.get('content-type')
if 'presentation' in content_type:
    # Save file
else:
    # Error occurred, check response.text
    print(response.text)
```

---

#### 6. Multi-Slide Format Not Working

**Problem:** Multi-slide data doesn't populate correctly.

**Cause:** Incorrect JSON structure.

**Solution:**
```json
{
  "slides": [
    {
      "slide_index": 0,
      "data": {
        "slide_title": "First Slide",
        "content": "Content 1"
      }
    },
    {
      "slide_index": 1,
      "data": {
        "slide_title": "Second Slide",
        "content": "Content 2"
      }
    }
  ]
}
```

**Key points:**
- Must have top-level `"slides"` array
- Each slide must have `"slide_index"` and `"data"`
- `slide_index` is 0-based (0 = first slide)

---

#### 7. Timeout Errors

**Problem:** Request times out before completing.

**Causes:**
- Large PowerPoint files
- Complex templates
- Slow network
- Server load

**Solutions:**

**a) Increase timeout:**
```python
response = requests.post(
    url,
    files=files,
    data=data,
    timeout=60  # Increase from default 30 seconds
)
```

**b) Optimize template:**
- Remove unnecessary images
- Simplify complex graphics
- Reduce file size

**c) Use async for multiple files:**
```python
import asyncio
import httpx

async def populate_async(template_path, data):
    async with httpx.AsyncClient(timeout=60.0) as client:
        # Your code here
        pass
```

---

#### 8. Formatting Lost After Population

**Problem:** Text formatting, colors, or fonts change after population.

**Cause:** This is expected behavior. The service replaces text content, which may reset some formatting.

**Solutions:**

**a) Design template with desired formatting:**
- Set fonts, colors, sizes in the template
- The service preserves most formatting

**b) What's preserved:**
- Font family, size, color (from template)
- Background images and colors
- Slide layouts and masters
- Table structure and borders

**c) What may change:**
- Bold/italic formatting within text
- Character-level formatting
- Complex text effects

**Workaround:** Keep formatting simple in templates for consistent results.

---

### Getting Help

Still having issues?

1. **Check the Quick Start** section at the top of this document
2. **Review Code Examples** for your language
3. **Test with cURL** to isolate issues
4. **Use the Web UI** at https://createpowerpoint-development.up.railway.app for visual testing
5. **Check Railway Logs** if you're self-hosting
6. **Open a GitHub Issue** with:
   - What you're trying to do
   - The error message (if any)
   - Sample code (redact sensitive data)
   - Template structure (from `/extract-data`)

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
