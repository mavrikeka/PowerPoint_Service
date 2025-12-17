#!/usr/bin/env python3
"""
FastAPI service for populating PowerPoint templates.
Upload a template + data, receive a populated PowerPoint file.
"""

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import Response, FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
import json
import tempfile
import os
from pathlib import Path
from typing import Optional
import shutil

app = FastAPI(
    title="PowerPoint Population Service",
    description="Upload a PowerPoint template and data to generate a populated presentation",
    version="1.0.0"
)

# Add CORS middleware to allow requests from your application
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Configure this for production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def find_shape_by_name(slide, name):
    """Find a shape by its name in a slide."""
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def populate_text_placeholder(shape, text):
    """Populate a text placeholder with the given text."""
    if shape is None:
        return False

    if hasattr(shape, "text_frame"):
        shape.text_frame.text = text
        return True
    elif hasattr(shape, "text"):
        shape.text = text
        return True

    return False


def populate_table(table_shape, data, skip_header=True):
    """
    Populate a table with data.

    Args:
        table_shape: The shape containing the table
        data: List of lists, where each inner list is a row
        skip_header: If True, start populating from row 1 (preserves header row)
    """
    if not table_shape.has_table:
        return False

    table = table_shape.table
    start_row = 1 if skip_header else 0

    # Populate the table
    for data_idx, row_data in enumerate(data):
        table_row_idx = data_idx + start_row
        if table_row_idx >= len(table.rows):
            break

        for col_idx, cell_value in enumerate(row_data):
            if col_idx >= len(table.columns):
                break

            cell = table.cell(table_row_idx, col_idx)
            cell.text = str(cell_value)

    return True


def populate_presentation_from_data(template_file, output_file, data: dict, slide_index: int = 0):
    """
    Populate a PowerPoint template with data.

    Args:
        template_file: Path to the template PPTX file
        output_file: Path to save the populated PPTX file
        data: Dictionary containing the data to populate
        slide_index: Which slide to populate (default: 0 for first slide)

    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        # Load the presentation
        prs = Presentation(template_file)

        if len(prs.slides) == 0:
            return False, "Template has no slides"

        if slide_index >= len(prs.slides):
            return False, f"Slide index {slide_index} out of range (template has {len(prs.slides)} slides)"

        slide = prs.slides[slide_index]
        populated_fields = []

        # Populate text placeholders
        for key, value in data.items():
            # Skip table data for now
            if key.endswith('_table') or isinstance(value, list):
                continue

            shape = find_shape_by_name(slide, key)
            if shape:
                success = populate_text_placeholder(shape, str(value))
                if success:
                    populated_fields.append(key)

        # Populate tables
        for key, value in data.items():
            if key.endswith('_table') and isinstance(value, list):
                table_shape = find_shape_by_name(slide, key)
                if table_shape:
                    # Check if first row should be treated as header
                    skip_header = data.get(f"{key}_skip_header", True)
                    success = populate_table(table_shape, value, skip_header=skip_header)
                    if success:
                        populated_fields.append(key)

        # Save the presentation
        prs.save(output_file)

        message = f"Successfully populated {len(populated_fields)} fields: {', '.join(populated_fields)}"
        return True, message

    except Exception as e:
        return False, f"Error processing template: {str(e)}"


@app.get("/")
async def root():
    """Serve the test UI."""
    index_path = Path(__file__).parent / "index.html"
    if index_path.exists():
        return FileResponse(index_path, media_type="text/html")
    else:
        return HTMLResponse("""
        <html>
            <body>
                <h1>PowerPoint Population Service</h1>
                <p>Service is running. API endpoint: POST /populate-pptx</p>
                <p>For API documentation, see <a href="https://github.com/CEO-Works/Create_Powerpoint">GitHub</a></p>
            </body>
        </html>
        """)


@app.get("/health")
async def health():
    """Health check for Railway."""
    return {"status": "healthy"}


@app.post("/populate-pptx")
async def populate_pptx(
    template: UploadFile = File(..., description="PowerPoint template file (.pptx)"),
    data: str = Form(..., description="JSON string containing field names and values to populate"),
    slide_index: Optional[int] = Form(0, description="Slide index to populate (default: 0)"),
    output_filename: Optional[str] = Form("output.pptx", description="Name for the output file")
):
    """
    Populate a PowerPoint template with data.

    Example data format:
    {
        "slide_title": "My Title",
        "role_name": "Software Engineer",
        "talent_name": "John Doe",
        "risk_action_table": [
            ["Risk 1", "Action 1", "Owner 1", "Date 1"],
            ["Risk 2", "Action 2", "Owner 2", "Date 2"]
        ]
    }
    """
    # Validate template file
    if not template.filename.endswith('.pptx'):
        raise HTTPException(status_code=400, detail="Template must be a .pptx file")

    # Parse JSON data
    try:
        data_dict = json.loads(data)
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=400, detail=f"Invalid JSON data: {str(e)}")

    # Create temporary directory for processing
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir_path = Path(temp_dir)

        # Save uploaded template
        template_path = temp_dir_path / "template.pptx"
        with open(template_path, "wb") as f:
            shutil.copyfileobj(template.file, f)

        # Generate output path
        output_path = temp_dir_path / "output.pptx"

        # Populate the presentation
        success, message = populate_presentation_from_data(
            str(template_path),
            str(output_path),
            data_dict,
            slide_index
        )

        if not success:
            raise HTTPException(status_code=400, detail=message)

        # Read the file into memory before temp dir is cleaned up
        with open(output_path, "rb") as f:
            file_content = f.read()

        # Return the populated file
        return Response(
            content=file_content,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f'attachment; filename="{output_filename}"',
                "X-Population-Message": message
            }
        )


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
