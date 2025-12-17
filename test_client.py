#!/usr/bin/env python3
"""
Test client for the PowerPoint population service.
Demonstrates how to call the service from your application.
"""

import requests
import json
import sys


def populate_powerpoint(
    service_url: str,
    template_path: str,
    data: dict,
    output_path: str = "output.pptx",
    slide_index: int = 0
):
    """
    Call the PowerPoint population service.

    Args:
        service_url: URL of the service (e.g., "http://localhost:8000")
        template_path: Path to the template .pptx file
        data: Dictionary with field names and values to populate
        output_path: Where to save the populated file
        slide_index: Which slide to populate (default: 0)

    Returns:
        bool: True if successful, False otherwise
    """
    endpoint = f"{service_url}/populate-pptx"

    try:
        # Prepare the request
        with open(template_path, 'rb') as template_file:
            files = {
                'template': (template_path, template_file, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
            }
            form_data = {
                'data': json.dumps(data),
                'slide_index': str(slide_index),
                'output_filename': output_path
            }

            print(f"Sending request to {endpoint}...")
            print(f"Template: {template_path}")
            print(f"Data fields: {list(data.keys())}")

            # Make the request
            response = requests.post(
                endpoint,
                files=files,
                data=form_data,
                timeout=30
            )

            # Check response
            if response.status_code == 200:
                # Save the populated file
                with open(output_path, 'wb') as output_file:
                    output_file.write(response.content)

                # Print success message from headers
                message = response.headers.get('X-Population-Message', 'Success')
                print(f"\n✓ {message}")
                print(f"✓ Saved to: {output_path}")
                return True
            else:
                print(f"\n✗ Error {response.status_code}: {response.text}")
                return False

    except requests.exceptions.ConnectionError:
        print(f"\n✗ Could not connect to service at {service_url}")
        print("Make sure the service is running!")
        return False
    except FileNotFoundError:
        print(f"\n✗ Template file not found: {template_path}")
        return False
    except Exception as e:
        print(f"\n✗ Error: {str(e)}")
        return False


def main():
    """Test the service with example data."""
    # Configuration
    SERVICE_URL = "http://localhost:8000"  # Change this to your Railway URL
    TEMPLATE_PATH = "valueactionplan_template.pptx"
    OUTPUT_PATH = "service_output.pptx"

    # Example data matching the Value Action Plan format
    data = {
        'slide_title': 'VALUE ACTION PLAN TO MITIGATE RISKS',
        'role_name': 'Chief Technology Officer',
        'talent_name': 'Jane Smith',
        'risk_action_table': [
            ['1', 'Technical debt accumulation', 'Implement quarterly tech debt sprints'],
            ['2', 'Team skill gaps in cloud architecture', 'AWS certification program for senior engineers'],
            ['3', 'Legacy system dependencies', 'Create migration roadmap to microservices'],
            ['4', 'Vendor lock-in risks', 'Evaluate multi-cloud strategy options'],
        ]
    }

    print("="*60)
    print("PowerPoint Population Service - Test Client")
    print("="*60)
    print()

    # Call the service
    success = populate_powerpoint(
        service_url=SERVICE_URL,
        template_path=TEMPLATE_PATH,
        data=data,
        output_path=OUTPUT_PATH
    )

    if success:
        print("\n" + "="*60)
        print("SUCCESS!")
        print("="*60)
    else:
        print("\n" + "="*60)
        print("FAILED!")
        print("="*60)
        sys.exit(1)


if __name__ == "__main__":
    # Check if custom URL provided
    if len(sys.argv) > 1:
        SERVICE_URL = sys.argv[1]
        print(f"Using custom service URL: {SERVICE_URL}")

    main()
