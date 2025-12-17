#!/usr/bin/env python3
"""
Test script for multi-slide population functionality.
"""

import requests
import json
from pathlib import Path

# Configuration
SERVICE_URL = "http://localhost:8001/populate-pptx"
TEMPLATE_FILE = "valueactionplan_template.pptx"
OUTPUT_FILE = "multi_slide_output.pptx"

def test_multi_slide():
    """Test multi-slide population with Option 1 format."""

    # Verify template exists
    template_path = Path(TEMPLATE_FILE)
    if not template_path.exists():
        print(f"❌ Template file not found: {TEMPLATE_FILE}")
        return False

    # Create test data in Option 1 format (multi-slide)
    test_data = {
        "slides": [
            {
                "slide_index": 0,
                "data": {
                    "slide_title": "Q4 Value Action Plan - Product Team",
                    "role_name": "Senior Software Engineer",
                    "talent_name": "Alice Johnson",
                    "risk_action_table": [
                        ["Technical debt", "Refactor legacy code", "Alice J.", "2024-01-15"],
                        ["Performance issues", "Optimize database queries", "Bob K.", "2024-01-20"],
                        ["Security gaps", "Update dependencies", "Carol M.", "2024-01-10"]
                    ]
                }
            },
            {
                "slide_index": 0,
                "data": {
                    "slide_title": "Q4 Value Action Plan - Engineering Team",
                    "role_name": "DevOps Engineer",
                    "talent_name": "Bob Kumar",
                    "risk_action_table": [
                        ["Infrastructure costs", "Migrate to cloud", "David N.", "2024-02-01"],
                        ["Deployment delays", "Automate CI/CD", "Emma P.", "2024-01-25"],
                        ["Monitoring gaps", "Implement observability", "Frank Q.", "2024-01-30"]
                    ]
                }
            },
            {
                "slide_index": 0,
                "data": {
                    "slide_title": "Q4 Value Action Plan - Design Team",
                    "role_name": "UX Designer",
                    "talent_name": "Carol Martinez",
                    "risk_action_table": [
                        ["User confusion", "Redesign navigation", "Grace R.", "2024-01-12"],
                        ["Accessibility issues", "WCAG 2.1 compliance", "Henry S.", "2024-01-18"],
                        ["Mobile experience", "Responsive design update", "Irene T.", "2024-01-22"]
                    ]
                }
            }
        ]
    }

    print("🧪 Testing multi-slide population...")
    print(f"📄 Template: {TEMPLATE_FILE}")
    print(f"📊 Creating {len(test_data['slides'])} slides")

    # Prepare the request
    files = {
        'template': (TEMPLATE_FILE, open(template_path, 'rb'), 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
    }

    data = {
        'data': json.dumps(test_data),
        'output_filename': OUTPUT_FILE
    }

    try:
        # Make the request
        print("\n📤 Sending request to service...")
        response = requests.post(SERVICE_URL, files=files, data=data)

        if response.status_code == 200:
            # Save the output file
            output_path = Path(OUTPUT_FILE)
            with open(output_path, 'wb') as f:
                f.write(response.content)

            # Get the population message from headers
            message = response.headers.get('X-Population-Message', 'No message')

            print(f"\n✅ Success!")
            print(f"📝 Message: {message}")
            print(f"💾 Output saved to: {OUTPUT_FILE}")
            print(f"📏 File size: {len(response.content):,} bytes")
            return True
        else:
            print(f"\n❌ Error: {response.status_code}")
            print(f"📝 Details: {response.text}")
            return False

    except requests.exceptions.ConnectionError:
        print(f"\n❌ Could not connect to service at {SERVICE_URL}")
        print("💡 Make sure the service is running: PORT=8001 python service.py")
        return False
    except Exception as e:
        print(f"\n❌ Unexpected error: {str(e)}")
        return False
    finally:
        # Close the file
        files['template'][1].close()


def test_single_slide_backward_compat():
    """Test that single-slide format still works (backward compatibility)."""

    # Verify template exists
    template_path = Path(TEMPLATE_FILE)
    if not template_path.exists():
        print(f"❌ Template file not found: {TEMPLATE_FILE}")
        return False

    # Create test data in original single-slide format
    test_data = {
        "slide_title": "Q4 Value Action Plan - Legacy Format Test",
        "role_name": "Test Engineer",
        "talent_name": "Test User",
        "risk_action_table": [
            ["Risk 1", "Action 1", "Owner 1", "2024-01-01"],
            ["Risk 2", "Action 2", "Owner 2", "2024-01-02"]
        ]
    }

    print("\n\n🧪 Testing single-slide format (backward compatibility)...")
    print(f"📄 Template: {TEMPLATE_FILE}")

    # Prepare the request
    files = {
        'template': (TEMPLATE_FILE, open(template_path, 'rb'), 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
    }

    data = {
        'data': json.dumps(test_data),
        'output_filename': 'single_slide_output.pptx'
    }

    try:
        # Make the request
        print("📤 Sending request to service...")
        response = requests.post(SERVICE_URL, files=files, data=data)

        if response.status_code == 200:
            # Save the output file
            output_path = Path('single_slide_output.pptx')
            with open(output_path, 'wb') as f:
                f.write(response.content)

            # Get the population message from headers
            message = response.headers.get('X-Population-Message', 'No message')

            print(f"\n✅ Success!")
            print(f"📝 Message: {message}")
            print(f"💾 Output saved to: single_slide_output.pptx")
            return True
        else:
            print(f"\n❌ Error: {response.status_code}")
            print(f"📝 Details: {response.text}")
            return False

    except requests.exceptions.ConnectionError:
        print(f"\n❌ Could not connect to service at {SERVICE_URL}")
        print("💡 Make sure the service is running: PORT=8001 python service.py")
        return False
    except Exception as e:
        print(f"\n❌ Unexpected error: {str(e)}")
        return False
    finally:
        # Close the file
        files['template'][1].close()


if __name__ == "__main__":
    print("=" * 70)
    print("Multi-Slide Population Test Suite")
    print("=" * 70)

    # Test multi-slide
    multi_result = test_multi_slide()

    # Test single-slide backward compatibility
    single_result = test_single_slide_backward_compat()

    # Summary
    print("\n" + "=" * 70)
    print("Test Summary")
    print("=" * 70)
    print(f"Multi-slide test:  {'✅ PASSED' if multi_result else '❌ FAILED'}")
    print(f"Single-slide test: {'✅ PASSED' if single_result else '❌ FAILED'}")
    print("=" * 70)

    if multi_result and single_result:
        print("\n🎉 All tests passed!")
    else:
        print("\n⚠️  Some tests failed. Please review the errors above.")
