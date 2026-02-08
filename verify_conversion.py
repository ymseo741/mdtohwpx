import os
import zipfile
import xml.etree.ElementTree as ET
import md2hwpx

SAMPLE_MD = """
# Test Table and Diagram

| Col 1 | Col 2 |
| :--- | :--- |
| Item 1 | Detail 1 |
| Item 2 | Detail 2 |

```mermaid
graph TD
    A --> B
```
"""

OUTPUT_FILE = "test_output.hwpx"

def verify():
    print("Starting conversion...")
    md2hwpx.convert_string(SAMPLE_MD, OUTPUT_FILE)
    
    if not os.path.exists(OUTPUT_FILE):
        print("Error: Output file not generated.")
        return

    print(f"File generated: {OUTPUT_FILE}")
    
    # Check contents of HWPX (ZIP)
    with zipfile.ZipFile(OUTPUT_FILE, 'r') as z:
        section0 = z.read('Contents/section0.xml').decode('utf-8')
        
    print("\n--- Verifying Tables ---")
    if '<hp:tbl' in section0:
        print("SUCCESS: Table found in section0.xml")
    else:
        print("FAILURE: No table found in section0.xml")

    print("\n--- Verifying Diagrams (Mermaid) ---")
    if 'graph TD' in section0:
        print("Note: Mermaid code found as text (expected current behavior).")
    
    if '<hp:pic' in section0 or '<hp:img' in section0:
        print("SUCCESS: Image/Picture found (unexpected for current version).")
    else:
        print("FAILURE: No image found for Mermaid diagram.")

if __name__ == "__main__":
    verify()
