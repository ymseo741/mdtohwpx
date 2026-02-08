import sys
import os
import md2hwpx

def convert_md_to_hwpx(input_file, output_file, reference_hwpx=None):
    """
    Converts a Markdown file to HWPX format.
    """
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found.")
        return

    with open(input_file, 'r', encoding='utf-8') as f:
        markdown_text = f.read()

    try:
        # If reference_hwpx is provided, use it for styling
        md2hwpx.convert_string(
            markdown_text,
            output_file,
            reference_doc=reference_hwpx
        )
        print(f"Successfully converted '{input_file}' to '{output_file}'.")
    except Exception as e:
        print(f"An error occurred during conversion: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python3 converter.py <input.md> <output.hwpx> [template.hwpx]")
    else:
        input_md = sys.argv[1]
        output_hwpx = sys.argv[2]
        template = sys.argv[3] if len(sys.argv) > 3 else None
        convert_md_to_hwpx(input_md, output_hwpx, template)
