"""
Streamlit app for generating product catalog PowerPoints.

Usage:
    streamlit run app.py
"""

import os
import glob
import shutil
import tempfile
import zipfile
import streamlit as st
from generate_catalog import generate_catalog

# ── Config ───────────────────────────────────────────────────────────────────

TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "templates")
FONTS_DIR = os.path.join(os.path.dirname(__file__), "fonts")


def install_fonts():
    """Install custom fonts from the fonts/ directory into the system font path."""
    system_font_dir = os.path.expanduser("~/.local/share/fonts")
    os.makedirs(system_font_dir, exist_ok=True)
    installed = False
    for ttf in glob.glob(os.path.join(FONTS_DIR, "*.ttf")):
        dest = os.path.join(system_font_dir, os.path.basename(ttf))
        if not os.path.exists(dest):
            shutil.copy2(ttf, dest)
            installed = True
    if installed:
        os.system("fc-cache -f")


def get_templates():
    """Return a dict of {display_name: filepath} for all .pptx files in templates/."""
    templates = {}
    if os.path.isdir(TEMPLATES_DIR):
        for f in sorted(os.listdir(TEMPLATES_DIR)):
            if f.endswith(".pptx"):
                display = f.replace(".pptx", "").replace("_", " ").title()
                templates[display] = os.path.join(TEMPLATES_DIR, f)
    return templates


def extract_images_zip(zip_file, dest_dir):
    """Extract uploaded zip to dest_dir. Handles nested folder structures."""
    with zipfile.ZipFile(zip_file, "r") as zf:
        zf.extractall(dest_dir)

    # If the zip contained a single top-level folder, move its contents up
    # e.g. images.zip/images/H-10078#1/... → dest_dir/H-10078#1/...
    entries = [e for e in os.listdir(dest_dir) if not e.startswith("__MACOSX")]
    if len(entries) == 1:
        inner = os.path.join(dest_dir, entries[0])
        if os.path.isdir(inner):
            for item in os.listdir(inner):
                shutil.move(os.path.join(inner, item), os.path.join(dest_dir, item))
            os.rmdir(inner)


# ── App ──────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="Product Catalog Generator", layout="centered")

    # Install fonts on first run
    if "fonts_installed" not in st.session_state:
        install_fonts()
        st.session_state.fonts_installed = True

    st.title("Product Catalog Generator")
    st.markdown("Generate a product catalog PowerPoint from your Excel data and product images.")

    # ── Template selector ────────────────────────────────────────────────
    templates = get_templates()
    if not templates:
        st.error(f"No .pptx templates found in `{TEMPLATES_DIR}/`. Add at least one template.")
        return

    template_name = st.selectbox("PPT Template", list(templates.keys()))
    template_path = templates[template_name]

    # ── File uploads ─────────────────────────────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        excel_file = st.file_uploader("Excel file", type=["xlsx", "xls"])
    with col2:
        images_zip = st.file_uploader("Images (.zip)", type=["zip"])

    # ── Generate ─────────────────────────────────────────────────────────
    if st.button("Generate Catalog", type="primary", disabled=not (excel_file and images_zip)):
        with st.spinner("Generating catalog..."):
            tmpdir = tempfile.mkdtemp()
            try:
                # Save uploaded excel
                excel_path = os.path.join(tmpdir, "data.xlsx")
                with open(excel_path, "wb") as f:
                    f.write(excel_file.getvalue())

                # Extract images zip
                images_dir = os.path.join(tmpdir, "images")
                os.makedirs(images_dir)
                zip_path = os.path.join(tmpdir, "images.zip")
                with open(zip_path, "wb") as f:
                    f.write(images_zip.getvalue())
                extract_images_zip(zip_path, images_dir)

                # Generate
                output_path = os.path.join(tmpdir, "catalog_output.pptx")
                summary = generate_catalog(excel_path, template_path, images_dir, output_path)

                # Show summary
                c1, c2, c3 = st.columns(3)
                c1.metric("Excel Rows", summary["total_rows"])
                c2.metric("Slides Created", summary["slides_created"])
                c3.metric("Images Found", f"{summary['images_found']}/{summary['total_rows']} ({summary['images_pct']}%)")

                if summary["missing_folders"]:
                    with st.expander("Missing image folders"):
                        for pid in summary["missing_folders"]:
                            st.text(f"  • {pid}")

                # Download button
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="Download Catalog (.pptx)",
                        data=f.read(),
                        file_name="product_catalog.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                    )

                st.success("Catalog generated successfully!")

            except Exception as e:
                st.error(f"Error generating catalog: {e}")
                raise
            finally:
                shutil.rmtree(tmpdir, ignore_errors=True)

    # ── Help ─────────────────────────────────────────────────────────────
    with st.expander("Expected file formats"):
        st.markdown("""
**Excel columns** (case-insensitive):
`product_id`, `length`, `width`, `height`, `FOB Port`, `Price`, `Inner carton`, `Outer carton`, `Unit`, `Cbm`

Only `product_id` is required. Other columns are included in the table if present.

**Images zip** structure:
```
images.zip
├── H-10078#1/
│   ├── IMG_2731.png      (main image, prefix "IMG")
│   ├── circle1_detail.png
│   └── circle2_detail.png
├── H-10078#2/
│   ├── IMG_0001.jpg
│   ├── circle1.png
│   ├── circle2.png
│   ├── circle3.png
│   └── circle4.png
└── ...
```
Folder names must match `product_id` values in the Excel file.
        """)


if __name__ == "__main__":
    main()
