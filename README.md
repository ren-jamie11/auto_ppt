# Product Catalog Generator

A Streamlit app that generates product catalog PowerPoint presentations from Excel data and product images.

## Repo Structure

```
catalog-app/
├── app.py                  # Streamlit interface
├── generate_catalog.py     # Core PPT generation logic
├── requirements.txt        # Python dependencies
├── packages.txt            # System packages (Streamlit Cloud)
├── .streamlit/
│   └── config.toml         # Streamlit theme config
├── templates/              # PPT templates (.pptx files)
│   └── sales_template_table.pptx
├── fonts/                  # Custom fonts (.ttf files)
│   └── GeorgiaProLight.ttf # ← YOU MUST ADD THIS
└── README.md
```

## Setup

### 1. Add your font file

Place your `Georgia Pro Light` `.ttf` file in the `fonts/` directory. The app installs it automatically at runtime.

### 2. Add your PPT template(s)

Place `.pptx` template files in the `templates/` directory. They'll appear in the template dropdown. Each template must have:
- Slide 1: cover slide (kept as-is)
- Slide 2: product slide template with a text box for product_id and a table

### 3. Deploy to Streamlit Community Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect the repo and set `app.py` as the main file
4. Deploy

### Local development

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Usage

1. Select a PPT template from the dropdown
2. Upload an Excel file with product data
3. Upload a zip of product images
4. Click **Generate Catalog**
5. Download the resulting `.pptx`

### Excel columns (all case-insensitive)

| Column | Required | Notes |
|--------|----------|-------|
| `product_id` | Yes | Must match image folder names |
| `length` | No | Combined into "Dim. (cm)" |
| `width` | No | Combined into "Dim. (cm)" |
| `height` | No | Combined into "Dim. (cm)" |
| `Price` | No | Shown as "$ (USD)" |
| `Inner carton` | No | Combined into "Packing" |
| `Outer carton` | No | Combined into "Packing" |
| `Unit` | No | Combined into "Packing" |
| `Cbm` | No | Combined into "Packing" |
| `FOB Port` | No | Shown as "FOB" |

### Images zip structure

```
images.zip
├── H-10078#1/
│   ├── IMG_2731.png       # Main image (prefix "IMG")
│   ├── circle1_detail.png # Close-up 1
│   └── circle2_detail.png # Close-up 2
├── H-10078#2/
│   ├── IMG_0001.jpg
│   ├── circle1.png
│   ├── circle2.png
│   ├── circle3.png
│   └── circle4.png
```

- Folder names must match `product_id` values
- Main image: filename contains "IMG"
- Close-ups: filenames contain "circle1", "circle2", etc. (supports 2–4)
- `.png` files are preferred over `.jpg` when both exist
