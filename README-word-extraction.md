# Word Document Content Extraction

This setup allows you to extract content from Word documents (.docx files) and automatically integrate it into your website.

## Setup

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Create a directory for your Word documents:**
   - The script will automatically create a `word-docs` directory if it doesn't exist
   - Or you can create it manually: `mkdir word-docs`

## Usage

1. **Place your Word documents in the `word-docs` directory:**
   - Name your files descriptively (e.g., `maps-location-cartographers.docx`, `about.docx`)
   - The script will try to match filenames to content keys automatically

2. **Run the extraction script:**
   ```bash
   python extract_from_word.py
   ```

   Or specify custom directories:
   ```bash
   python extract_from_word.py word-docs .
   ```

3. **The script will:**
   - Extract text content from Word documents
   - Identify headings and sections
   - Extract images (saved to `assets/` directory)
   - Update `assets/content.json` with the extracted content

## Content Key Mapping

The script automatically maps filenames to content keys:
- `maps`, `location`, `cartographer` → `maps`
- `tectonic`, `earthquake`, `volcano` → `tectonics`
- `weathering`, `erosion`, `mass` → `weathering`
- `fluvial`, `ocean`, `coastline` → `fluvial`
- `climate`, `biome` → `climate`
- `about` → `about`
- `home` → `home`
- `reference` → `references`

If the script can't determine the key, it will prompt you.

## Word Document Structure

For best results, structure your Word documents like this:

1. **Title** (Heading 1 or Title style)
2. **Introduction** (regular paragraph)
3. **Section Headings** (Heading 2 or bold text)
4. **Section Content** (regular paragraphs)

Example:
```
Maps, Location, and Cartographers

This page covers the cartographic history of Arizona...

Maps of Arizona

[Content about maps...]

Location and Geographic Information

[Content about location...]
```

## Extracted Images

Images from Word documents will be extracted and saved to the `assets/` directory. You'll need to manually update your HTML to reference these images, or modify the script to generate HTML with image references.

## Troubleshooting

- **No content extracted?** Make sure your Word document has text content and uses proper heading styles
- **Sections not recognized?** Use Word's built-in heading styles (Heading 1, Heading 2, etc.) or format headings as bold
- **Images not extracted?** Check that images are embedded in the Word document (not linked)

