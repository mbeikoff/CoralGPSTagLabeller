# Photo Matcher Tool

A desktop GUI application that helps match photos taken from multiple cameras to event timestamps recorded in an Excel (.xlsx) file. It extracts datetime information from photo EXIF metadata and finds the closest-matching photo within a user-defined time threshold.

Perfect for synchronizing photos from different devices (e.g., action cameras labeled PDP1/PDP2/PDP3) with a log of GPS-tagged events.

## Features

- Load an Excel file containing at least `Time` and `Camera` columns
- Select separate folders for three camera groups (Green → PDP1, White → PDP2, Third → PDP3)
- Automatically extract EXIF creation times from JPEG photos
- Thumbnail grid view with checkboxes to manually include/exclude photos
- Configurable matching threshold in seconds (0 = exact match only)
- Progress bars for photo loading and matching
- Preview of all successful matches in a scrollable table
- Export matched rows (with added Subfolder and Filename columns) to a new Excel file
- Robust time parsing that handles various Excel date/time formats, including Australian day-first formats

## Requirements

- Python 3.6+
- Required packages:
  ```bash
  pip install pandas pillow openpyxl
  ```

`tkinter` is included with standard Python installations.

## Usage

1. Save the script as `photo_matcher.py` (or any name you prefer).
2. Run it:
   ```bash
   python photo_matcher.py
   ```
3. **Initial screen**:
   - Choose your input `.xlsx` file (must have `Time` and `Camera` columns with values PDP1, PDP2, or PDP3).
   - Select a folder for each camera color (Green, White, Third).
   - Set the match threshold in seconds (e.g., `30` for ±30 seconds).
   - Optionally edit the output base filename.
   - Click **Load Photos for Selection**.

4. **Photo selection screen**:
   - Browse between cameras using Previous/Next buttons.
   - Check the boxes next to photos you want to include in matching (by default none are selected).
   - Use **Select All** / **Deselect All** for convenience.
   - Click **Proceed to Matching**.

5. **Preview screen**:
   - Review the matched results in the table.
   - Click **Export Matches** to save a new Excel file with only the matched rows and added photo information.

## Notes

- Only `.jpg` and `.jpeg` files are processed.
- Photos without valid EXIF datetime are skipped and counted in the summary.
- Excel times are parsed robustly (supports many formats including partial times and Excel serial dates).
- The tool runs entirely locally — no internet connection required.

## Screenshots

*(Add screenshots here once you have them — they display nicely on GitHub)*

Example workflow:
- Initial setup screen
- Photo selection grid for one camera
- Matching preview table

## Contributing

Feel free to open issues or pull requests! Suggestions for improvements (e.g., better thumbnail highlighting, drag-and-drop support, or additional file formats) are welcome.

## License

Apache 2.0 License — see the [LICENSE](LICENSE) file for details.

---

