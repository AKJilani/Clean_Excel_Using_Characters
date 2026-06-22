# Excel Unicode Cleaner Tool

A Python Tkinter desktop application that cleans Excel files by removing rows containing characters outside a defined allowed character set. It supports English letters (a–z), automatic lowercase handling, and custom Unicode characters like à, á, â, é, è, ò, ó, ö, ë, etc.

Users can select an Excel file, choose a sheet name, and enter allowed characters in a single space-separated input field. The tool automatically includes English a–z characters by default. It then scans each row in the selected sheet and deletes any row that contains characters not present in the allowed set. Processing progress is shown using a progress bar, and the cleaned file is saved as a new Excel file with a "_cleaned.xlsx" suffix. The application can also be packaged into an EXE using PyInstaller with the command: pyinstaller --onefile --windowed.
