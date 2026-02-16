
# PrintSpect-VBA-Optimization

A professional Word VBA solution designed for engineering and architectural specification workflows. This macro automates the generation of PDFs from Word documents while resolving common legacy issues related to 64-bit Office environments and duplex printing.

## 🚀 Key Features

* **Even-Page Logic:** Automatically detects if a document has an odd number of pages and inserts a blank page at the end. This ensures that when printing in duplex (double-sided), the next section starts on a fresh sheet.
* **64-Bit Compatibility:** Fully updated API declarations for `winspool.drv` and `kernel32` to run natively on modern 64-bit versions of Microsoft Word.
* **Dynamic Printer Selection:** Replaces the old "switch system default" method. It triggers the standard Print Dialog, allowing the user to choose Bluebeam, Adobe PDF, or any physical printer without affecting Windows settings.
* **Forced Color Mode:** Automatically sets the printer to Color mode for the duration of the print job to ensure project-specific highlighting (like instructional blue fields) is preserved.
* **Header Sanitization:** Automates the insertion of project-specific headers (`projname.doc`) and strips unwanted formatting baggage (like "boxes" or indents) that often plague legacy specification templates.

## 🛠 Installation

1.  Open your Microsoft Word document or Global Template (`Normal.dotm`).
2.  Press `Alt + F11` to open the Visual Basic for Applications (VBA) editor.
3.  Right-click on **Modules** in the Project Explorer and select **Insert > Module**.
4.  Copy the code from `PrintSpectPDF07.bas` in this repository and paste it into the new module.
5.  Ensure the file `projname.doc` exists in the same directory as your specification file.

## 📖 How To Use

1.  Click the **Macros** button on the **Developer** tab (or press `Alt + F8`).
2.  Select `MAIN` (or `PrintSpectPDF07.MAIN`) and click **Run**.
3.  When the Print Dialog appears, select your PDF printer.
4.  Click **Print**. The macro will handle the page counts, color settings, and formatting overrides automatically.

## 📋 Technical Requirements

* **Environment:** Microsoft Word (32-bit or 64-bit).
* **OS:** Windows 10 or 11.
* **Permissions:** Ability to run VBA macros and access to the local folder path for header insertion.

## ⚖️ License
This project is open-source and available for professional use and modification.
