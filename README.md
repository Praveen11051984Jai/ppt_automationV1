# PowerPoint Automation Tool

## 📋 Overview

The **PowerPoint Automation Tool** is a powerful Python-based application designed to streamline and automate the creation of customized PowerPoint presentations. With an intuitive GUI interface, this tool allows users to combine content from multiple sources, search documents for relevant information, and generate professional presentations in minutes.

## ✨ Key Features

### 1. **Customizable Title & Ending Slides**
- Add company and customer logos
- Personalize with author name and date
- Choose custom background colors
- Multiple ending slide templates (Thank You, Questions?, Contact Us, Custom Message)

### 2. **Smart Content Selection**
- Import and scan multiple PowerPoint presentations
- Preview all available slides with titles
- Select specific slides to include in the new presentation
- Preserve original formatting, images, tables, and charts

### 3. **Document Intelligence**
- **Text Search**: Search keywords in PDF, Word, and Excel files
- **Image Search**: Find and insert relevant images from folders
- **Chart Generation**: Extract numerical data from Excel and create charts
- Context-aware content extraction

### 4. **Safe Slide Copying**
- Shape-by-shape duplication to prevent XML corruption
- Maintains all formatting, colors, and styles
- Supports text, images, tables, and complex layouts
- No data loss or corruption issues

## 🚀 Installation

### Prerequisites
- Python 3.7 or higher
- pip package manager

### Required Libraries

```bash
pip install python-pptx
pip install PyPDF2
pip install python-docx
pip install openpyxl
pip install pillow
pip install pdf2image
```

### Quick Install (All at once)

```bash
pip install python-pptx PyPDF2 python-docx openpyxl pillow pdf2image
```

## 📖 How to Use

### Step 1: Launch the Application

```bash
python ppt_automation_tool.py
```

The GUI window will open with a tabbed interface.

### Step 2: Configure Title & End Slides

1. Navigate to **Step 1** tab
2. Enter your presentation title
3. Enter author name
4. Upload company logo (optional)
5. Upload customer logo (optional)
6. Choose background color
7. Select ending slide type

### Step 3: Select Content from Existing Presentations

1. Navigate to **Step 2** tab
2. Click **"Add Presentation(s)"** to select source PowerPoint files
3. Click **"Scan Slides"** to analyze all presentations
4. Review the slide list with titles
5. Select the slides you want to include (use Ctrl+Click for multiple selections)

### Step 4: Add New Content from Documents

1. Navigate to **Step 3** tab
2. Click **"Browse"** to select your support documents folder
3. Choose content type:
   - **Text**: Search for keywords in documents
   - **Images**: Find images matching topics
   - **Chart Data**: Search Excel files for numerical data
4. Enter a topic/keyword
5. Click **"Search"** to find relevant content
6. Review search results
7. Click **"Add to Presentation"** to include the topic as a new slide
8. Repeat for multiple topics
9. Click **"Refresh Preview"** to see all selected content

### Step 5: Generate Your Presentation

1. Navigate to **Step 4** tab
2. Review or modify the output filename
3. Choose output directory (default: Desktop)
4. Enable/disable automatic slide numbering
5. Click **"Generate Presentation"**
6. View the generation summary
7. Your presentation is ready!

## 🏗️ Technical Architecture

### Core Components

- **GUI Framework**: Tkinter with ttk widgets
- **PowerPoint Processing**: python-pptx library
- **Document Readers**: PyPDF2, python-docx, openpyxl
- **Image Processing**: Pillow (PIL)

### Architecture Pattern

- Object-Oriented Design
- Model-View-Controller (MVC) pattern
- Event-driven GUI
- Modular component structure

### Slide Copying Method

The tool uses a **safe shape-by-shape duplication** method:
1. Creates a temporary intermediate presentation
2. Copies each shape individually using native APIs
3. Preserves formatting, images, tables, and text
4. Avoids XML namespace corruption
5. Comprehensive error handling and logging

## 🎯 Use Cases

### Business Presentations
- Create customer-specific presentations from template slides
- Combine content from multiple product presentations
- Generate proposal presentations with dynamic content

### Training & Education
- Compile training materials from various sources
- Create customized course presentations
- Combine slides from different modules

### Sales & Marketing
- Generate personalized sales decks
- Create pitch presentations with customer logos
- Assemble case studies and testimonials

### Reporting
- Create monthly/quarterly reports
- Combine data visualizations from Excel
- Generate executive summaries

## 💡 Tips & Best Practices

### Organization
- ✓ Keep source presentations in a dedicated folder
- ✓ Use descriptive filenames for easy identification
- ✓ Organize support documents by category

### Content Preparation
- ✓ Use high-resolution logos (PNG format recommended)
- ✓ Name images with relevant keywords for easy search
- ✓ Keep Excel files clean with clear headers
- ✓ Ensure PDF documents have selectable text

### Using the Tool
- ✓ Test with a small presentation first
- ✓ Preview content before generating final presentation
- ✓ Run from command line to see debug information
- ✓ Check console output for detailed logging

### Performance
- ✓ Close other applications when processing large presentations
- ✓ Use a limited number of source presentations (5-10 recommended)
- ✓ Optimize images before adding to presentations

## 🐛 Troubleshooting

### Issue: "No module named 'pptx'"
**Solution**: Install python-pptx: `pip install python-pptx`

### Issue: Generated presentation needs repair
**Solution**: The latest version uses safe copying methods. Ensure you're using the updated code.

### Issue: Images not found
**Solution**: Check that the support documents folder path is correct and images have matching keywords in filenames.

### Issue: Text not copying correctly
**Solution**: Ensure source presentations are not password-protected or corrupted.

### Issue: Chart data not found
**Solution**: Make sure Excel files contain numerical data and sheet names match search keywords.

## 🔍 Console Output

Run the application from command line to see detailed logging:

```
=== Starting Slide Copy Process ===
Copying slide 1 from presentation.pptx
  Copying slide with 5 shapes
    Copying shape 1/5: Type=17
      ✓ Text copied: 2 paragraph(s)
    Copying shape 2/5: Type=13
      ✓ Picture copied
  ✓ Safely copied 5 shapes
✓ Slide 1 copied successfully
```

## 📝 Supported File Formats

### Input
- PowerPoint: `.pptx`, `.ppt`
- Documents: `.pdf`, `.docx`
- Spreadsheets: `.xlsx`, `.xls`
- Images: `.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`

### Output
- PowerPoint: `.pptx`

## 🔐 Limitations

- Charts from source presentations are copied as placeholders (require manual recreation)
- Password-protected files cannot be read
- Very large presentations (100+ slides) may take longer to process
- Embedded videos are not copied
- Animations and transitions are not preserved

## 🛠️ System Requirements

- **Operating System**: Windows, macOS, or Linux
- **Python**: 3.7 or higher
- **RAM**: 4GB minimum (8GB recommended)
- **Disk Space**: 100MB for application and libraries
- **Display**: 1200x800 minimum resolution

## 📦 Project Structure

```
ppt_automation_tool/
│
├── ppt_automation_tool.py      # Main application file
├── README.md                    # This file
├── requirements.txt             # Python dependencies
└── examples/                    # Example presentations (optional)
```

## 🤝 Contributing

Contributions are welcome! Areas for improvement:
- Add support for more chart types
- Implement animation preservation
- Add template library
- Improve search algorithms
- Add batch processing mode

## 📄 License

This project is provided as-is for educational and commercial use.

## 📞 Support

For issues and questions:
1. Check the console output for error messages
2. Review the Troubleshooting section
3. Ensure all dependencies are installed correctly

## 🎉 Version History

### Version 2.0 (Current)
- ✓ Added background color selection
- ✓ Implemented safe slide copying (no corruption)
- ✓ Added image search functionality
- ✓ Added chart generation from Excel data
- ✓ Enhanced text formatting preservation
- ✓ Improved error handling and logging

### Version 1.0
- Initial release with basic functionality

## 🌟 Acknowledgments

Built with:
- python-pptx by Steve Canny
- Tkinter (Python Standard Library)
- PyPDF2, python-docx, openpyxl communities

---

**PowerPoint Automation Tool** - Streamline Your Presentation Workflow

Created with ❤️ using Python
# ppt_automationV1
PPT_Automation_tool
