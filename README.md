# **🔮 Universal Document Converter**

## **✨ Seamless, High-Fidelity Conversion powered by Pandoc**

Welcome to the **Universal Document Converter**, a robust web utility built with Python (Flask) and Pandoc designed to streamline your document workflow. Convert structured text formats—like Markdown and complex LaTeX projects—into high-quality Microsoft Word (.docx) files with perfect formatting, right from your browser.

The frontend features a modern, interactive **Glassmorphism** design using deep purple gradients for a visually engaging and professional experience.

### **🚀 Key Features**

* **Multi-Format Support:** Easily convert **Markdown**, **LaTeX**, **HTML**, and **Plain Text** to DOCX.  
* **📐 Full LaTeX Project Support:** Upload a .zip archive containing your multi-file LaTeX document, figures, and bibliography (.bib). (Requires main.tex in the root).  
* **➗ Equation Rendering:** LaTeX equations ($...$) are converted using MathML for native, editable rendering in Word documents.  
* **🎨 Vibing UI:** A dark-themed, responsive web interface utilizing **Glassmorphism** and vibrant purple gradients for a modern look and feel.  
* **💾 Custom Filenames:** Specify your desired output filename before downloading.  
* **⚡ Server-Side Reliability:** Built on Python/Flask, utilizing the powerful **Pandoc** command-line tool for guaranteed, consistent conversion quality.

## **🛠️ Tech Stack**

| Component | Technology | Role |
| :---- | :---- | :---- |
| **Backend** | Python 3.x, Flask | Handles routing, file management, and process control. |
| **Converter Core** | Pandoc | The engine responsible for format conversion (LaTeX \-\> DOCX, etc.). |
| **Frontend** | HTML5, Tailwind CSS | Modern, responsive UI with custom CSS for Glassmorphism. |
| **Styling/UX** | Font Awesome 6, Custom CSS | Icons, hover effects, and the dark/purple visual theme. |

