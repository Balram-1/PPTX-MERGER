# PPTFusion: Local Powerpoint Combiner

PPTFusion is a beautiful, locally-hosted web application that allows you to seamlessly combine multiple PowerPoint (`.pptx`) presentations into a single master deck. It features a premium "Luminous Void" design system with fluid animations, drag-and-drop support, and file rearrangement capabilities entirely within your browser.

## Features

- **Drag and Drop Interface:** Quickly upload PPTX files by dragging them into the dropzone.
- **Native File Dialog:** Click the dropzone to browse files from your computer.
- **Drag to Arrange:** Easily reorder the presentations by dragging the grip icon. The final document is merged in top-to-bottom order.
- **Remove Files:** Easily remove unwanted presentations using the remove button.
- **Client-Side Ordering:** Retains order visually and submits them sequentially to the backend.
- **Seamless Merging:** Generates a single `.pptx` file and automatically downloads it upon completion.
- **Premium Design:** Glassmorphic UI, dynamic ambient blur backgrounds, and smooth interactive elements.

## Prerequisites

- Python 3.8+
- `flask`
- `python-pptx`

## Installation & Setup

1. **Clone or Download the Repository**
2. **Install Dependencies**
   Run the following command in the project directory:
   ```bash
   pip install flask python-pptx
   ```
3. **Run the Application**
   ```bash
   python app.py
   ```
4. **Access the Web Interface**
   Open your browser and navigate to: `http://127.0.0.1:5000`

## Usage

1. Open `http://127.0.0.1:5000` in your web browser.
2. Click the upload zone to browse for `.pptx` files, or drag and drop them onto the dashboard.
3. Drag the loaded cards up and down using their grip handles to change the merge order.
4. Click **Merge Presentations**.
5. Wait for the loading animation. Once complete, your new merged presentation will download automatically.

## Project Structure

- `app.py`: The Flask server and Python merging logic (`python-pptx`).
- `templates/index.html`: The frontend UI (Vanilla HTML/CSS/JS).
- `uploads/`: Temporary folder for uploaded files during merge.
- `output/`: Temporary folder for generated merged presentations.

## License

This project is intended for personal and educational use.
