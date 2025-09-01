# Inventory Web Application

This is a web-based Inventory Management Application built with Python (likely Flask) that allows users to manage inventory items through a web interface. The application serves dynamic HTML pages using templates and handles static assets such as images and stylesheets.

## Features

- Web interface for managing inventory
- Dynamic HTML rendering using templates
- Static assets support (images, icons, styles)
- File upload support (uploads folder present)
- Simple and lightweight design

## Project Structure

- `app.py` - Main application script to run the web server
- `requirements.txt` - Python dependencies for the project
- `templates/` - HTML templates for rendering web pages
- `static/` - Static files such as images and icons
- `uploads/` - Directory for uploaded files
- `downloads/` - Directory for downloadable files

## Setup and Installation

1. Create a virtual environment (recommended):

   ```bash
   python -m venv venv
   ```

2. Activate the virtual environment:

   - On Windows:

     ```bash
     venv\Scripts\activate
     ```

   - On macOS/Linux:

     ```bash
     source venv/bin/activate
     ```

3. Install the required dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

Run the application using:

```bash
python app.py
```

By default, the application will start a local development server (usually on `http://127.0.0.1:5000/`).

Open your web browser and navigate to the URL to access the inventory web app.

## Usage

- Use the web interface to view and manage inventory items.
- Upload files if supported by the application.
- Download files from the downloads section if available.
