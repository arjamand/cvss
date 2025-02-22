# Import necessary modules for creating the Flask app, handling templates, requests, and file operations
from flask import Flask, render_template, request, redirect, send_file
from analysis import readXML, createExcel, get_single_zip_file  # Custom functions for XML processing and Excel creation
import zipfile  # For handling ZIP file extraction
import os  # For file system operations

# Initialize the Flask app
app = Flask(__name__)

# Route for the homepage
@app.route("/")
def home():
    return render_template("home.html")  # Render the homepage template

# Route for handling file uploads and analysis
@app.route("/analysis", methods=['POST'])
def analysis():
    if request.method == "POST":  # Ensure the request method is POST
        # Retrieve uploaded files from the form
        system = request.form.get("system")
        tester = request.form.get("tester")
        date = request.form.get("date")
        xml = request.files['xml']  # XML file
        zip = request.files['zip']  # ZIP file

        # Create a unique directory for each analysis using an incremented ID
        id = len(os.listdir("./files/"))  # Get the number of existing directories
        os.mkdir(f"./files/{id}/")  # Create a new directory for the analysis
        xml.save(f"./files/{id}/file.xml")  # Save the XML file
        zip.save(f"./files/{id}/{zip.filename}.zip")  # Save the ZIP file
        
        # Extract the ZIP file contents into the newly created directory
        with zipfile.ZipFile(f"./files/{id}/{zip.filename}.zip", 'r') as zip_ref:
            zip_ref.extractall(f"./files/{id}/zip/")

        # Read and process the XML file
        with open(f"./files/{id}/file.xml", "r+") as f:
            data = readXML(f.read())  # Extract data from the XML file using the readXML function

        # Generate an Excel report based on the extracted data
        createExcel(id, data, system, tester, date)

        # Render the analysis results template with extracted data and ID
        return render_template("analysis.html", data=data, id=id, enumerate=enumerate)
    return redirect("/")  # Redirect to the homepage if the method is not POST

# Route for accessing previous analysis results by ID
@app.route("/analysis/<id>")
def prevanalysis(id):
    # Open the XML file and extract its data
    with open(f"./files/{id}/file.xml", "r+") as f:
        data = readXML(f.read())  # Read and process the XML file

    # Render the analysis results template for the given ID
    return render_template("analysis.html", data=data, id=id, enumerate=enumerate)

# Route for displaying the source code of a specific issue from the analysis
@app.route("/source/<id>/<ino>")
def source(id, ino):
    with open(f"./files/{id}/file.xml", "r+") as f:
        ino = int(ino) - 1  # Adjust index for zero-based indexing
        data = readXML(f.read())  # Read and process the XML file
        lineno = data[ino]['linenumber']  # Get the line number of the issue
        filepath = data[ino]['filepath']  # Get the file path of the issue

    filename = get_single_zip_file(f"./files/{id}/")
    # Open the relevant file from the extracted ZIP content
    with open(f"./files/{id}/zip/{filename}/" + filepath, "r+") as f:
        content = f.read()  # Read the file content

    # Render the source code template with the relevant details
    return render_template("source.html", lineno=lineno, content=content, filepath=filepath)

# Route for downloading the generated Excel report
@app.route("/download/<id>")
def download(id):
    # Send the generated Excel file to the client for download
    return send_file(f"./files/{id}/file.xlsx")

# Main entry point for running the Flask app
if __name__ == "__main__":
    app.run(debug=True)  # Run the app in debug mode for development