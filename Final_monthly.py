from flask import Flask, render_template_string, request
import openpyxl

app = Flask(__name__)

html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Excel Processor</title>
  <style>
    body {
      font-family: 'Arial', sans-serif;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
      background-color: #ecf0f1;
    }
    .card {
      background-color: white;
      padding: 20px;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
      border-radius: 5px;
      text-align: center;
    }
    .file-input {
      margin-top: 10px;
    }
  </style>
</head>
<body>
  <div class="card">
    <h3>Choose Your Excel File</h3>
    <form method="POST" enctype="multipart/form-data">
      <input type="file" name="file" class="file-input" accept=".xls,.xlsx" required>
      <br><br>
      <button type="submit">Upload and Process</button>
    </form>
  </div>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        file = request.files["file"]
        if file:
            filepath = "uploaded_file.xlsx"
            file.save(filepath)
            
            # Process the uploaded Excel file using your existing logic
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook['Sheet1']
            
            # Example of your existing logic
            Final_machine_reading_tokiem = sheet['F7'].value
            Initial_machine_reading_tokiem = sheet['F8'].value
            print(f"Final: {Final_machine_reading_tokiem}, Initial: {Initial_machine_reading_tokiem}")
            
            # Save modifications (optional)
            workbook.save("Modified_File.xlsx")
            
            return "File processed successfully! Check the server for the modified file."
    return render_template_string(html_template)

if __name__ == "__main__":
    app.run(debug=True)

