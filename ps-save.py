import win32com.client # Import the win32com.client module

# Create an instance of the Photoshop Application object
ps_app = win32com.client.Dispatch("Photoshop.Application")

# Get the active document
doc = ps_app.ActiveDocument

# Save the document
try:
    doc.Save()
    print("Document saved successfully")
except Exception as e:
    print(f"Unable to save the document, {e}")
