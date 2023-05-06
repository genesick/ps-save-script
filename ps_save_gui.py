import tkinter as tk
import win32com.client
import pywintypes
import time

class PS_SAVE:
    def __init__(self):
        self.name:str = ""
        self.ps_app = None
        self.create_window()
        self.create_buttons()
        self.create_status()
        self.auto_save()
        self.window.mainloop() # Run the window

    def get_document_name(self):
        try:
            if self.ps_app is not None:
                doc = self.ps_app.ActiveDocument
                self.change_message(f"Opened {doc.Name}", "black")
                return
        except pywintypes.com_error as e:
            print(e)

        self.window.after(1000, self.get_document_name)


    def save_document(self):
        try:
            if self.ps_app is not None:
                # Get the active document
                doc = self.ps_app.ActiveDocument
                if doc is None:
                    print("No document selected")
                    self.change_message("No document selected", "black")
                    return
                # Save the document
                doc.Save()
                self.name = doc.Name
                self.change_message(f"{self.name} saved successfully!", "black")
        except pywintypes.com_error as e:
            message = "Error: " + str(e)
            self.change_message(message, "red")


    def auto_save(self):
        self.save_document() # save document
        self.change_message(f"Auto-saving {self.name}...", "black") # update message
        self.window.after(1000, self.change_message, f"{self.name} saved successfully!", "black") # update message again after a second
        self.window.after(2000, self.auto_save) # auto-save every second



    def change_message(self, message: str, color: str):
        self.message_label.config(text=message)
        self.message_label.config(fg=color)
        print(message)

    def open_photoshop(self):
        try:
            self.ps_app = win32com.client.Dispatch("Photoshop.Application")
            self.change_message("Opened Photoshop succesfully!", "black")
            self.get_document_name()
        except pywintypes.com_error as e:
            message = "Error: " + str(e)
            self.change_message(message, "red")

    def create_window(self):
        # Create a window
        self.window = tk.Tk()
        self.window.title("Save in Photoshop")
        self.window.geometry("200x110")  # Set the size of the window
        self.window.resizable(False, False)  # Set the window to not be resizable
        self.window.configure(bg="#81849f")

    def create_buttons(self):
        self.open_photoshop_button = tk.Button(self.window, text="OPEN PHOTOSHOP", 
                                    command=self.open_photoshop, 
                                    padx=30, 
                                    pady=20)
        self.open_photoshop_button.configure(bg="#fff")
        self.open_photoshop_button.pack(pady=10)
        
        
    def create_status(self):
        # Create a label to display messages
        self.message_label = tk.Label(self.window, text="")
        self.message_label.pack()


if __name__ == "__main__":
    ps_save = PS_SAVE()
