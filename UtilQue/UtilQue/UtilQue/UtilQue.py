import os
import tkinter as tk
from tkinter import filedialog, messagebox
import re
from datetime import datetime

class UtilQueApp:
    def __init__(self, master):
        self.master = master
        master.title("UtilQue")
        self.ticket_widgets = []
        self.folder_path = None  # Initialize folder_path attribute
        self.created_tickets = set()  # Initialize set to store ticket numbers of created widgets

        # Create a frame to contain the status text and scrollbar
        self.status_frame = tk.Frame(height=50, width=200)
        self.status_frame.pack(fill=tk.BOTH, expand=False)

        # Create a frame to contain the list of ticket widgets and scrollbar
        self.widget_frame = tk.Frame(master)
        self.widget_frame.pack(fill=tk.BOTH, expand=True)

        # Create a canvas to contain the list of ticket widgets
        self.canvas = tk.Canvas(self.widget_frame, bg="white")
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add scrollbar to the canvas
        self.scrollbar_widgets = tk.Scrollbar(self.widget_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar_widgets.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.config(yscrollcommand=self.scrollbar_widgets.set)

        # Create a frame to contain the list of ticket widgets inside the canvas
        self.widget_list_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.widget_list_frame, anchor=tk.NW)

        # Update the scroll region of the canvas when the size of the widget list changes
        self.widget_list_frame.bind("<Configure>", lambda event: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Add the widget_list_frame to the canvas with proper expansion behavior
        self.canvas.create_window((0, 0), window=self.widget_list_frame, anchor=tk.NW, tags="frame")

        # Configure the canvas to expand the widget_list_frame
        self.canvas.bind("<Configure>", lambda event, canvas=self.canvas: canvas.itemconfigure("frame", width=event.width))

        # Bind the MouseWheel event to the canvas for scrolling
        self.widget_frame.bind_all("<MouseWheel>", self.on_mousewheel)

        # Create a text box to display status messages
        self.status_text = tk.Text(self.status_frame, height=5, width=50)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Create the Select Folder button
        self.select_folder_button = tk.Button(master, text="Select Folder", command=self.select_folder)
        self.select_folder_button.pack()

        # Create the Refresh button but keep it hidden initially
        self.refresh_button = tk.Button(master, text="Refresh", command=self.refresh_clicked)
        self.refresh_button.pack_forget()
        
        # Create the Exit button
        self.exit_button = tk.Button(master, text="Exit", command=master.quit)
        self.exit_button.pack(side=tk.TOP, anchor="ne", padx=10, pady=10)
        
        # Display initial status message
        self.show_status("Awaiting folder directory...")

    def select_folder(self):
        self.folder_path = filedialog.askdirectory(title="Select Folder")  # Prompt to select folder
        if self.folder_path:
            if self.create_required_folders(self.folder_path):
                unprocessed_folder = os.path.join(self.folder_path, "Unprocessed")  # Get the path of the Unprocessed folder
                self.process_eml_files(unprocessed_folder)  # Process EML files
                markouts_folder = os.path.join(self.folder_path, "Markouts")  # Get the path of the Markouts folder
                self.process_markout_files(markouts_folder)  # Process markout files
                self.select_folder_button.pack_forget()  # Hide the Select Folder button
                self.refresh_button.pack()  # Show the Refresh button
            else:
                messagebox.showinfo("Folders Created", "All required folders were created.")
        else:
            messagebox.showwarning("Folder not selected", "Please select a folder.")

    def refresh_clicked(self):
        if self.folder_path:
            unprocessed_folder = os.path.join(self.folder_path, "Unprocessed")  # Get the path of the Unprocessed folder
            self.process_eml_files(unprocessed_folder)  # Process EML files
            markouts_folder = os.path.join(self.folder_path, "Markouts")  # Get the path of the Markouts folder
            self.process_markout_files(markouts_folder)  # Process markout files
        else:
            messagebox.showwarning("Folder not selected", "Please select a folder.")

    def create_required_folders(self, folder_path):
        # List of required folders
        required_folders = ["Unprocessed", "Processed", "Markouts", "Completed"]

        # Create the selected folder if it doesn't exist
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            print(f"Created folder: {folder_path}")

        # Check if all required folders exist within the selected folder
        all_folders_exist = all(os.path.exists(os.path.join(folder_path, folder)) for folder in required_folders)
        if not all_folders_exist:
            print("Some or all required folders do not exist. Creating...")
            for folder in required_folders:
                folder_path = os.path.join(self.folder_path, folder)  # Use selected folder path as base
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path)
                    print(f"Created folder: {folder_path}")
                else:
                    print(f"Folder already exists: {folder_path}")
            return False  # Some or all required folders were not found and were created
        else:
            return True  # All required folders were found

    def process_eml_files(self, unprocessed_folder):
        # Process .eml files in the "Unprocessed" folder
        print("Folder path:", unprocessed_folder)  # Debugging print statement
        self.show_status("Processing .eml files...")
        eml_found = False  # Flag to track if EML files were found
        for filename in os.listdir(unprocessed_folder):
            if filename.endswith(".eml"):
                eml_found = True
                eml_file_path = os.path.join(unprocessed_folder, filename)
                with open(eml_file_path, "r") as eml_file:
                    eml_content = eml_file.read()
                # Extract content between "[EXTERNAL EMAIL]" and "End Request"
                start_index = eml_content.find("[EXTERNAL EMAIL]")
                end_index = eml_content.find("End Request", start_index)
                if start_index != -1 and end_index != -1:
                    end_index += len("End Request")  # Include "End Request" in the processed content
                    processed_content = eml_content[start_index:end_index]
                    # Save processed content as .txt file in "Markouts" folder
                    markouts_folder = os.path.join(os.path.dirname(unprocessed_folder), "Markouts")
                    txt_file_path = os.path.join(markouts_folder, filename.replace(".eml", ".txt"))
                    with open(txt_file_path, "w") as txt_file:
                        txt_file.write(processed_content)
                    # Move processed .eml file to "Processed" folder
                    processed_folder = os.path.join(os.path.dirname(unprocessed_folder), "Processed")
                    processed_file_path = os.path.join(processed_folder, filename)
                    os.rename(eml_file_path, processed_file_path)
        if not eml_found:
            self.show_status("No EML Files Found")
        else:
            self.show_status("Processed .eml files.")

    def process_markout_files(self, folder_path):
        # Process .txt files in the "Markouts" folder
        print("Folder path:", folder_path)  # Debugging print statement
        self.show_status("Processing .txt files in Markouts folder...")
        markouts_found = False  # Flag to track if markouts files were found
        for filename in os.listdir(folder_path):
            if filename.endswith(".txt"):
                markouts_found = True
                txt_file_path = os.path.join(folder_path, filename)
                with open(txt_file_path, "r") as txt_file:
                    txt_content = txt_file.read()
                # Extract relevant information from the text content
                ticket_info = {
                    "Ticket Number": self.extract_ticket_number(txt_content),
                    "Routine or Emergency": self.extract_ticket_type(txt_content),
                    "Street Address": self.extract_street_address(txt_content),
                    "Type of Work": self.extract_type_of_work(txt_content),
                    "Extent of Work": self.extract_extent_of_work(txt_content),
                    "Start Time": self.extract_start_time(txt_content)
                }
                # Check if widget for this ticket already exists
                ticket_number = ticket_info["Ticket Number"]
                if ticket_number not in self.created_tickets:
                    # Create ticket widget
                    ticket_widget = TicketWidget(self.widget_list_frame, self, ticket_info, filename)
                    ticket_widget.pack(fill="x", padx=10, pady=5)
                    self.ticket_widgets.append(ticket_widget)
                    self.created_tickets.add(ticket_number)  # Add ticket number to the set
        if not markouts_found:
            messagebox.showinfo("No Markouts Found", "No markouts files were found in the Markouts folder.")
        self.show_status("Processed .txt files.")

    def extract_ticket_number(self, text_content):
        # Extract ticket number from text content
        pattern = r'Request No.: (\d+)'
        match = re.search(pattern, text_content)
        if match:
            return match.group(1)  # Extracting the ticket number
        else:
            return None  # Return None if no ticket number found

    def extract_ticket_type(self, text_content):
        # Extract routine or emergency status from text content
        if "*** E M E R G E N C Y" in text_content:
            return "Emergency"
        else:
            return "Routine"

    def extract_street_address(self, text_content):
        # Extract street address from text content
        pattern = r'Street:\s*(.*?)\n'
        match = re.search(pattern, text_content, re.DOTALL)
        if match:
            return match.group(1).strip()
        else:
            return None

    def extract_type_of_work(self, text_content):
        # Extract type of work from text content
        pattern = r'Type of Work:\s*(.*?)\n'
        match = re.search(pattern, text_content, re.DOTALL)
        if match:
            return match.group(1).strip()
        else:
            return None

    def extract_extent_of_work(self, text_content):
        # Extract extent of work from text content, including all text between "Extent of Work" and "Remarks"
        pattern = r'Extent of Work:(.*?)Remarks'
        match = re.search(pattern, text_content, re.DOTALL)
        if match:
            return match.group(1).strip()
        else:
            return None

    def extract_start_time(self, text_content):
        # Extract start time from text content
        pattern = r'Start Date/Time:\s*(.*?)\n'
        match = re.search(pattern, text_content, re.DOTALL)
        if match:
            return match.group(1).strip()
        else:
            return None

    def show_status(self, message):
        # Display status message in the text box
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)  # Scroll to the end of the text box
        
    def on_mousewheel(self, event):
        # Perform vertical scrolling when the mouse wheel is scrolled
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def deselect_all(self):
        for widget in self.ticket_widgets:
            widget.deselect()

    def deselect_others(self, selected_widget):
        for widget in self.ticket_widgets:
            if widget != selected_widget:
                widget.deselect()

class TicketWidget(tk.Frame):
    def __init__(self, master, util_que_app, ticket_info, filename):
        super().__init__(master)
        self.util_que_app = util_que_app  # Pass the instance of UtilQueApp
        self.ticket_info = ticket_info
        self.filename = filename  # Store the filename
        self.selected = False
        self.create_widgets()

    def create_widgets(self):
        # Check if ticket type is "Emergency" and set background color to red
        bg_color = "#ff9999" if self.ticket_info.get("Routine or Emergency") == "Emergency" else "white"
        self.configure(bg=bg_color, highlightbackground="white", highlightcolor="white", highlightthickness=1)

        # Create labels for ticket information
        for i, (key, value) in enumerate(self.ticket_info.items()):
            label = tk.Label(self, text=f"{key}: {value}", anchor="w", padx=5, pady=3)
            label.grid(row=i//3, column=i%3, sticky="w")

        # Create the Open File button
        open_file_button = tk.Button(self, text="Show Ticket", command=self.open_txt_file)
        open_file_button.grid(row=3, column=0, sticky="we")

        # Create a blank column
        blank_label = tk.Label(self, text="", padx=5, pady=3)
        blank_label.grid(row=3, column=1)

        # Create the Complete Ticket button
        complete_button = tk.Button(self, text="Complete", command=self.complete_ticket)
        complete_button.grid(row=3, column=2, sticky="we")

        self.bind("<Enter>", self.hover)
        self.bind("<Leave>", self.unhover)

    def hover(self, event=None):
        if not self.selected:
            if self.ticket_info.get("Routine or Emergency") == "Emergency":
                bg_color = "red"  # Red for emergency
            else:
                bg_color = "lightblue"  # Light blue for routine
            self.configure(bg=bg_color)
            for label in self.winfo_children():
                label.configure(background=bg_color, foreground="black")

    def unhover(self, event=None):
        if not self.selected:
            if self.ticket_info.get("Routine or Emergency") == "Emergency":
                bg_color = "#ff9999"  # Keep red for emergency
            else:
                bg_color = "white"  # Reset to white for routine
            self.configure(bg=bg_color)
            for label in self.winfo_children():
                label.configure(background=bg_color, foreground="black")


    def open_txt_file(self):
        # Open the corresponding .txt file using the default text editor
        txt_filepath = os.path.join(self.util_que_app.folder_path, "Markouts", self.filename)
        if os.path.exists(txt_filepath):
            os.startfile(txt_filepath)
        else:
            messagebox.showerror("File Not Found", f"File '{self.filename}' not found.")

    def complete_ticket(self):
        # Prompt the user to confirm completion
        confirm = messagebox.askyesno("Complete Ticket", "Are you sure you want to mark this ticket as complete?")
        if confirm:
            # Open and modify the .txt file
            txt_filepath = os.path.join(self.util_que_app.folder_path, "Markouts", self.filename)
            if os.path.exists(txt_filepath):
                with open(txt_filepath, "a") as txt_file:
                    txt_file.write(f"\nCompleted at {datetime.now()}")

                # Move the .txt file to the "Completed" folder
                dest_folder = os.path.join(self.util_que_app.folder_path, "Completed")
                dest_filepath = os.path.join(dest_folder, self.filename)
                try:
                    os.rename(txt_filepath, dest_filepath)
                    # Remove the widget from the ticket list
                    self.destroy()
                    self.util_que_app.ticket_widgets.remove(self)
                    # Update the canvas scroll region
                    self.util_que_app.canvas.update_idletasks()
                    self.util_que_app.canvas.configure(scrollregion=self.util_que_app.canvas.bbox("all"))
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to complete ticket: {e}")
            else:
                messagebox.showerror("File Not Found", f"File '{self.filename}' not found.")


def main():
    root = tk.Tk()
    root.title("UtilQue")
    root.state('zoomed')

    # Configure background color of the root window and foreground color of all text
    root.configure(bg="#333333")

    # Configure specific widget colors
    root.option_add("*Button.Background", "#444444")  # Dark gray for buttons
    root.option_add("*Button.Foreground", "white")     # White text on buttons
    root.option_add("*Label.Background", "#333333")    # Dark gray for labels
    root.option_add("*Label.Foreground", "white")      # White text on labels

    # Create and pack widgets
    app = UtilQueApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()