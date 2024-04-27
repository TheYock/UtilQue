import tkinter as tk

class TextStyler:
    def __init__(self, master):
        self.master = master
        master.title("Text Styler")

        self.text = tk.Text(master, wrap=tk.WORD, height=10, width=50, font=("Arial", 12))
        self.text.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        style_frame = tk.Frame(master)
        style_frame.pack(pady=5)

        self.bold_button = tk.Button(style_frame, text="Bold", command=self.bold_text)
        self.bold_button.pack(side=tk.LEFT, padx=5)

        self.italic_button = tk.Button(style_frame, text="Italic", command=self.italic_text)
        self.italic_button.pack(side=tk.LEFT, padx=5)

        self.underline_button = tk.Button(style_frame, text="Underline", command=self.underline_text)
        self.underline_button.pack(side=tk.LEFT, padx=5)

    def bold_text(self):
        self.modify_text("bold")

    def italic_text(self):
        self.modify_text("italic")

    def underline_text(self):
        self.modify_text("underline")

    def modify_text(self, option):
        sel = self.text.tag_ranges(tk.SEL)
        if option == "bold":
            self.text.tag_configure("bold", font=("Arial", 12, "bold"))
            if sel:
                self.text.tag_add("bold", sel[0], sel[1])
        elif option == "italic":
            self.text.tag_configure("italic", font=("Arial", 12, "italic"))
            if sel:
                self.text.tag_add("italic", sel[0], sel[1])
        elif option == "underline":
            self.text.tag_configure("underline", underline=True)
            if sel:
                self.text.tag_add("underline", sel[0], sel[1])