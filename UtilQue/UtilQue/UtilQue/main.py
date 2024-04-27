import tkinter as tk
from UtilQue import UtilQueApp

def main():
    root = tk.Tk()
    root.title("UtilQue")
    root.configure(bg="black")  # Set background color to black
    app = UtilQueApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()