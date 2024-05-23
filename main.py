import tkinter as tk
from tkinter import ttk
import os
import openpyxl
import add_user
import take_attendance

def reset_attendance():
    if os.path.exists('attendance.xlsx'):
        os.remove('attendance.xlsx')

def open_add_user():
    add_user.add_user()

def open_take_attendance():
    take_attendance.take_attendance()

def view_attendance():
    view_window = tk.Toplevel(root)
    view_window.title("View Attendance")
    view_window.geometry("600x400")

    tree = ttk.Treeview(view_window, columns=("Roll No", "Name", "Status", "Timestamp"), show="headings")
    tree.heading("Roll No", text="Roll No")
    tree.heading("Name", text="Name")
    tree.heading("Status", text="Status")
    tree.heading("Timestamp", text="Timestamp")
    tree.pack(fill=tk.BOTH, expand=True)

    if os.path.exists('attendance.xlsx'):
        try:
            workbook = openpyxl.load_workbook('attendance.xlsx')
            sheet = workbook["Attendance"]

            for row in sheet.iter_rows(min_row=2, values_only=True):
                tree.insert("", tk.END, values=row)
        except Exception as e:
            print("An error occurred while viewing attendance:", e)
            tk.messagebox.showerror("Error", "An error occurred while viewing attendance.")
    else:
        print("Attendance file does not exist.")

root = tk.Tk()
root.title("Camera-Based Attendance System")
root.geometry("400x250")

# Apply theme
style = ttk.Style(root)
root.tk.call("source", "azure.tcl")
style.theme_use("azure")

# Layout configuration
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

frame = ttk.Frame(root, padding="10 10 10 10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Reset attendance on startup
reset_attendance()

# Add buttons with styling
add_user_button = ttk.Button(frame, text="Add User", command=open_add_user)
add_user_button.grid(row=0, column=0, pady=10, sticky=tk.EW)

take_attendance_button = ttk.Button(frame, text="Take Attendance", command=open_take_attendance)
take_attendance_button.grid(row=1, column=0, pady=10, sticky=tk.EW)

view_attendance_button = ttk.Button(frame, text="View Attendance", command=view_attendance)
view_attendance_button.grid(row=2, column=0, pady=10, sticky=tk.EW)

reset_attendance_button = ttk.Button(frame, text="Reset Attendance", command=reset_attendance)
reset_attendance_button.grid(row=3, column=0, pady=10, sticky=tk.EW)

root.mainloop()
