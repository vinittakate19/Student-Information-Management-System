import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os

file_path = r"C:\Users\HP\OneDrive\Desktop\Py_1\doc.xlsx"

# Create Excel file with headers if not exists
if not os.path.exists(file_path):
    initial_df = pd.DataFrame(columns=['Name', 'Age', 'Email', 'Std', 'Div', 'Roll No', 'PRN'])
    initial_df.to_excel(file_path, index=False)

def open_data_entry():
    login_window.destroy()
    data_entry_form()

def login():
    username = "MECH_E3"
    password = "G12"
    if username_entry.get() == username and password_entry.get() == password:
        messagebox.showinfo(title="Login Success", message="You successfully logged in.")
        open_data_entry()
    else:
        messagebox.showerror(title="Error", message="Invalid login.")

def save_data():
    name = name_entry.get()
    age = age_entry.get()
    email = email_entry.get()
    std = std_entry.get()
    div = div_entry.get()
    rollno = rollno_entry.get()
    prn = prn_entry.get()

    if not name or not age or not email or not std or not div or not rollno or not prn:
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    try:
        age = int(age)
        rollno = int(rollno)
    except ValueError:
        messagebox.showerror("Error", "Age and Roll No must be numbers.")
        return

    new_data = {
        'Name': [name],
        'Age': [age],
        'Email': [email],
        'Std': [std],
        'Div': [div],
        'Roll No': [rollno],
        'PRN': [prn]
    }

    df = pd.DataFrame(new_data)

    try:
        existing_df = pd.read_excel(file_path)
        updated_df = pd.concat([existing_df, df], ignore_index=True)
        updated_df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", "Data saved successfully!")
        show_data()  # Refresh table
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        return

    for entry in [name_entry, age_entry, email_entry, std_entry, div_entry, rollno_entry, prn_entry]:
        entry.delete(0, tk.END)

def delete_all_data():
    """Deletes all rows after confirmation."""
    confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete ALL data?")
    if confirm:
        try:
            empty_df = pd.DataFrame(columns=['Name', 'Age', 'Email', 'Std', 'Div', 'Roll No', 'PRN'])
            empty_df.to_excel(file_path, index=False)
            messagebox.showinfo("Deleted", "All data deleted.")
            show_data()
        except Exception as e:
            messagebox.showerror("Error", f"Could not delete data: {e}")

def delete_specific_row():
    """Deletes a row based on Roll No input."""
    roll_to_delete = delete_entry.get()
    if not roll_to_delete:
        messagebox.showerror("Error", "Enter a Roll No to delete.")
        return

    try:
        roll_to_delete = int(roll_to_delete)
        df = pd.read_excel(file_path)
        df_before = len(df)
        df = df[df['Roll No'] != roll_to_delete]
        df_after = len(df)

        if df_before == df_after:
            messagebox.showinfo("Not Found", f"No data found with Roll No: {roll_to_delete}")
        else:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Deleted", f"Entry with Roll No {roll_to_delete} deleted.")
            delete_entry.delete(0, tk.END)
            show_data()

    except Exception as e:
        messagebox.showerror("Error", f"Problem deleting row: {e}")

def show_data():
    """Loads and displays Excel data in text box."""
    try:
        df = pd.read_excel(file_path)
        text_display.delete(1.0, tk.END)
        if df.empty:
            text_display.insert(tk.END, "No data available.")
        else:
            text_display.insert(tk.END, df.to_string(index=False))
    except Exception as e:
        text_display.insert(tk.END, f"Error reading file: {e}")

def data_entry_form():
    global name_entry, age_entry, email_entry, std_entry, div_entry, rollno_entry, prn_entry, delete_entry, text_display

    root = tk.Tk()
    root.title("Student Data Entry")

    labels = ["Name", "Age", "Email", "Class", "Division", "Roll No", "PRN"]
    entries = []

    for i, label in enumerate(labels):
        tk.Label(root, text=label + ":").grid(row=i, column=0, padx=5, pady=5, sticky="e")
        entry = tk.Entry(root)
        entry.grid(row=i, column=1, padx=5, pady=5)
        entries.append(entry)

    name_entry, age_entry, email_entry, std_entry, div_entry, rollno_entry, prn_entry = entries

    tk.Button(root, text="Save", width=15, command=save_data).grid(row=0, column=2, padx=10)
    tk.Button(root, text="Delete All Data", width=15, bg="red", fg="white", command=delete_all_data).grid(row=1, column=2, padx=10)

    tk.Label(root, text="Delete by Roll No:").grid(row=2, column=2, sticky="w")
    delete_entry = tk.Entry(root)
    delete_entry.grid(row=3, column=2, padx=5)
    tk.Button(root, text="Delete Entry", width=15, command=delete_specific_row).grid(row=4, column=2)

    # Display saved data
    tk.Label(root, text="Saved Entries:").grid(row=7, column=0, columnspan=3, pady=(20, 5))
    text_display = tk.Text(root, height=12, width=80)
    text_display.grid(row=8, column=0, columnspan=3, padx=10)

    show_data()

    root.mainloop()

# Login Form
login_window = tk.Tk()
login_window.title("Login Form")
login_window.geometry('340x250')
login_window.configure(bg='#333333')

frame = tk.Frame(login_window, bg='#333333')
frame.pack(pady=20)

tk.Label(frame, text="Login", bg='#333333', fg="#FF3399", font=("Arial", 24)).grid(row=0, column=0, columnspan=2, pady=10)

tk.Label(frame, text="Username", bg='#333333', fg="#FFFFFF").grid(row=1, column=0, pady=5)
username_entry = tk.Entry(frame)
username_entry.grid(row=1, column=1)

tk.Label(frame, text="Password", bg='#333333', fg="#FFFFFF").grid(row=2, column=0, pady=5)
password_entry = tk.Entry(frame, show="*")
password_entry.grid(row=2, column=1)

tk.Button(frame, text="Login", bg="#FF3399", fg="#FFFFFF", command=login).grid(row=3, column=0, columnspan=2, pady=10)

login_window.mainloop()
