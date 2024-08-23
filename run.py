import tkinter as tk
from tkinter import messagebox
import pyodbc
import threading
import time


def run_sp():
    try:
        # Show that the execution is running
        start_animation()

        conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                              'SERVER=tcp:10.1.1.103\SQLEXPRESS;'
                              'DATABASE=AccountingSavedQueries;'
                              'UID=run_store_procedure;'
                              'PWD=3TXw-#3$24;'
                              'TrustServerCertificate=yes;'
                              'Connection Timeout=30;')
        cursor = conn.cursor()
        query = (f"DECLARE @return_value int "
                 f"EXEC @return_value = [dbo].[usp_RefreshOverheadAbsorption] "
                 f"SELECT @return_value AS 'Return Value'")

        cursor.execute(query)
        return_value = cursor.fetchone()[0]

        if return_value == 0:
            messagebox.showinfo(
                "Success", "Stored Procedure executed successfully!")
        else:
            messagebox.showwarning(
                "Warning", f"Stored Procedure executed with warnings! Return Value: {return_value}")

    except Exception as e:
        messagebox.showerror(
            "Error", f"Failed to execute stored procedure: {e}")

    finally:
        # Stop the animation after execution
        stop_animation()
        cursor.close()
        conn.close()


def start_animation():
    global stop_animation_flag
    stop_animation_flag = False
    animate()


def stop_animation():
    global stop_animation_flag
    stop_animation_flag = True


def animate():
    global stop_animation_flag
    if not stop_animation_flag:
        current_text = loading_label.cget("text")
        new_text = "Loading" if current_text == "Loading..." else current_text + "."
        loading_label.config(text=new_text)
        root.after(500, animate)


def run_sp_thread():
    threading.Thread(target=run_sp).start()


# Create the main window
root = tk.Tk()
root.title("Run Stored Procedure")

# Set the size of the window (width x height)
root.geometry("300x150")

# Create a label to show loading animation
loading_label = tk.Label(root, text="")
loading_label.pack(pady=10)

# Create a button to run the stored procedure with adjusted size
run_button = tk.Button(root, text="Run Stored Procedure",
                       command=run_sp_thread, width=20, height=2)
run_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
