import tkinter as tk
from tkinter import ttk
import pyvisa

def convert_temperature(temp_format):
    # Convert the temperature format (multiple of 10) to actual temperature
    temp = float(temp_format) / 10.0
    return temp

def convert_temperature_format(temp):
    # Convert the actual temperature to temperature format (multiple of 10)
    temp_format = str(int(float(temp) * 10))
    return temp_format

def send_query_command(command):
    # Send the command to the device and read the response
    retries = 1
    while retries > 0:
        try:
            response = device.query(command + '\r')
            return response
        except pyvisa.VisaIOError as e:
            retries -= 1
            print(f"An error occurred while sending the command, please try again")
    print("Failed to communicate with the device. Please check the connection and try again.")
    return None

def read_actual_temperature():
    response = send_query_command('R 100,1')
    if response:
        temperature = convert_temperature(response)
        temperature_str = str(temperature) + " C"
        result_label.config(text="Current Chamber Temperature: " + temperature_str)

def read_set_point_temperature():
    response = send_query_command('R 300,1')
    if response:
        temperature = convert_temperature(response)
        temperature_str = str(temperature) + " C"
        result_label.config(text="Set Point Temperature: " + temperature_str)

def set_temperature():
    temp = temperature_entry.get()
    try:
        temp_format = convert_temperature_format(temp)
        command = 'W 300,' + temp_format
        response = device.write(command)
        temperature_str = str(temp) + " C"
        result_label.config(text="Temperature set to " + temperature_str)
    except ValueError:
        result_label.config(text="Invalid temperature value. Please try again.")

def on_exit():
    device.close()
    root.destroy()

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tipwindow = tk.Toplevel(self.widget)
        self.tipwindow.wm_overrideredirect(True)
        self.tipwindow.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tipwindow, text=self.text, justify=tk.LEFT, background="#FFFFE0", relief=tk.SOLID, borderwidth=1, font=("Helvetica", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tooltip(self, event):
        if self.tipwindow:
            self.tipwindow.destroy()

rm = pyvisa.ResourceManager()
device = rm.open_resource('GPIB0::1::INSTR')  # Replace 'GPIB0::1::INSTR' with the appropriate device address

root = tk.Tk()
root.title("Temperature Controller")
root.geometry("600x600")  # Set the initial size of the window

# Custom styles
root_style = ttk.Style()
root_style.configure('Custom.TFrame', background='#cce6ff')  # Light blue background
ttk.Style().configure('TButton', font=('Helvetica', 18), background='#4d94ff')  # Reduced font size and medium blue color for buttons
ttk.Style().configure('TLabel', font=('Helvetica', 27), foreground='#003366')  # Reduced font size and dark blue color for labels

command_label = ttk.Label(root, text="Enter your command below.\n", style='TLabel')
command_label.pack()

actual_temp_button = ttk.Button(root, text="Read actual chamber temperature", style='TButton', command=read_actual_temperature)
actual_temp_button.pack(padx=10, pady=5)

set_point_button = ttk.Button(root, text="Read set point temperature", style='TButton', command=read_set_point_temperature)
set_point_button.pack(padx=10, pady=5)

temperature_label = ttk.Label(root, text="Enter temperature value to set to:", style='TLabel')
temperature_label.pack()

temperature_entry = ttk.Entry(root, font=('Helvetica', 18))
temperature_entry.pack(padx=10, pady=5)

set_temperature_button = ttk.Button(root, text="Set the temperature", style='TButton', command=set_temperature)
set_temperature_button.pack(padx=10, pady=5)

result_label = ttk.Label(root, text="", style='TLabel')
result_label.pack()

exit_button = ttk.Button(root, text="Exit", style='TButton', command=on_exit)
exit_button.pack()

root.mainloop()