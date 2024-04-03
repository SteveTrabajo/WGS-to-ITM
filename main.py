import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
import pandas as pd
from pyproj import Proj, transform
import os
import threading
from tkinterdnd2 import DND_FILES, TkinterDnD

# Function to convert WGS84 coordinates to ITM
def wgs84_to_itm(lon, lat):
    # Define WGS84 and ITM projections
    wgs84 = Proj(init='epsg:4326')
    itm = Proj(init='epsg:2039')  # Israeli Transverse Mercator
    
    # Convert coordinates
    itm_lon, itm_lat = transform(wgs84, itm, lon, lat)
    
    return itm_lon, itm_lat

# Function to process Excel or CSV file in a separate thread
def convert_coordinates_threaded(input_file, progress_event):
    error_label.config(text="")
    try:
        # Read the data from the file
        root.drop_target_unregister()
        if input_file.lower().endswith('.csv'):
            df = pd.read_csv(input_file, header=None)
        else:
            df = pd.read_excel(input_file, header=None)
        
        # Determine the total number of lines in the file
        total_lines = len(df)
        
        # Skip the first row if it doesn't contain valid coordinate values in both the second and third columns
        if len(df.columns) >= 3:
            first_row = df.iloc[0, :]
            try:
                float(first_row[1])
                float(first_row[2])
            except ValueError:
                df = df.iloc[1:, :]  # Skip the first row
        
        # Determine the column positions
        id_column = 0  # First column
        longitude_column = None
        latitude_column = None
        
        # Check each row for valid longitude and latitude values
        for index, row in df.iterrows():
            if len(row) < 3:
                continue  # Skip rows with less than 3 elements
            
            # Skip rows with empty cells in the second and third columns
            if pd.isnull(row[1]) or pd.isnull(row[2]):
                continue
            
            # Skip rows with non-numeric values in the second and third columns
            try:
                longitude = float(row[1])
                latitude = float(row[2])
            except ValueError:
                continue
            
            # Ensure at least three columns are present and attempt conversion
            if 29 <= latitude <= 33 and 34 <= longitude <= 36:
                # Detected Longitude and Latitude columns
                longitude_column = 1
                latitude_column = 2
                break
            elif 29 <= longitude <= 33 and 34 <= latitude <= 36:
                # Detected Longitude and Latitude columns (flipped order)
                longitude_column = 2
                latitude_column = 1
                break
        
        if longitude_column is None or latitude_column is None:
            progress_bar.pack_forget()
            root.drop_target_register(DND_FILES)
            error_label.config(text="Unable to detect Longitude and Latitude columns.")
            return
        
        # Convert coordinates
        df['ITM Longitude'], df['ITM Latitude'] = zip(*df.apply(lambda row: wgs84_to_itm(float(row[longitude_column]), float(row[latitude_column])) if not pd.isnull(row[longitude_column]) and not pd.isnull(row[latitude_column]) else (None, None), axis=1))
        
        # Filter out rows where conversion failed (resulted in None for both ITM Longitude and ITM Latitude)
        df = df.dropna(subset=['ITM Longitude', 'ITM Latitude'])
        
        if df.empty:
            error_label.config(text="No valid Longitude and Latitude data found after skipping empty cells.")
            return
        
        # Add headers for Easting and Northing
        df.rename(columns={id_column: 'ID', longitude_column: 'Longitude', latitude_column: 'Latitude', 'ITM Longitude': 'Easting', 'ITM Latitude': 'Northing'}, inplace=True)
        
        # Extract original file name without extension
        original_filename = os.path.splitext(os.path.basename(input_file))[0]
        
        # Construct output file name
        output_filename = f"{original_filename}_output.csv"
        
        # Write to new CSV file with headers
        df.to_csv(output_filename, index=False, header=True)
        progress_bar.pack_forget()
        root.drop_target_register(DND_FILES)
        messagebox.showinfo("Conversion Complete", f"Conversion completed successfully. Output saved as '{output_filename}'.")
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
    finally:
        # Set the event to signal that the conversion is complete
        progress_event.set()

        # Re-enable the drag and drop functionality and the browse button
        drop_label.config(state=tk.NORMAL)
        browse_button.config(state=tk.NORMAL)

# Function to handle files dropped onto the window
def on_drop(event):
    error_label.config(text="")
    # Disable the drag and drop functionality and the browse button
    drop_label.config(state=tk.DISABLED)
    browse_button.config(state=tk.DISABLED)

    input_file = root.tk.splitlist(event.data)[0]  # Extract file path from drop event
    progress_event.clear()
    threading.Thread(target=convert_coordinates_threaded, args=(input_file, progress_event), daemon=True).start()
    # Start the progress bar loop in a separate thread
    threading.Thread(target=progress_bar_loop, daemon=True).start()
    # Show the progress bar
    progress_bar.pack(pady=10)

# Function to handle button click event for browsing files
def browse_files():
    # Reset the error label
    error_label.config(text="")
    
    # Disable the drag and drop functionality and the browse button
    drop_label.config(state=tk.DISABLED)
    browse_button.config(state=tk.DISABLED)
    
    filename = filedialog.askopenfilename(initialdir="/", title="Select a File", filetypes=[("Excel files", "*.xls *.xlsx *.xlsm *.xlsb"), ("CSV files", "*.csv")])
    if filename:
        progress_event.clear()
        threading.Thread(target=convert_coordinates_threaded, args=(filename, progress_event), daemon=True).start()
        # Start the progress bar loop in a separate thread
        threading.Thread(target=progress_bar_loop, daemon=True).start()
        # Show the progress bar
        progress_bar.pack(pady=10)

# Function to update the progress bar
def progress_bar_loop():
    while not progress_event.is_set():
        progress_bar.step(10)
        progress_bar.update_idletasks()
        progress_event.wait(0.1)  # Adjust the wait time to control the speed of the progress bar

# Create the GUI window
root = TkinterDnD.Tk()
root.title("Excel Coordinate Converter")

# Calculate the screen width and height
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Set the window size and position to center
window_width = 400
window_height = 250
x_position = (screen_width - window_width) // 2
y_position = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

# Add a label for the title
title_label = tk.Label(root, text="WGS84 to ITM Converter", font=("Calibri", 16))
title_label.pack(pady=10)

# Add a label to display errors
error_label = tk.Label(root, fg="red")
error_label.pack()

# Add a label for drag-and-drop functionality
drop_label = tk.Label(root, text="Drag and drop an Excel or CSV file here", borderwidth=2, relief="groove")
drop_label.pack(fill=tk.BOTH, expand=True)

# Bind events for drag-and-drop functionality
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)

# Add a button to browse files
browse_button = tk.Button(root, text="Browse File", command=browse_files)
browse_button.pack(pady=10)

# Add a progress bar (initially hidden)
progress_bar = Progressbar(root, orient="horizontal", mode="indeterminate", length=300, maximum=100)
progress_bar.pack(pady=5)
progress_bar.pack_forget()


# Event to signal when the conversion is complete
progress_event = threading.Event()

# Run the GUI
root.mainloop()
