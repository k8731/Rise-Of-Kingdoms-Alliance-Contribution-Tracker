import os
import re
import tkinter
from tkinter import ttk
import customtkinter
from PIL import Image, ImageTk
import cv2
import pandas as pd
import numpy as np
from tkinter import filedialog, messagebox
from tkinter.constants import *
from ocr_utils import OCR
from autocomplete_entry import AutocompleteEntry
import threading

customtkinter.set_appearance_mode(
    "Light"
)  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme(
    "dark-blue"
)  # Themes: "blue" (standard), "green", "dark-blue"
image_extensions = (".png", ".jpg", ".jpeg")

class AutoRecordApp(customtkinter.CTk):

    
    def __init__(self):
        super().__init__()
        self.button_state = True
        
        self.namelist_file = None
        self.requirement = 0
        self.current_image = None
        self.processed_image = None
        self.name_list = None # ID, Exact Name, Name

        self.image_files = None
        self.gray_image_files = None

        self.image_folder = None
        self.gray_image_folder = None

        self.data = []
        self.unmatched_data = []
        self.match_dict = (
            {}
        )  # key = matched_id, value = (matched_name, ocr_name, exact_name, number)
        self.match_flag = set()

        default_font = customtkinter.CTkFont("Segoe UI Symbol", size=16)
        default_font_bold = customtkinter.CTkFont("Segoe UI Symbol", size=18, weight="bold")
        self.ocr = OCR()

        # Configure window
        self.title("Rise Of Kingdoms: Alliance Contribution Tracker")
        self.geometry("1400x900")

        # =========================
        # Main Layout (5x3 grid)
        # =========================
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=0)
        self.grid_rowconfigure(3, weight=1)
        self.grid_rowconfigure(4, weight=0)

        # Left Panel: Basic Settings
        self.create_basic_setting_frame(default_font)
        # Right Panel: Control Panel
        self.create_control_panel_frame(default_font, default_font_bold)
        # Middle: Main Display
        self.create_main_display_frame(default_font, default_font_bold)
       
        
    def create_basic_setting_frame(self, default_font):
        # =========================
        # Left Panel: Basic Settings
        # =========================
        self.basic_setting_frame = customtkinter.CTkFrame(self)
        self.basic_setting_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        
        # ├─ Label: Basic Setting
        self.basic_setting_label = customtkinter.CTkLabel(
            self.basic_setting_frame,
            text="Basic Setting",
            font=customtkinter.CTkFont(size=20, weight="bold"),
        )
        self.basic_setting_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # ├─ Empty spacer
        self.basic_setting_frame.grid_rowconfigure(1, weight=1)
        
        # └─ Appearance Mode Setting
        self.appearance_mode_label = customtkinter.CTkLabel(
            self.basic_setting_frame,
            text="Appearance Mode:",
            anchor="w",
            font=default_font,
        )
        self.appearance_mode_label.grid(row=2, column=0, padx=20, pady=(10, 0))

        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(
            self.basic_setting_frame,
            values=["Light", "Dark", "System"],
            command=self.change_appearance_mode_event,
            font=default_font,
        )
        self.appearance_mode_optionemenu.grid(row=3, column=0, padx=20, pady=(10, 10))
        
    def create_control_panel_frame(self, default_font, default_font_bold):
        # =========================
        # Right Panel: Control Panel
        # =========================
        self.control_frame = customtkinter.CTkFrame(self)
        self.control_frame.grid(row=0, column=2, rowspan=4, sticky="nsew")
        
        # ├─ Label: Control Panel
        self.control_panel_label = customtkinter.CTkLabel(
            self.control_frame,
            text="Control Panel",
            font=customtkinter.CTkFont(size=20, weight="bold"),
        )
        self.control_panel_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # ├─ Max Player Frame
        self.max_player_frame = customtkinter.CTkFrame(self.control_frame)
        self.max_player_frame.grid(
            row=1, column=0, padx=(20, 20), pady=(20, 0), sticky="nsew"
        )
        self.max_player_label = customtkinter.CTkLabel(self.max_player_frame, text="Max Player:", font=default_font_bold)
        self.max_player_label.grid(row=0, column=0, pady=10, padx=20, sticky="n")

        self.max_player_var = tkinter.StringVar(value="300")
        self.max_player_entry = customtkinter.CTkEntry(self.max_player_frame, textvariable=self.max_player_var)
        self.max_player_entry.grid(row=0, column=1, pady=10, padx=20, sticky="n")
        
        # ├─ Buttons: Get Name List, Choose Image Folder
        self.namelist_button = customtkinter.CTkButton(
            self.control_frame,
            text="Get Name List",
            command=self.get_name_list,
            font=default_font_bold
        )
        self.namelist_button.grid(row=2, column=0, padx=10, pady=10)
        
        self.load_button = customtkinter.CTkButton(
            self.control_frame,
            text="Choose Image Folder",
            command=self.select_image_folder,
            font=default_font_bold
        )
        self.load_button.grid(row=3, column=0, padx=10, pady=10)
        
        # ├─ Radio Buttons: Choose Stats Mode
        self.radiobutton_frame = customtkinter.CTkFrame(self.control_frame)
        self.radiobutton_frame.grid(
            row=4, column=0, padx=(20, 20), pady=(20, 0), sticky="nsew"
        )
        self.radio_var = tkinter.IntVar(value=0)
        self.label_radio_group = customtkinter.CTkLabel(
            self.radiobutton_frame, text="Choose Stats Mode:", font=default_font_bold
        )
        self.label_radio_group.grid(row=0, columnspan=1, padx=10, pady=10, sticky="")
        self.radio_button_1 = customtkinter.CTkRadioButton(
            self.radiobutton_frame,
            text="alliance stats",
            variable=self.radio_var,
            value=0,
            font=default_font,
        )
        self.radio_button_1.grid(row=1, column=0, pady=10, padx=20, sticky="n")
        self.radio_button_2 = customtkinter.CTkRadioButton(
            self.radiobutton_frame,
            text="mobilization",
            variable=self.radio_var,
            value=1,
            font=default_font,
        )
        self.radio_button_2.grid(row=1, column=1, pady=10, padx=20, sticky="n")
        
        # ├─ Requirement Frame
        self.requirement_frame = customtkinter.CTkFrame(self.control_frame)
        self.requirement_frame.grid(
            row=5, column=0, padx=(20, 20), pady=(20, 0), sticky="nsew"
        )
        self.requirement_label = customtkinter.CTkLabel(self.requirement_frame, text="Requirement:", font=default_font_bold)
        self.requirement_label.grid(row=0, column=0, pady=10, padx=20, sticky="n")

        self.requirement_var = tkinter.StringVar()
        self.requirement_entry = customtkinter.CTkEntry(self.requirement_frame, textvariable=self.requirement_var)
        self.requirement_entry.grid(row=0, column=1, pady=10, padx=20, sticky="n")
        
        # ├─ Process Image Button
        self.process_image_button = customtkinter.CTkButton(
            self.control_frame,
            text="Process Image",
            command=self.process_image,
            font=default_font_bold
        )
        self.process_image_button.grid(row=6, column=0, padx=10, pady=10)

        # ├─ Change Image Frame
        self.change_image_frame = customtkinter.CTkFrame(self.control_frame)
        self.change_image_frame.grid(
            row=7, column=0, padx=(20, 20), pady=(20, 0), sticky="nsew"
        )
        self.current_image_label = customtkinter.CTkLabel(
            self.change_image_frame,
            # text="Change Image",
            text=self.get_current_image(),
            font=default_font,
        )
        self.current_image_label.grid(row=0, columnspan=1, padx=10, pady=10, sticky="")
        # next image
        self.next_image_button = customtkinter.CTkButton(
            self.change_image_frame,
            text="Next Image",
            command=self.next_image,
            font=default_font,
        )
        self.next_image_button.grid(row=1, column=0, pady=10, padx=20, sticky="n")
        # previous image
        self.previous_image_button = customtkinter.CTkButton(
            self.change_image_frame,
            text="Previous Image",
            command=self.previous_image,
            font=default_font,
        )
        self.previous_image_button.grid(row=1, column=1, pady=10, padx=20, sticky="n")

        # ├─ Extract Info Button
        self.extract_info_button = customtkinter.CTkButton(
            self.control_frame,
            text="Extract Info and Match",
            command=self.get_info,
            font=default_font_bold
        )
        self.extract_info_button.grid(row=8, column=0, padx=10, pady=10)

        # ├─ Progress Bar
        self.progress_bar = customtkinter.CTkProgressBar(self.control_frame, width=100)
        self.progress_bar.grid(row=9, column=0, pady=10)
        self.progress_bar.set(0)

        # ├─ Show Result Frame
        self.show_result_frame = customtkinter.CTkFrame(self.control_frame)
        self.show_result_frame.grid(
            row=10, column=0, padx=(20, 20), pady=(20, 0), sticky="nsew"
        )
        # show matched result
        self.show_matched_result_button = customtkinter.CTkButton(
            self.show_result_frame,
            text="Show Matched Result",
            command=self.show_result,
            font=default_font,
        )
        self.show_matched_result_button.grid(row=0, column=0, padx=10, pady=10)
        # show unmatched data
        self.show_unmatched_data_button = customtkinter.CTkButton(
            self.show_result_frame,
            text="Show Unmatched Data",
            command=self.show_unmatched_data,
            font=default_font,
        )
        self.show_unmatched_data_button.grid(row=0, column=1, padx=10, pady=10)

        # ├─ Save Excel Button
        self.save_excel_button = customtkinter.CTkButton(
            self.control_frame,
            text="Save Excel File",
            command=self.save_excel,
            font=default_font_bold
        )
        self.save_excel_button.grid(row=11, column=0, padx=10, pady=10)
        
        # └─ Reset Button
        self.reset_button = customtkinter.CTkButton(
            self.control_frame,
            text="Reset",
            command=self.reset,
            font=default_font,
        )
        self.reset_button.grid(row=12, column=0, padx=10, pady=10)

    def create_main_display_frame(self, default_font, default_font_bold):
        # ===============================
        # Right Panel: Main Display Frame
        # ===============================
        self.main_display_frame = customtkinter.CTkFrame(self)
        self.main_display_frame.grid(row=0, column=1, rowspan=5, sticky="nsew")
        self.main_display_frame.rowconfigure(3, weight=1)
        
        # ├─ Original Image Frame
        self.original_image_frame = customtkinter.CTkFrame(self.main_display_frame)
        self.original_image_frame.grid(row=0, column=0, sticky="nsew")

        self.original_image_label = customtkinter.CTkLabel(
            self.original_image_frame,
            text="No Image Loaded",
            height=300,
            width=700,
            font=default_font,
        )
        self.original_image_label.pack(padx=1, pady=1)
        self.original_image_label.pack_propagate(False)
        
        # ├─ Processed Image Frame
        self.processed_image_frame = customtkinter.CTkFrame(self.main_display_frame)
        self.processed_image_frame.grid(row=1, column=0, sticky="nsew")

        self.processed_image_label = customtkinter.CTkLabel(
            self.processed_image_frame,
            text="No Image Loaded",
            height=300,
            width=700,
            font=default_font,
        )
        self.processed_image_label.pack(padx=1, pady=1)
        
        # ├─ Output Text Frame
        self.output_text_frame = customtkinter.CTkFrame(self.main_display_frame)
        self.output_text_frame.grid(row=2, column=0, sticky="nsew")
        self.output_textbox = customtkinter.CTkTextbox(
            self.output_text_frame, width=700, height=25, font=default_font
        )
        self.output_textbox.configure(state="disabled")
        self.output_textbox.pack(padx=5, pady=5)
        
        # ├─ Table Frame
        self.table_frame = customtkinter.CTkFrame(self.main_display_frame)
        self.table_frame.grid(row=3, column=0, padx=10, pady=10)
        style = ttk.Style()
        style.configure("Treeview", font=default_font)
        style.configure("Treeview.Heading", font=default_font_bold)
        self.table = ttk.Treeview(
            self.table_frame,
            columns=["ID", "OCR Name", "Match Name", "Exact Name", "Number"],
            show="headings",
        )
        for col in ["ID", "OCR Name", "Match Name", "Exact Name", "Number"]:
            self.table.heading(col, anchor="w",text=col)
            self.table.column(col, anchor="w")
        self.update_table_columns(
            ["ID", "OCR Name", "Match Name", "Exact Name", "Number", "Diff"],
            [100, 160, 160, 160, 100, 100],
        )
        self.table.pack(expand=True, fill="both")
        self.table.bind("<Double-1>", self.on_table_double_click)
        
        # └─ add row button
        self.add_row_button = customtkinter.CTkButton(
            self.main_display_frame,
            text="Add new blank row",
            command=self.add_blank_row,
            font = default_font
        )
        self.add_row_button.grid(row=4, column=0, padx=10, pady=10)


        
    def enable_button(self):
        if self.button_state:
            self.namelist_button.configure(state="normal")
            self.load_button.configure(state="normal") 
            self.process_image_button.configure(state="normal")
            self.next_image_button.configure(state="normal")
            self.previous_image_button.configure(state="normal")
            #self.extract_info_button.configure(state="normal")
            self.show_matched_result_button.configure(state="normal")
            self.show_unmatched_data_button.configure(state="normal")
            self.save_excel_button.configure(state="normal")
            self.reset_button.configure(state="normal")   
            self.add_row_button.configure(state="normal")   
        else:
            self.namelist_button.configure(state="disabled")
            self.load_button.configure(state="disabled")
            self.process_image_button.configure(state="disabled")
            self.next_image_button.configure(state="disabled")
            self.previous_image_button.configure(state="disabled")
            self.extract_info_button.configure(state="disabled")
            self.show_matched_result_button.configure(state="disabled")
            self.show_unmatched_data_button.configure(state="disabled")
            self.save_excel_button.configure(state="disabled")
            self.reset_button.configure(state="disabled")
            self.add_row_button.configure(state="disabled")
        # configure button radius
        self.namelist_button.configure(corner_radius=8)
        self.load_button.configure(corner_radius=8)
        self.process_image_button.configure(corner_radius=8)
        self.next_image_button.configure(corner_radius=8)
        self.previous_image_button.configure(corner_radius=8)
        self.extract_info_button.configure(corner_radius=8)
        self.show_matched_result_button.configure(corner_radius=8)
        self.show_unmatched_data_button.configure(corner_radius=8)
        self.save_excel_button.configure(corner_radius=8)
        self.reset_button.configure(corner_radius=8)  
        self.add_row_button.configure(corner_radius=8)          

    def get_current_image(self):
        if not self.image_files:
            return "No Image"
        else:
            return "Current Image: " + self.image_files[self.image_count]

    def open_input_dialog_event(self):
        dialog = customtkinter.CTkInputDialog(
            text="Type in a number:", title="CTkInputDialog"
        )
        print("CTkInputDialog:", dialog.get_input())

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def show_output_textbox(self, context, msg_type="info"):
        """
        Display a message in the output textbox with color based on type.
        
        Args:
            context (str): The message to display.
            msg_type (str): Type of message: "info" (green) or "error" (red). Default is "info".
        """
        self.output_textbox.configure(state="normal")
        self.output_textbox.delete("0.0", "end")
        
        # Determine color
        if msg_type == "error":
            color = "red"
        elif msg_type == "warning":
            color = "orange"
        else:  # "info" or any other
            color = "green"
        # Insert message with color
        self.output_textbox.tag_config("colored", foreground=color)    
        self.output_textbox.insert("0.0", context + "\n", "colored")
        
        self.output_textbox.configure(state="disabled")

    def get_name_list(self):
        """
        Open a file dialog to select a rok toolkit excel file and load the name list.
        Validates the max player count and displays a success or error message.
        """
        self.namelist_file = filedialog.askopenfilename()
        if not self.namelist_file:
            return
        
        # Get max player count (default 300)
        max_player = self.max_player_var.get()
        max_player = int(max_player) if max_player.isdigit() else 300
        
        try:
            self.name_list = self.ocr.load_namelist_from_excel(self.namelist_file, max_player)
            self.show_output_textbox("Successfully obtained the name list.")
        except Exception as e:
            self.show_output_textbox(f"Failed to load name list: {e}", msg_type="error")

    def show_image(self):
        """
        Display the current image and its processed version.
        Updates labels and handles errors when opening images.
        """
        # Check if there are images to display
        if not self.image_files:
            self.show_output_textbox("No image in the folder.", msg_type="warning")
            return

        self.current_image_label.configure(text=self.get_current_image())
        # Display original image
        orig_path = os.path.join(self.image_folder, self.image_files[self.image_count])
        try:
            with Image.open(orig_path) as img:
                self.current_image = img.copy()  # keep a copy
                resized_img = self.ocr.resize_image(orig_path)
                ctk_img = customtkinter.CTkImage(light_image=resized_img, size=resized_img.size)
                self.original_image_label.configure(image=ctk_img, text="")
                self.original_image_label.image = ctk_img
        except Exception as e:
            self.show_output_textbox(f"Error reading original image: {e}", msg_type="error")
            return

        # Display processed (grayscale) image
        if self.gray_image_folder:
            gray_path = os.path.join(self.gray_image_folder, self.gray_image_files[self.image_count])
            try:
                with Image.open(gray_path) as img:
                    self.processed_image = img.copy()
                    resized_img = self.ocr.resize_image(gray_path)
                    ctk_img = customtkinter.CTkImage(light_image=resized_img, size=resized_img.size)
                    self.processed_image_label.configure(image=ctk_img, text="")
                    self.processed_image_label.image = ctk_img
            except Exception as e:
                self.show_output_textbox(f"Error reading processed image: {e}", msg_type="error")

    def select_image_folder(self):
        """
        Let user select a folder containing images, rename files, 
        load image list, and display the first image.
        """
        # Open a folder dialog for the user to select an image folder
        self.image_folder = filedialog.askdirectory()

        if not self.image_folder:
            return

        # Rename files in the folder
        self.ocr.rename_file(self.image_folder)

        # Only include image files with supported extensions
        image_extensions = (".png", ".jpg", ".jpeg")
        try:
            files = [
                f
                for f in os.listdir(self.image_folder)
                if f.lower().endswith(image_extensions)
            ]
            # Try to sort by numeric filename, fallback to alphabetical
            def sort_key(f):
                name, _ = os.path.splitext(f)
                try:
                    return int(name)
                except ValueError:
                    return float('inf'), name  # non-numeric filenames go last, sorted alphabetically

            self.image_files = sorted(files, key=sort_key)

        except Exception as e:
            self.show_output_textbox(f"Error loading images: {e}", msg_type="error")
            return

        self.show_output_textbox("Folder selected successfully.")

        # Show the first image
        self.image_count = 0
        self.show_image()

    def process_image(self):
        """
        Process images in the selected, 
        generate grayscale versions, and display the processed image.
        """
        if not self.image_folder:
            self.show_output_textbox("Please choose image folder first.", msg_type="warning")
            return
        
        # Process the images and get the folder of processed grayscale images
        try:
            self.gray_image_folder = self.ocr.process_image_folder(
                self.image_folder, self.radio_var.get()
            )
        except Exception as e:
            self.show_output_textbox(f"Error processing images: {e}", msg_type="error")
            return

        # Helper function to extract numeric value from filename for sorting
        def extract_number(filename):
            match = re.search(r"\d+", filename)
            return int(match.group()) if match else float("inf")

        try:
            self.gray_image_files = [
                f
                for f in os.listdir(self.gray_image_folder)
                if f.lower().endswith(image_extensions)
            ]
            self.gray_image_files = sorted(self.gray_image_files, key=extract_number)
        except Exception as e:
            self.show_output_textbox(f"Error loading processed images: {e}", msg_type="error")
            return
        
        # Display the processed image
        self.show_image()
        self.show_output_textbox("Images have already been processed.")

    def next_image(self):
        """
        Move to the next image in the folder and update the display and results.
        """
        if not self.image_folder:
            self.show_output_textbox("Please choose image folder first.", msg_type="warning")
            return
        if self.image_count >= (len(self.image_files) - 1):
            return
        self.image_count = self.image_count + 1

        self.show_image()
        self.show_result()

    def previous_image(self):
        """
        Move to the previous image in the folder and update the display and results.
        """
        if not self.image_folder:
            self.show_output_textbox("Please choose image folder first.", msg_type="warning")
            return
        if self.image_count <= 0:
            return
        self.image_count = self.image_count - 1
        
        self.show_image()
        self.show_result()

    def get_info(self):
        """
        Start matching data between names and processed images in a separate thread.
        Ensures preconditions are met and updates UI accordingly.
        """
        # Check prerequisites
        if not self.name_list:
            self.show_output_textbox("Please get name list first.", msg_type="warning")
            return
        if not self.image_folder:
            self.show_output_textbox("Please choose image folder first.", msg_type="warning")
            return
        if not self.gray_image_folder:
            self.show_output_textbox("Please process image first.", msg_type="warning")
            return
        
        # Disable buttons while processing
        self.button_state = False
        self.enable_button()
        
        # Validate requirement input
        self.requirement = self.requirement_var.get()
        self.requirement = int(self.requirement) if self.requirement.isdigit() else 0
        
        self.show_output_textbox("Data matching in progress...")
        
        # Start matching in a separate daemon thread to avoid blocking UI
        try:
            thread = threading.Thread(
                target=self.ocr.match_data,
                args=(
                    self.gray_image_folder,
                    self.name_list,
                    self.progress_callback,
                    self.match_complete_callback,
                ),
                daemon=True,
            )
            thread.start()
        except Exception as e:
            self.show_output_textbox(f"Error during data matching: {e}", msg_type="error")

    def progress_callback(self, value):
        self.progress_bar.set(value)
        self.update_idletasks()

    def reset(self):
        """
        Reset all data and UI states to initial values.
        Clears image info, name list, processed data, and match results.
        """
        # Clear file references
        self.namelist_file = None
        self.image_folder = None
        self.gray_image_folder = None

        # Clear image objects
        self.current_image = None
        self.processed_image = None

        # Clear data structures
        self.name_list = None
        self.image_files = None
        self.gray_image_files = None
        if self.data: self.data.clear()
        if self.unmatched_data: self.unmatched_data.clear()
        if self.match_dict: self.match_dict.clear()
        if self.match_flag: self.match_flag.clear()

        # Reset UI button
        self.extract_info_button.configure(state="normal")

        self.show_output_textbox("All data has been reset.")
        
    def match_complete_callback(self, data, unmatched_data, match_dict, match_flag):
        """
        Callback function called when data matching is finished.
        Updates internal data structures, re-enables buttons, 
        displays results, and notifies the user.
        """
        # Update internal data
        self.data = data
        self.unmatched_data = unmatched_data
        self.match_dict = match_dict
        self.match_flag = match_flag
        
        self.show_output_textbox("Data matching completed successfully.")
        
        # Re-enable buttons
        self.button_state = True
        self.enable_button()
        
        # Display result table
        self.show_result()
        
        # Provide user hint after a short delay
        self.after(1000, lambda: self.show_output_textbox('Double-click "Match Name column" or "Number" to edit.'))
        

    def update_table_columns(self, columns, widths):
        """
        Recreate the Treeview table with new columns and column widths.
        Binds double-click event to allow editing.
        """
        # Check input
        if not columns or not widths or len(columns) != len(widths):
            self.show_output_textbox("Columns and widths must be non-empty and match in length.", msg_type="warning")
            return

        # Destroy old table if exists
        try:
            self.table.destroy()
        except AttributeError:
            pass  # Table not created yet

        # Create new Treeview with specified columns
        self.table = ttk.Treeview(
            self.table_frame,
            columns=columns,
            show="headings",
        )

        # Configure columns and headings
        self.table["columns"] = columns
        for col, width in zip(columns, widths):
            self.table.heading(col, anchor="w", text=col)
            self.table.column(col, anchor="w", width=width)

        # Display the table and Bind double-click event for editing
        self.table.pack(expand=True, fill="both")
        self.table.bind("<Double-1>", self.on_table_double_click)

    def show_result(self):
        """
        Display the current image's matching results in the Treeview table.
        Updates table columns, sorts data by 'Number', and calculates 'Diff' based on requirement.
        """
        # Validate requirement input
        self.requirement = self.requirement_var.get()
        self.requirement = int(self.requirement) if self.requirement.isdigit() else 0

        # Check if data exists
        if not self.data or not self.data[self.image_count]:
            self.show_output_textbox("Please extract info first.", msg_type="warning")
            return
        
        # Clear previous table rows
        self.table.delete(*self.table.get_children())
        
        # Update table columns
        self.update_table_columns(
            ["ID", "OCR Name", "Match Name", "Exact Name", "Number", "Diff"],
            [100, 160, 160, 160, 100, 100],
        )
        
        # Sort data by 'Number' descending
        sorted_data = sorted(
            self.data[self.image_count],
            key=lambda x: int(x["Number"]),
            reverse=True,
        )
        self.data[self.image_count] = sorted_data
        
        # Insert sorted data into table
        for item in sorted_data:
            id_ = int(item.get("ID", 0))
            name = item.get("OCR Name", "")
            match_name = item.get("Match Name", "")
            exact_name = item.get("Exact Name", "")
            number = int(item.get("Number", 0))
            diff = max(0, self.requirement-number)
            self.table.insert(
                "", "end", values=(id_, name, match_name, exact_name, number, diff)
            )

    def show_unmatched_data(self):
        """
        Display unmatched OCR results in the Treeview table.
        Shows OCR Name, Number, and corresponding image filename.
        """
        # Check if data exists
        if not self.data or not self.data[self.image_count]:
            self.show_output_textbox("Please extract info first.", msg_type="warning")
            return
        
        # Clear previous table rows
        self.table.delete(*self.table.get_children())
        
        # Update table columns
        self.update_table_columns(["OCR Name", "Number", "Image"], [160, 100, 100])
        
        # Insert unmatched data into table
        for item in self.unmatched_data:
            # Defensive check: ensure item has at least 3 elements
            if len(item) < 3:
                continue
            name = item[0] if item[0] is not None else ""
            number = item[1] if item[1] is not None else 0
            image_cnt = f"{item[2]+1}.png" if item[2] is not None else ""
            self.table.insert("", "end", values=(name, number, image_cnt))
    
    def on_table_double_click(self, event):
        """
        Handle double-click on Treeview cell for inline editing.
        Supports editing 'Match Name' with autocomplete and 'Number' with numeric validation.
        Updates self.data and self.unmatched_data accordingly.
        """
        # Check if the clicked area is a valid table cell
        region = self.table.identify("region", event.x, event.y)
        if region != "cell":
            return
        
        # Get row ID (Treeview item ID)
        row_id = self.table.identify_row(event.y)
        if not row_id:
            return
        # Get row index (0-based index)
        row_index = self.table.index(row_id)
        
        # Get column ID and name
        col_id = self.table.identify_column(event.x)
        try:
            col_index = int(col_id.replace("#", "")) - 1
            col_name = self.table["columns"][col_index]
        except (ValueError, IndexError):
            return  # Invalid column

        # Only allow editing "Match Name" or "Number"
        if col_name not in ("Match Name", "Number"):
            return

        # Get the bounding box of the clicked cell
        bbox = self.table.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, width, height = bbox

        # Get the original value of the cell
        original_value = self.table.set(row_id, col_name)

        # Create appropriate entry widget
        if col_name == "Match Name":
            # Build autocomplete options (from both Name and Exact Name fields)
            name_options = list({item['Name'] for item in self.name_list} | 
                                {item['Exact Name'] for item in self.name_list})
            entry = AutocompleteEntry(self.table_frame, name_options)
            entry.var.trace_add("write", lambda *args: entry.changed())
        else:  
            # Number entry with validation
            vcmd = (self.table_frame.register(lambda P: P.isdigit() or P == ""), "%P")
            entry = tkinter.Entry(self.table_frame, validate="key", validatecommand=vcmd)

        # Place entry over the cell
        entry.row = row_index
        entry.place(x=x + self.table.winfo_x(), y=y + self.table.winfo_y(), width=width, height=height)
        entry.insert(0, original_value)
        entry.focus()

        # Confirm changes on Enter or FocusOut
        def on_entry_confirm(event):
            """
            Confirm the input and update table + data structures.
            Triggered by <Return> or losing focus.
            """
            new_value = entry.get().strip()
            
            if not new_value:
                entry.destroy()
                return
            if original_value == new_value:
                entry.destroy()
                return
            
            # Handle Number field
            if col_name == "Number":
                new_value = int(new_value) if new_value.isdigit() else 0
            # Handle Match Name field
            elif col_name == "Match Name":
                if not new_value:  # Keep original if empty
                    new_value = original_value

            # Update table and data
            self.table.set(row_id, col_name, str(new_value))
            try:
                self.data[self.image_count][row_index][col_name] = new_value
            except IndexError:
                entry.destroy()
                return

            # Update related fields for "Match Name"
            if col_name == "Match Name":
                # Find match dictionary
                match_dict = next(
                    (item for item in self.name_list 
                     if item["Name"] == new_value or item["Exact Name"] == new_value), 
                    None)
                if match_dict:
                    # Update Exact Name and ID in both data and table
                    self.data[self.image_count][row_index]["Exact Name"] = match_dict["Exact Name"]
                    self.data[self.image_count][row_index]["ID"] = match_dict["ID"]
                    self.table.set(row_id, "ID", match_dict["ID"])
                    self.table.set(row_id, "Exact Name", match_dict["Exact Name"])
                    
                    # Remove from unmatched_data
                    origin_ocr_name = self.data[self.image_count][row_index]["OCR Name"]
                    for i, (ocr_name, number, image_cnt) in enumerate(self.unmatched_data):
                        if ocr_name == origin_ocr_name and image_cnt == self.image_count:
                            self.unmatched_data.pop(i)
                            break
                        
            self.show_output_textbox(f"Update: {original_value} -> {new_value}")  
            entry.destroy()
            self.show_result()     
        
    
        # Bind confirmation actions
        entry.bind("<Return>", on_entry_confirm)
        entry.bind("<FocusOut>", on_entry_confirm)        
            
    def add_blank_row(self):
        """
        Add a new blank row to both self.data and the Treeview table.

        - Creates a default blank_data dictionary with empty/default values.
        - Appends this blank row to self.data for the current image_count.
        - Inserts a corresponding row into the Treeview with iid = row index.
        """
        # --- Default blank data structure ---
        blank_data = {
            "ID": "",
            "OCR Name": "",
            "Match Name": "",
            "Exact Name": "",
            "Number": 0
        }
        
        # Append new row to self.data
        self.data[self.image_count].append(blank_data)
        
        # Row index = position in list
        row_index = len(self.data[self.image_count]) - 1
        
        # Insert into Treeview
        self.table.insert(
            "",
            "end",
            iid=row_index,
            values=(
            blank_data.get("ID", ""),
            blank_data.get("OCR Name", ""),
            blank_data.get("Match Name", ""),
            blank_data.get("Exact Name", ""),
            str(blank_data.get("Number", 0)),
            "0"
            )
        )

    def save_excel(self):
        """
        Save the matched results into an Excel file.

        - Validates requirement input.
        - Prompts user to select a file path for saving.
        - Collects all match data from self.data across all images.
        - Calls OCR.save_excel to generate the Excel file.
        """
        
        # Validate requirement input
        self.requirement = self.requirement_var.get()
        self.requirement = int(self.requirement) if self.requirement.isdigit() else 0
        
        #Check if data exists
        if not self.data or not self.image_files:
            self.show_output_textbox("Please extract info first.", msg_type="warning")
            return
        
        # Ask user where to save the Excel file
        file_name = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Excel File",
        )
        if not file_name:  # User canceled the dialog
            self.show_output_textbox("Save cancelled.", msg_type="warning")
            return
        
        # Collect data for all images
        excel_data = []
        for i in range(len(self.image_files)):
            for item in self.data[i]:
                id = item.get("ID", "")
                name = item.get("OCR Name", "")
                match_name = item.get("Match Name", "")
                exact_name = item.get("Exact Name", "")
                number = int(item.get("Number", 0)) if str(item.get("Number", "")).isdigit() else 0
                diff = max(0, int(self.requirement)-number)
                excel_data.append((id, name, match_name, exact_name, number, diff))
                
        # Save to Excel file
        try:
            if excel_data:
                self.ocr.save_excel(file_name, excel_data)
                self.show_output_textbox(f"Data saved to Excel file: {file_name}")
            else:
                self.show_output_textbox("No data available to save.", msg_type="warning")
        except Exception as e:
            self.show_output_textbox(f"Error saving Excel file: {e}", msg_type="error")

    


if __name__ == "__main__":
    auto_record_app = AutoRecordApp()
    auto_record_app.mainloop()
