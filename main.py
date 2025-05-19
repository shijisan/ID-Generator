import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk, ImageDraw, ImageFont
import os
import uuid
import qrcode
import pandas as pd
import csv
from datetime import datetime

class Recipient:
    def __init__(self, field1, field2, field3, image_path):
        self.field1 = field1
        self.field2 = field2
        self.field3 = field3
        self.image_path = image_path
        self.uuid = str(uuid.uuid4())

class IDGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ID Card Generator")
        self.root.geometry("800x600")
        
        # Check for required packages
        self.check_required_packages()
        
        # Variables
        self.recipients = []
        self.background_image_path = None
        
        # Create UI
        self.create_ui()
    
    def check_required_packages(self):
        """Check if openpyxl is installed and show a warning if not"""
        try:
            import openpyxl
            self.excel_support = True
        except ImportError:
            self.excel_support = False
            messagebox.showinfo("Package Information", 
                               "Note: The 'openpyxl' package is not installed.\n\n"
                               "The program will still work, but data will be exported as CSV instead of Excel.\n\n"
                               "To enable Excel export, install openpyxl using:\n"
                               "pip install openpyxl")
    
    def create_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input frame (left side)
        input_frame = ttk.LabelFrame(main_frame, text="Add Recipient", padding="10")
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Field 1 (Required)
        ttk.Label(input_frame, text="Field 1 (Required):").grid(column=0, row=0, sticky=tk.W, pady=5)
        self.field1_entry = ttk.Entry(input_frame, width=30)
        self.field1_entry.grid(column=1, row=0, sticky=tk.W, pady=5)
        
        # Field 2
        ttk.Label(input_frame, text="Field 2:").grid(column=0, row=1, sticky=tk.W, pady=5)
        self.field2_entry = ttk.Entry(input_frame, width=30)
        self.field2_entry.grid(column=1, row=1, sticky=tk.W, pady=5)
        
        # Field 3
        ttk.Label(input_frame, text="Field 3:").grid(column=0, row=2, sticky=tk.W, pady=5)
        self.field3_entry = ttk.Entry(input_frame, width=30)
        self.field3_entry.grid(column=1, row=2, sticky=tk.W, pady=5)
        
        # Profile Image
        ttk.Label(input_frame, text="Profile Image:").grid(column=0, row=3, sticky=tk.W, pady=5)
        self.profile_image_button = ttk.Button(input_frame, text="Select Image", command=self.select_profile_image)
        self.profile_image_button.grid(column=1, row=3, sticky=tk.W, pady=5)
        self.profile_image_path = None
        self.profile_image_label = ttk.Label(input_frame, text="No image selected")
        self.profile_image_label.grid(column=0, row=4, columnspan=2, sticky=tk.W, pady=5)
        
        # Add Recipient Button
        self.add_recipient_button = ttk.Button(input_frame, text="Add Recipient", command=self.add_recipient)
        self.add_recipient_button.grid(column=0, row=5, columnspan=2, pady=10)
        
        # Background Image Button
        self.bg_image_button = ttk.Button(input_frame, text="Select Background Image", command=self.select_background_image)
        self.bg_image_button.grid(column=0, row=6, columnspan=2, pady=5)
        self.bg_image_label = ttk.Label(input_frame, text="No background selected")
        self.bg_image_label.grid(column=0, row=7, columnspan=2, sticky=tk.W, pady=5)
        
        # Generate Button
        self.generate_button = ttk.Button(input_frame, text="Generate ID Cards", command=self.generate_ids)
        self.generate_button.grid(column=0, row=8, columnspan=2, pady=10)
        
        # Recipients list frame (right side)
        list_frame = ttk.LabelFrame(main_frame, text="Recipients List", padding="10")
        list_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Treeview for recipients
        columns = ('field1', 'field2', 'field3')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings')
        self.tree.heading('field1', text='Field 1')
        self.tree.heading('field2', text='Field 2')
        self.tree.heading('field3', text='Field 3')
        self.tree.column('field1', width=100)
        self.tree.column('field2', width=100)
        self.tree.column('field3', width=100)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Remove button
        self.remove_button = ttk.Button(list_frame, text="Remove Selected", command=self.remove_recipient)
        self.remove_button.pack(pady=10)
    
    def select_profile_image(self):
        self.profile_image_path = filedialog.askopenfilename(
            title="Select Profile Image",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg")]
        )
        
        if self.profile_image_path:
            self.profile_image_label.config(text=os.path.basename(self.profile_image_path))
    
    def select_background_image(self):
        self.background_image_path = filedialog.askopenfilename(
            title="Select Background Image",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg")]
        )
        
        if self.background_image_path:
            self.bg_image_label.config(text=os.path.basename(self.background_image_path))
    
    def add_recipient(self):
        field1 = self.field1_entry.get()
        field2 = self.field2_entry.get()
        field3 = self.field3_entry.get()
        
        if not field1:
            messagebox.showwarning("Input Error", "Field 1 is required.")
            return
        
        if not self.profile_image_path:
            messagebox.showwarning("Input Error", "Please select a profile image.")
            return
        
        # Add to recipients list
        recipient = Recipient(field1, field2, field3, self.profile_image_path)
        self.recipients.append(recipient)
        
        # Add to treeview
        self.tree.insert('', tk.END, values=(field1, field2, field3))
        
        # Clear form
        self.field1_entry.delete(0, tk.END)
        self.field2_entry.delete(0, tk.END)
        self.field3_entry.delete(0, tk.END)
        self.profile_image_path = None
        self.profile_image_label.config(text="No image selected")
    
    def remove_recipient(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Selection Error", "Please select a recipient to remove.")
            return
        
        # Get the index of the selected item
        index = self.tree.index(selected_item[0])
        
        # Remove from the list and treeview
        self.recipients.pop(index)
        self.tree.delete(selected_item)
    
    def generate_ids(self):
        if not self.background_image_path:
            messagebox.showwarning("Input Error", "Please select a background image.")
            return
        
        if not self.recipients:
            messagebox.showwarning("Input Error", "Please add at least one recipient.")
            return
        
        # Create output directory with datetime
        now = datetime.now()
        timestamp = now.strftime("%Y%m%d_%H%M%S")
        output_dir = f"Generated_IDs_{timestamp}"
        os.makedirs(output_dir, exist_ok=True)
        
        # Excel data
        excel_data = []
        
        for recipient in self.recipients:
            # Generate QR code
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            qr.add_data(recipient.uuid)
            qr.make(fit=True)
            qr_img = qr.make_image(fill_color="black", back_color="white")
            qr_img = qr_img.resize((100, 100))
            
            # Create ID card
            try:
                # Load and resize background image
                background = Image.open(self.background_image_path).convert("RGBA")
                background = background.resize((400, 600), Image.Resampling.LANCZOS)
                
                # Load and resize profile image
                profile_img = Image.open(recipient.image_path).convert("RGBA")
                profile_size = (200, 200)
                profile_img = profile_img.resize(profile_size, Image.Resampling.LANCZOS)
                
                # Calculate positions
                profile_pos = ((400 - profile_size[0]) // 2, 50)  # Center in top half
                qr_pos = ((400 - 100) // 2, 450)  # Bottom position for QR
                
                # Create a drawing context
                draw = ImageDraw.Draw(background)
                
                # Use a default font
                try:
                    font = ImageFont.truetype("arial.ttf", size=20)
                except IOError:
                    font = ImageFont.load_default()
                
                # Paste profile image
                background.paste(profile_img, profile_pos, profile_img)
                
                # Paste QR code
                background.paste(qr_img, qr_pos)
                
                # Add fields
                field_y = 300
                for field, value in [
                    ("Field 1:", recipient.field1),
                    ("Field 2:", recipient.field2),
                    ("Field 3:", recipient.field3)
                ]:
                    if value:  # Only add non-empty fields
                        text = f"{field} {value}"
                        # Calculate text position for centering
                        text_bbox = draw.textbbox((0, 0), text, font=font)
                        text_width = text_bbox[2] - text_bbox[0]
                        text_height = text_bbox[3] - text_bbox[1]
                        text_x = (400 - text_width) // 2
                        
                        # Draw white background with black border for text
                        padding = 10  # Padding around text
                        rect_x0 = text_x - padding
                        rect_y0 = field_y - padding
                        rect_x1 = text_x + text_width + padding
                        rect_y1 = field_y + text_height + padding
                        
                        # Draw white background rectangle
                        draw.rectangle([rect_x0, rect_y0, rect_x1, rect_y1], 
                                      fill="white", 
                                      outline="black", 
                                      width=2)
                        
                        # Draw text on top of white background
                        draw.text((text_x, field_y), text, fill="black", font=font)
                        field_y += 60  # Increased spacing between fields
                
                # Save the ID card
                file_name = f"{recipient.field1.replace(' ', '_')}.png"
                output_path = os.path.join(output_dir, file_name)
                background.save(output_path)
                
                # Add data for Excel
                excel_data.append({
                    'UUID': recipient.uuid,
                    'Field1': recipient.field1,
                    'Field2': recipient.field2,
                    'Field3': recipient.field3,
                    'ImagePath': recipient.image_path,
                    'OutputPath': output_path
                })
                
            except Exception as e:
                messagebox.showerror("Error", f"Error processing {recipient.field1}: {str(e)}")
        
        # Create Excel file
        if excel_data:
            df = pd.DataFrame(excel_data)
            excel_path = os.path.join(output_dir, "recipients_database.xlsx")
            csv_path = os.path.join(output_dir, "recipients_database.csv")
            
            try:
                # Try to save as Excel first
                try:
                    df.to_excel(excel_path, index=False)
                    messagebox.showinfo("Success", f"Generated {len(self.recipients)} ID cards in '{output_dir}' folder.\nExcel database saved as 'recipients_database.xlsx'.")
                except ImportError as e:
                    if "openpyxl" in str(e):
                        # If openpyxl is missing, save as CSV instead and inform the user
                        df.to_csv(csv_path, index=False)
                        messagebox.showinfo("Success - CSV Only", 
                                           f"Generated {len(self.recipients)} ID cards in '{output_dir}' folder.\n\n"
                                           f"Note: Excel file could not be created because 'openpyxl' module is missing.\n"
                                           f"Data has been saved as CSV instead at 'recipients_database.csv'.\n\n"
                                           f"To enable Excel export, install openpyxl using:\n"
                                           f"pip install openpyxl")
                    else:
                        raise e
            except Exception as e:
                # Last resort: manual CSV creation
                try:
                    with open(csv_path, 'w', newline='') as csvfile:
                        writer = csv.DictWriter(csvfile, fieldnames=excel_data[0].keys())
                        writer.writeheader()
                        writer.writerows(excel_data)
                    
                    messagebox.showinfo("Success - CSV Only", 
                                       f"Generated {len(self.recipients)} ID cards in '{output_dir}' folder.\n\n"
                                       f"Note: Excel file could not be created due to an error: {str(e)}\n"
                                       f"Data has been saved as CSV instead at 'recipients_database.csv'.")
                except Exception as csv_error:
                    messagebox.showwarning("Warning", 
                                         f"Generated {len(self.recipients)} ID cards in '{output_dir}' folder, "
                                         f"but could not create database file due to errors: \n{str(e)}\n{str(csv_error)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = IDGeneratorApp(root)
    root.mainloop()