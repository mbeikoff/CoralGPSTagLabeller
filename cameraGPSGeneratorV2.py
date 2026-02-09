import pandas as pd
from PIL import Image, ImageTk
from PIL.ExifTags import TAGS
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import platform

def get_exif_datetime(file_path):
    try:
        with Image.open(file_path) as image:
            exifdata = image.getexif()
            priority_tags = ['DateTimeOriginal', 'DateTimeDigitized', 'DateTime']
            for tag_name in priority_tags:
                tag_id = next((tid for tid, tn in TAGS.items() if tn == tag_name), None)
                if tag_id is None:
                    continue
                data = exifdata.get(tag_id)
                if data:
                    dt_str = str(data).strip()
                    if ' ' in dt_str:
                        return datetime.strptime(dt_str, '%Y:%m:%d %H:%M:%S'), tag_name
                    else:
                        return datetime.strptime(dt_str[:10] + ' 00:00:00', '%Y:%m:%d %H:%M:%S'), tag_name
    except Exception as e:
        pass  # Silent error handling
    
    return None, None

class PhotoMatcherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Photo Matcher Tool")
        self.root.geometry("1200x800")
        
        self.xlsx_path = tk.StringVar()
        self.threshold_var = tk.StringVar(value="0")
        self.threshold = 0.0
        self.output_base = tk.StringVar()
        self.camera_map = {'PDP1': 'Green', 'PDP2': 'White', 'PDP3': 'Third'}
        self.colors = ['Green', 'White', 'Third']
        self.dir_vars = {color: tk.StringVar() for color in self.colors}
        self.photos = {color: [] for color in self.colors}
        self.photo_paths = {color: [] for color in self.colors}
        self.selected = {color: {} for color in self.colors}
        self.df = None
        self.matched_df = None
        self.matched_photos = {}
        
        self.initial_frame = None
        self.photo_select_frame = None
        self.preview_frame = None
        self.tree = None
        self.selectable_frames = {}
        self.canvases = {}
        self.camera_frames = {}
        self.current_camera = None
        self.camera_container = None
        self.button_frame = None
        self.camera_label = None
        self.match_progress = None
        
        self.sys_platform = platform.system()
        
        self.create_widgets()
    
    def create_widgets(self):
        self.initial_frame = tk.Frame(self.root)
        self.initial_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(self.initial_frame, text="Select XLSX File:", font=("Arial", 10, "bold")).pack(pady=10)
        entry_frame1 = tk.Frame(self.initial_frame)
        entry_frame1.pack(pady=5)
        tk.Entry(entry_frame1, textvariable=self.xlsx_path, width=50, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(entry_frame1, text="Browse XLSX", command=self.select_xlsx).pack(side=tk.RIGHT, padx=5)
        
        for color in self.colors:
            tk.Label(self.initial_frame, text=f"Select {color} Photo Directory:", font=("Arial", 10, "bold")).pack(pady=10)
            entry_frame = tk.Frame(self.initial_frame)
            entry_frame.pack(pady=5)
            tk.Entry(entry_frame, textvariable=self.dir_vars[color], width=50, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True)
            tk.Button(entry_frame, text=f"Browse {color} Folder", command=lambda c=color: self.select_dir(c)).pack(side=tk.RIGHT, padx=5)
        
        tk.Label(self.initial_frame, text="Match Threshold (seconds):", font=("Arial", 10, "bold")).pack(pady=10)
        entry_frame3 = tk.Frame(self.initial_frame)
        entry_frame3.pack(pady=5)
        tk.Entry(entry_frame3, textvariable=self.threshold_var, width=10).pack(side=tk.LEFT)
        tk.Label(entry_frame3, text="seconds (0 for exact match)").pack(side=tk.LEFT, padx=(5, 0))
        
        tk.Label(self.initial_frame, text="Output Base Name:", font=("Arial", 10, "bold")).pack(pady=10)
        tk.Label(self.initial_frame, text="(threshold & .xlsx will be appended; leave empty for default)", font=("Arial", 9)).pack(pady=(0,5))
        entry_frame4 = tk.Frame(self.initial_frame)
        entry_frame4.pack(pady=5)
        tk.Entry(entry_frame4, textvariable=self.output_base, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tk.Button(self.initial_frame, text="Load Photos for Selection", command=self.load_photos, bg="lightgreen", font=("Arial", 12, "bold")).pack(pady=20)
        
        self.progress = ttk.Progressbar(self.initial_frame, orient='horizontal', length=400, mode='determinate')
        self.progress.pack(pady=10)
        
        self.status_label = tk.Label(self.initial_frame, text="", font=("Arial", 9))
        self.status_label.pack(pady=5)
    
    def select_xlsx(self):
        filename = filedialog.askopenfilename(
            title="Select XLSX File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.xlsx_path.set(filename)
            base_name = os.path.splitext(os.path.basename(filename))[0] + '_matched'
            self.output_base.set(base_name)
    
    def select_dir(self, color):
        directory = filedialog.askdirectory(title=f"Select {color} Photo Directory")
        if directory:
            self.dir_vars[color].set(directory)
    
    def bind_mousewheel(self, widget, canvas):
        """Recursively bind mousewheel to widget and all its children"""
        def on_mousewheel(event):
            if hasattr(event, 'delta') and event.delta != 0:
                delta = int(-1 * (event.delta / 120))
            elif event.num == 4:
                delta = -1
            elif event.num == 5:
                delta = 1
            else:
                return
            canvas.yview_scroll(delta, "units")
        
        widget.bind("<MouseWheel>", on_mousewheel)
        widget.bind("<Button-4>", on_mousewheel)
        widget.bind("<Button-5>", on_mousewheel)
        
        # Recursively bind to all children
        for child in widget.winfo_children():
            self.bind_mousewheel(child, canvas)
    
    def select_all(self, color):
        for var in self.selected[color].values():
            var.set(True)
    
    def deselect_all(self, color):
        for var in self.selected[color].values():
            var.set(False)
    
    def make_thumbnail(self, file_path, size=(200, 200)):
        try:
            with Image.open(file_path) as img:
                img.thumbnail(size, Image.Resampling.LANCZOS)
                return ImageTk.PhotoImage(img)
        except:
            return None
    
    def parse_time_robust(self, raw_value):
        """Try multiple parsing methods and return best datetime or error info."""
        if pd.isna(raw_value):
            return None, "Empty/NaN value"
        
        raw_str = str(raw_value).strip()
        
        if isinstance(raw_value, (pd.Timestamp, datetime)):
            return raw_value, None
        
        try:
            num = float(raw_value)  # More aggressive
            dt = pd.to_datetime(num, origin=pd.Timestamp('1899-12-30'), unit='D', errors='coerce')
            if pd.notna(dt):
                return dt, None
        except:
            pass
        
        # Method 3: String formats (AU-specific, including partial times)
        formats_to_try = [
            '%d/%m/%Y %H:%M:%S',      # Full 24h
            '%d/%m/%Y %I:%M:%S %p',   # Full 12h AM/PM
            '%d/%m/%Y %H:%M',         # Partial 24h (adds :00)
            '%d/%m/%Y %I:%M %p',      # Partial 12h (adds :00)
            '%d/%m/%Y %H:%M:%S.%f',   # With microseconds if any
            '%d/%m/%Y %I:%M:%S.%f %p',
            '%d/%m/%Y',               # Date only (00:00:00)
            '%d/%m/%y %H:%M:%S',
            '%d/%m/%y %I:%M:%S %p',
            '%d/%m/%y %H:%M',
            '%d/%m/%y %I:%M %p'
        ]
        for fmt in formats_to_try:
            try:
                dt = pd.to_datetime(raw_str, format=fmt, dayfirst=True, errors='coerce')
                if pd.notna(dt):
                    return dt, None
            except:
                pass
        
        try:
            dt = pd.to_datetime(raw_str, dayfirst=True, infer_datetime_format=True, errors='coerce')
            if pd.notna(dt):
                return dt, None
        except:
            pass
        
        if ':' in raw_str and len(raw_str.split(':')) == 2:  # HH:MM only
            try:
                parts = raw_str.split()
                date_part = parts[0]
                time_part = parts[1] + ':00'
                full_str = f"{date_part} {time_part}"
                dt = pd.to_datetime(full_str, dayfirst=True, errors='coerce')
                if pd.notna(dt):
                    return dt, None
            except:
                pass
        
        error_msg = f"Failed all parses (raw: '{raw_str}') | Type: {type(raw_value)}"
        return None, error_msg

    def load_photos(self):
        xlsx_file = self.xlsx_path.get()
        if not xlsx_file:
            messagebox.showerror("Error", "Please select XLSX file.")
            return
        
        dirs = {color: self.dir_vars[color].get() for color in self.colors}
        if not all(dirs.values()):
            messagebox.showerror("Error", "Please select all three photo directories.")
            return
        
        threshold_str = self.threshold_var.get().strip()
        try:
            self.threshold = float(threshold_str) if threshold_str else 0.0
        except ValueError:
            messagebox.showerror("Error", "Invalid threshold value. Please enter a number (e.g., 30 for 30 seconds).")
            return
        
        try:
            self.df = pd.read_excel(xlsx_file, sheet_name=0, engine='openpyxl')
            if 'Time' not in self.df.columns or 'Camera' not in self.df.columns:
                raise ValueError("Column 'Time' or 'Camera' not found in the XLSX file.")
            
            # Strip whitespace from Camera column
            self.df['Camera'] = self.df['Camera'].astype(str).str.strip()
            
            self.status_label.config(text="Parsing times...")
            self.root.update()
            
            self.df['Time_parsed'] = None
            self.df['Parse_Error'] = ''
            invalid_count = 0
            for idx, row in self.df.iterrows():
                dt, error = self.parse_time_robust(row['Time'])
                self.df.at[idx, 'Time_parsed'] = dt
                self.df.at[idx, 'Parse_Error'] = error if error else ''
                if error:
                    invalid_count += 1
            
            self.df['Time'] = pd.to_datetime(self.df['Time_parsed'], errors='coerce')
            self.df.drop('Time_parsed', axis=1, inplace=True)
            
            cameras = self.df['Camera'].dropna().unique()
            invalid_cameras = [cam for cam in cameras if cam not in self.camera_map]
            if invalid_cameras:
                raise ValueError(f"Unknown cameras found: {invalid_cameras}. Expected PDP1, PDP2, PDP3.")
            
            self.progress['maximum'] = 1
            self.root.update()
            self.status_label.config(text="Loading photos...")
            self.root.update()
            
            # Calculate total files
            total_files = 0
            photo_count = {color: 0 for color in self.colors}
            skipped_count = {color: 0 for color in self.colors}
            for color in self.colors:
                photo_dir = dirs[color]
                count = sum(1 for _, _, files in os.walk(photo_dir) for f in files if f.lower().endswith(('.jpg', '.jpeg')))
                total_files += count
            
            processed = 0
            self.progress['maximum'] = total_files
            self.progress['value'] = 0
            
            self.photos = {color: [] for color in self.colors}
            self.photo_paths = {color: [] for color in self.colors}
            
            for color in self.colors:
                photo_dir = dirs[color]
                for root, _, files in os.walk(photo_dir):
                    for file in files:
                        if file.lower().endswith(('.jpg', '.jpeg')):
                            full_path = os.path.join(root, file)
                            photo_time, source = get_exif_datetime(full_path)
                            if photo_time:
                                subfolder = os.path.relpath(root, photo_dir)
                                if subfolder == '.':
                                    subfolder = ''
                                self.photos[color].append((photo_time, subfolder, file, source))
                                self.photo_paths[color].append(full_path)
                                photo_count[color] += 1
                            else:
                                skipped_count[color] += 1
                            
                            processed += 1
                            self.progress['value'] = processed
                            self.root.update_idletasks()
            
            self.progress['value'] = self.progress['maximum']
            self.root.update()
            
            # Sort photos by time
            for color in self.colors:
                self.photos[color].sort(key=lambda x: x[0])
            
            # Build photo selection frame
            if self.photo_select_frame is None:
                self.photo_select_frame = tk.Frame(self.root)
                
                # Instructions
                instr_label = tk.Label(self.photo_select_frame, text="Select photos with yellow tags (none selected by default). Navigate between cameras using the buttons below.", 
                                       font=("Arial", 10), wraplength=1100)
                instr_label.pack(pady=10)
                
                # Camera container
                self.camera_container = tk.Frame(self.photo_select_frame)
                self.camera_container.pack(fill=tk.BOTH, expand=True)
                
                # Create camera frames
                self.camera_frames = {}
                self.selectable_frames = {}
                self.canvases = {}
                for color in self.colors:
                    camera_frame = tk.Frame(self.camera_container)
                    
                    title_label = tk.Label(camera_frame, text=f"{color} Photos ({photo_count[color]} images, {skipped_count[color]} skipped)", 
                                           font=("Arial", 12, "bold"))
                    title_label.pack(pady=10)
                    
                    toggle_frame = tk.Frame(camera_frame)
                    toggle_frame.pack(pady=5)
                    tk.Button(toggle_frame, text="Select All", command=lambda c=color: self.select_all(c)).pack(side=tk.LEFT, padx=5)
                    tk.Button(toggle_frame, text="Deselect All", command=lambda c=color: self.deselect_all(c)).pack(side=tk.LEFT, padx=5)
                    
                    scroll_frame = tk.Frame(camera_frame)
                    scroll_frame.pack(fill=tk.BOTH, expand=True)
                    
                    canvas = tk.Canvas(scroll_frame, bg='white')
                    v_scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
                    canvas.configure(yscrollcommand=v_scrollbar.set)
                    
                    canvas.pack(side="left", fill="both", expand=True)
                    v_scrollbar.pack(side="right", fill="y")
                    
                    inner_frame = tk.Frame(canvas)
                    canvas.create_window((0, 0), window=inner_frame, anchor="nw")
                    
                    inner_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
                    
                    self.canvases[color] = canvas
                    self.selectable_frames[color] = inner_frame
                    
                    self.selected[color] = {}
                    grid_row = 0
                    col = 0
                    for i, path in enumerate(self.photo_paths[color]):
                        var = tk.BooleanVar(value=False)
                        self.selected[color][path] = var
                        
                        thumb = self.make_thumbnail(path)
                        if thumb:
                            img_label = tk.Label(inner_frame, image=thumb, borderwidth=1, relief="solid")
                            img_label.image = thumb  # Keep reference
                            img_label.grid(row=grid_row, column=col * 2, padx=10, pady=10)
                        
                        filename = os.path.basename(path)
                        short_name = filename[:30] + '...' if len(filename) > 30 else filename
                        chk = tk.Checkbutton(inner_frame, variable=var, text=short_name, anchor='w')
                        chk.grid(row=grid_row, column=col * 2 + 1, sticky='ew', padx=10, pady=10)
                        
                        col += 1
                        if col == 3:  # 3 images per row
                            col = 0
                            grid_row += 1
                    
                    # Configure columns for even spacing
                    for j in range(6):
                        inner_frame.grid_columnconfigure(j, weight=1)
                    
                    # Bind mousewheel to camera_frame and all its children
                    self.bind_mousewheel(camera_frame, canvas)
                    
                    self.camera_frames[color] = camera_frame
                
                # Button frame
                self.button_frame = tk.Frame(self.photo_select_frame)
                self.button_frame.pack(fill=tk.X, pady=10)
                
                tk.Button(self.button_frame, text="Back to Initial", command=lambda: self.show_frame('initial'), 
                          bg="lightgray", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
                
                tk.Button(self.button_frame, text="Previous Camera", command=self.prev_camera).pack(side=tk.LEFT, padx=5)
                
                self.camera_label = tk.Label(self.button_frame, text="", font=("Arial", 10, "bold"))
                self.camera_label.pack(side=tk.LEFT, padx=10)
                
                tk.Button(self.button_frame, text="Next Camera", command=self.next_camera).pack(side=tk.LEFT, padx=5)
                
                tk.Button(self.button_frame, text="Proceed to Matching", command=self.match_photos, 
                          bg="lightblue", font=("Arial", 12, "bold")).pack(side=tk.RIGHT, padx=10)
            
            self.initial_frame.pack_forget()
            self.photo_select_frame.pack(fill=tk.BOTH, expand=True)
            self.current_camera = 'Green'
            self.show_camera(self.current_camera)
            self.status_label.config(text="")
            
        except Exception as e:
            self.status_label.config(text="Error occurred!")
            messagebox.showerror("Error", f"An error occurred: {str(e)}\nCheck console for details.")
    
    def show_camera(self, color):
        for c in self.colors:
            self.camera_frames[c].pack_forget()
        self.camera_frames[color].pack(in_=self.camera_container, fill=tk.BOTH, expand=True)
        self.camera_label.config(text=f"{color} Photos")
    
    def next_camera(self):
        idx = self.colors.index(self.current_camera)
        next_idx = (idx + 1) % len(self.colors)
        self.current_camera = self.colors[next_idx]
        self.show_camera(self.current_camera)
    
    def prev_camera(self):
        idx = self.colors.index(self.current_camera)
        prev_idx = (idx - 1) % len(self.colors)
        self.current_camera = self.colors[prev_idx]
        self.show_camera(self.current_camera)
    
    def match_photos(self):
        # Filter selected photos
        self.matched_photos = {}
        for color in self.colors:
            self.matched_photos[color] = []
            for i, (p_time, subfolder, filename, source) in enumerate(self.photos[color]):
                path = self.photo_paths[color][i]
                if self.selected[color][path].get():
                    self.matched_photos[color].append((p_time, subfolder, filename, source))
        
        # Match
        self.df['Subfolder'] = ''
        self.df['Filename'] = ''
        
        matched_count = 0
        no_match_count = 0
        total_rows = len(self.df)
        
        if self.match_progress is None:
            self.match_progress = ttk.Progressbar(self.photo_select_frame, orient='horizontal', length=400, mode='determinate')
            self.match_progress.pack(in_=self.photo_select_frame, before=self.button_frame, fill=tk.X, padx=10, pady=5)
        
        self.match_progress['maximum'] = total_rows
        self.match_progress['value'] = 0
        self.status_label.config(text=f"Matching photos (Threshold: {self.threshold}s)...")
        self.root.update()
        
        for idx, row in self.df.iterrows():
            gps_time = row['Time']
            
            if pd.isna(gps_time):
                no_match_count += 1
            else:
                cam = str(row['Camera']).strip()
                color = self.camera_map.get(cam)
                if not color:
                    no_match_count += 1
                else:
                    photos_list = self.matched_photos.get(color, [])
                    if not photos_list:
                        no_match_count += 1
                    else:
                        min_diff_seconds = float('inf')
                        closest_subfolder = ''
                        closest_filename = ''
                        
                        for p_time, subfolder, filename, source in photos_list:
                            diff = abs((p_time - gps_time).total_seconds())
                            if diff < min_diff_seconds:
                                min_diff_seconds = diff
                                closest_subfolder = subfolder
                                closest_filename = filename
                        
                        if min_diff_seconds <= self.threshold:
                            matched_count += 1
                            self.df.at[idx, 'Subfolder'] = closest_subfolder
                            self.df.at[idx, 'Filename'] = closest_filename
                        else:
                            no_match_count += 1
            self.match_progress['value'] += 1
            self.root.update_idletasks()
        
        self.match_progress['value'] = self.match_progress['maximum']
        self.root.update()
        
        self.matched_df = self.df[self.df['Filename'] != ''].copy()
        
        # Build preview frame if not exists
        if self.preview_frame is None:
            self.preview_frame = tk.Frame(self.root)
            tk.Label(self.preview_frame, text=f"Preview: {len(self.matched_df)} matches found", 
                     font=("Arial", 12, "bold")).pack(pady=10)
            
            tree_frame = tk.Frame(self.preview_frame)
            tree_frame.pack(fill=tk.BOTH, expand=True)
            
            self.tree = ttk.Treeview(tree_frame)
            v_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
            h_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
            self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
            
            v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
            self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            button_frame = tk.Frame(self.preview_frame)
            button_frame.pack(fill=tk.X, pady=10)
            tk.Button(button_frame, text="Export Matches", command=self.export_matches, 
                      bg="lightgreen", font=("Arial", 12, "bold")).pack(side=tk.LEFT, padx=10)
            tk.Button(button_frame, text="Back to Photo Selection", 
                      command=lambda: self.show_frame('select'), bg="lightgray", font=("Arial", 10)).pack(side=tk.LEFT, padx=10)
            tk.Button(button_frame, text="Back to Initial", 
                      command=lambda: self.show_frame('initial'), bg="lightgray", font=("Arial", 10)).pack(side=tk.RIGHT, padx=10)
        
        # Populate tree
        self.tree.delete(*self.tree.get_children())
        if len(self.matched_df) > 0:
            columns = list(self.matched_df.columns)
            self.tree['columns'] = columns
            self.tree.heading('#0', text='')  # No tree column
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, minwidth=50)
            
            for _, row_data in self.matched_df.iterrows():
                self.tree.insert('', tk.END, values=[str(val) for val in row_data])
        else:
            self.tree['columns'] = ()
            messagebox.showinfo("No Matches", "No matches found with the selected photos and threshold.")
        
        self.photo_select_frame.pack_forget()
        self.preview_frame.pack(fill=tk.BOTH, expand=True)
        self.status_label.config(text="Review matches and export.")
    
    def export_matches(self):
        if self.matched_df is None or self.matched_df.empty:
            messagebox.showwarning("No Matches", "No matches to export.")
            return
        
        output_base_input = self.output_base.get().strip()
        input_base, input_ext = os.path.splitext(os.path.basename(self.xlsx_path.get()))
        base = output_base_input if output_base_input else input_base
        base, ext = os.path.splitext(base)
        output_filename = f"{base}_threshold_{self.threshold}s{ext or '.xlsx'}"
        output_file = os.path.join(os.path.dirname(self.xlsx_path.get()), output_filename)
        
        self.matched_df.to_excel(output_file, index=False)
        
        messagebox.showinfo("Success", 
            f"Exported {len(self.matched_df)} matches to:\n{output_file}\n"
            f"Threshold used: {self.threshold} seconds"
        )
    
    def show_frame(self, frame_name):
        if frame_name == 'initial':
            if self.preview_frame:
                self.preview_frame.pack_forget()
            if self.photo_select_frame:
                self.photo_select_frame.pack_forget()
            self.initial_frame.pack(fill=tk.BOTH, expand=True)
            self.status_label.config(text="")
            self.progress['value'] = 0
            if self.match_progress is not None:
                self.match_progress.pack_forget()
                self.match_progress.destroy()
                self.match_progress = None
        elif frame_name == 'select':
            if self.preview_frame:
                self.preview_frame.pack_forget()
            if self.initial_frame:
                self.initial_frame.pack_forget()
            self.photo_select_frame.pack(fill=tk.BOTH, expand=True)
            if hasattr(self, 'camera_container'):
                self.camera_container.pack(fill=tk.BOTH, expand=True)
            if self.match_progress is not None:
                self.match_progress.pack(in_=self.photo_select_frame, before=self.button_frame, fill=tk.X, padx=10, pady=5)
            if self.current_camera is None:
                self.current_camera = 'Green'
            self.show_camera(self.current_camera)

if __name__ == "__main__":
    root = tk.Tk()
    app = PhotoMatcherGUI(root)
    root.mainloop()