import os
import platform
import shutil
import pytesseract
import cv2
from PIL import Image, ImageTk
import pandas as pd
import numpy as np
import re
import shutil, stat
import unicodedata
from fuzzywuzzy import process

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


class OCR:
    def __init__(self):
        super().__init__()
        self.data = []
        self.unmatched_data = []
        self.fail_extract_data = []
        self.match_dict = (
            {}
        )  # key = matched_id, value = (matched_name, ocr_name, exact_name, number, image_cnt)
        self.match_flag = set()
        
        system = platform.system()
        if system == "Windows":
            default_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        elif system == "Darwin":  # macOS
            default_path = "/usr/local/bin/tesseract"
        else:  # Linux
            default_path = "/usr/bin/tesseract"
        tesseract_path = os.getenv("TESSERACT_PATH", default_path)
        if not os.path.exists(tesseract_path):
            raise FileNotFoundError("[ERROR] Can not find Tesseract exe file: {tesseract_path}")
        pytesseract.pytesseract.tesseract_cmd = tesseract_path



    def _normalize_text(self, text):
        """
        Normalize text by removing spaces, punctuation, symbols, and superscripts,
        while keeping all kinds of letters (including non-Latin) and digits.
        """
        smallcaps_map = {
            'ᴀ': 'a', 'ʙ': 'b', 'ᴄ': 'c', 'ᴅ': 'd', 'ᴇ': 'e', 'ꜰ': 'f', 'ɢ': 'g', 
            'ʜ': 'h', 'ɪ': 'i', 'ᴊ': 'j', 'ᴋ': 'k', 'ʟ': 'l', 'ᴍ': 'm', 'ɴ': 'n', 
            'ᴏ': 'o', 'ᴘ': 'p', 'ǫ': 'q', 'ʀ': 'r', 's': 's', 'ᴛ': 't', 'ᴜ': 'u', 
            'ᴠ': 'v', 'ᴡ': 'w', 'x': 'x', 'ʏ': 'y', 'ᴢ': 'z', 'ı': 'i', 'ʊ': 'u', 
            'ɑ': 'a', 'ɛ': 'e'
        }

        def is_cjk_or_hangul(ch):
            return (
                '\u4e00' <= ch <= '\u9fff'  # CJK
                or '\u3400' <= ch <= '\u4dbf'  # CJK 擴展
                or '\u3040' <= ch <= '\u309f'  # 日文平假名
                or '\u30a0' <= ch <= '\u30ff'  # 日文片假名
                or '\uac00' <= ch <= '\ud7af'  # Hangul 音節
                or '\u1100' <= ch <= '\u11ff'  # Hangul Jamo
            )
        text = ''.join(smallcaps_map.get(c, c) for c in text)
        normalized = []

        for ch in text:
            ch = unicodedata.normalize('NFKC', ch)
            if is_cjk_or_hangul(ch):
                normalized.append(ch)
            else:
                for c in unicodedata.normalize('NFKD', ch):
                    c = smallcaps_map.get(c, c)
                    if not unicodedata.combining(c) and unicodedata.category(c) in ('Lu', 'Ll', 'Lt', 'Lo', 'Nd'):
                        normalized.append(c)
        return ''.join(normalized)

    def load_namelist_from_excel(self, file_path, max_player):
        """
        Load Character ID and Username from ROK TOOLKIT Excel file,
        normalize usernames by removing spaces and special characters,
        and return them as a list of dictionaries.

        Args:
            file_path (str): Path to the Excel file.
            max_player (int): Maximum number of players to load.

        Returns:
            list: A list of dictionaries with keys 'ID', 'Exact Name', and 'Name'.
        """
        # Validate file existence
        if not file_path or not os.path.isfile(file_path):
            raise FileNotFoundError(f"Excel file not found: {file_path}")
        
        # Read the Excel file into a DataFrame
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            raise ValueError(f"Failed to read Excel file: {e}")


        # Ensure required columns exist
        required_columns = {"Character ID", "Username"}
        if not required_columns.issubset(df.columns):
            raise ValueError("Excel file must contain 'Character ID' and 'Username' columns")


        ids = df['Character ID']
        usernames = df['Username'].astype(str)

        name_list = [] # ID, Exact Name, Name
        count = 0
        for id_, username in zip(ids, usernames):
            if count >= max_player:
                break
            count = count + 1
            
            # Skip rows where Character ID is missing
            if pd.isna(id_):
                continue

            # Convert ID to int
            try:
                id_int = int(id_)
            except ValueError:
                print(f"Warning: Invalid Character ID '{id_}', skipped.")
                continue

            # Normalize username
            normalized_name = self._normalize_text(username)
            name_list.append({'ID': id_int,'Exact Name': username, 'Name': normalized_name})       
        
        # Save processed namelist to Excel file
        if name_list:
            try:
                file_name = "namelist_" + os.path.basename(file_path)
                excel_data = [
                    (player["ID"], player["Exact Name"], player["Name"])
                    for player in name_list
                ]
                df_out = pd.DataFrame(excel_data, columns=["ID", "Exact Name", "Name"])
                df_out.to_excel(file_name, index=False)
                print(f"Successfully saved processed namelist to {file_name}")
            except Exception as e:
                print(f"Warning: Failed to save processed namelist: {e}")
            
        return name_list
    
    def rename_file(self, folder_path):
        """
        Rename all PNG files in the specified folder to a sequential order (1.png, 2.png, ...).

        Steps:
        1. Rename all files to temporary names to avoid overwrite conflicts.
        2. Rename temporary files to final sequential names.

        Args:
            folder_path (str): Path to the folder containing PNG images.
        """
        # Check if folder exists
        if not folder_path or not os.path.isdir(folder_path):
            print(f"Error: Folder does not exist: {folder_path}")
            return
        
        try:
            files = sorted(f for f in os.listdir(folder_path) if f.lower().endswith(".png"))
            if not files:
                print(f"No PNG files found in folder: {folder_path}")
                return
            # Step 1: rename to temporary filenames
            for i, file_name in enumerate(files):
                old_path = os.path.join(folder_path, file_name)
                temp_path = os.path.join(folder_path, f"temp_{i}.png")
                os.rename(old_path, temp_path)
                
            # Step 2: rename temp files to final sequential names
            temp_files = sorted(f for f in os.listdir(folder_path) if f.startswith("temp_"))
            for i, file_name in enumerate(temp_files, start=1):
                old_path = os.path.join(folder_path, file_name)
                new_path = os.path.join(folder_path, f"{i}.png")
                os.rename(old_path, new_path)
                
            print("Name change completed successfully.")
        except Exception as e:
            print(f"Error renaming files: {e}")

    def resize_image(self, image_path, target_width=500, target_height=300):
        """
        Resize the given image to fit within target dimensions while maintaining aspect ratio.

        Args:
            image_path (str): Path to the image file.
            target_width (int): Maximum width of resized image.
            target_height (int): Maximum height of resized image.

        Returns:
            PIL.Image.Image or None: Resized image object, or None if failed.
        """
        # Check if file exists
        if not image_path or not os.path.isfile(image_path):
            print(f"Error: Image file does not exist: {image_path}")
            return None
        
        try:
            with Image.open(image_path) as image:
                img_width, img_height = image.size
                if img_width == 0 or img_height == 0:
                    print(f"Error: Image has invalid dimensions: {image_path}")
                    return None
                # Calculate resize ratio to maintain aspect ratio
                ratio = min(target_width / img_width, target_height / img_height)
                new_size = (max(1, int(img_width * ratio)), max(1, int(img_height * ratio)))  # prevent zero size
                image = image.resize(new_size, Image.LANCZOS)
                return image
            
        except Exception as e:
            print(f"Error resizing image {image_path}: {e}")
            return None
    
    def image_process(self, image, mode):
        """
        Process an image to enhance OCR readability by converting to grayscale,
        thresholding, resizing, and applying morphological operations.

        The image is divided vertically:
        - Left half (names): can be thinned via erosion to reduce noise.
        - Right half (numbers): can be thickened via dilation to enhance digits.

        Args:
            image (PIL.Image.Image): Input image to process.
            mode (int): Processing mode. 0 = alliance, 1 = mobilization.

        Returns:
            PIL.Image.Image: Processed image suitable for OCR.
        """
        if image is None:
            print("Error: input image is None")
            return None

        # Convert to RGB numpy array
        try:
            image = np.array(image.convert("RGB"))
            gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        except Exception as e:
            print(f"Error converting image to grayscale: {e}")
            return None
        
        h, w = gray.shape

        # Split image into left (names) and right (numbers) halves
        left = gray[:, :w//2]
        right = gray[:, w//2:]

        
        # Threshold
        if mode == 0:  # alliance
            _, left_bin = cv2.threshold(left, 180, 250, cv2.THRESH_BINARY)
            _, right_bin = cv2.threshold(right, 180, 250, cv2.THRESH_BINARY)
        else:  # mobilization
            _, left_bin = cv2.threshold(left, 150, 250, cv2.THRESH_BINARY_INV)
            _, right_bin = cv2.threshold(right, 150, 250, cv2.THRESH_BINARY_INV)

        left_bin = cv2.resize(left_bin, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        right_bin = cv2.resize(right_bin, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        
        # Morphological operations
        if mode == 0: # alliance
            kernel_left = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 1))
            left_bin = cv2.erode(left_bin, kernel_left, iterations=1)
            kernel_right = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
            right_bin = cv2.dilate(right_bin, kernel_right, iterations=2)
        else: # mobilization
            kernel_left = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 1))
            left_bin = cv2.erode(left_bin, kernel_left, iterations=2)
            kernel_right = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 1))
            right_bin = cv2.dilate(right_bin, kernel_right, iterations=2)
        # Combine left and right halves
        processed = np.hstack((left_bin, right_bin))
        return Image.fromarray(processed)

    def process_image_folder(self, image_folder, mode):
        """
        Process all PNG images in the given folder and save processed grayscale images
        in a subfolder called 'gray'. Existing gray folder will be deleted first.

        Args:
            image_folder (str): Path to the folder containing original images.
            mode (int): Processing mode for image_process (0=alliance, 1=mobilization).

        Returns:
            str: Path to the folder containing processed gray images.
        """
        if not image_folder or not os.path.isdir(image_folder):
            print(f"Error: Image folder does not exist: {image_folder}")
            return None
        
        gray_folder = os.path.normpath(os.path.join(image_folder, "gray"))
        
        # Delete existing gray folder if exists
        if os.path.exists(gray_folder):
            try:
                os.chmod(gray_folder, stat.S_IWRITE)
                shutil.rmtree(gray_folder)
                print("Deleted existing gray folder.")
            except Exception as e:
                print(f"Error deleting gray folder: {e}")

        # Create gray folder
        try:
            os.makedirs(gray_folder, exist_ok=True)
            print("Created gray folder.")
        except Exception as e:
            print(f"Error creating gray folder: {e}")
            return None
        
        # Process each PNG image
        for file_name in os.listdir(image_folder):
            if not file_name.lower().endswith(".png"):
                continue
            
            old_path = os.path.join(image_folder, file_name)
            new_path = os.path.join(gray_folder, f"gray{file_name}")

            try:
                with Image.open(old_path) as image:
                    processed_image = self.image_process(image, mode)
                    if processed_image is not None:
                        processed_image.save(new_path)
                    else:
                        print(f"Processing returned None for {file_name}")
            except Exception as e:
                print(f"Error processing {file_name}: {e}")
                
        return gray_folder
    
    def _match_with_namelist(self, name_list, threshold):
        """
        Attempt to match unmatched OCR names with the provided name list using fuzzy matching.
        Updates match_dict and match_flag for successful matches, and maintains unmatched_data
        for names that cannot be matched above the threshold.

        Args:
            name_list (list): List of dictionaries containing 'ID', 'Name', and 'Exact Name'.
            threshold (int): Minimum fuzzy match score to consider a valid match.

        Behavior:
            - Filters out already matched IDs from name_list.
            - For each unmatched OCR entry:
                - Uses fuzzy matching (process.extract) against remaining name_list.
                - Filters out candidates with length difference > 2.
                - If best match score >= threshold, adds to match_dict and match_flag.
                - Otherwise, keeps entry in unmatched_data.
        """
        # Filter out already matched IDs
        name_list_unmatched = [item for item in name_list if item['ID'] not in self.match_flag]

        unmatched_data2 = []
        
        for unmatched in self.unmatched_data:
            name, number, image_cnt = unmatched
            
            # Fuzzy match: top 5 candidates
            candidates = [n['Name'] for n in name_list_unmatched]
            if not candidates:
                unmatched_data2.append(unmatched)
                continue
            matches = process.extract(name, candidates, limit=5)

            # Filter candidates with large length difference
            matches = [(m, s) for m, s in matches if abs(len(m) - len(name)) <= 3]
            if not matches:
                unmatched_data2.append(unmatched)
                continue

            best_match, score = matches[0]
            
            if score >= threshold:
                # Find the matched entry in name_list_unmatched
                matched_entry = next(
                    (entry for entry in name_list_unmatched if entry['Name'] == best_match),
                    None,
                )
                if matched_entry:
                    matched_id = matched_entry['ID']
                    exact_name = matched_entry['Name']
                    
                    if matched_id in self.match_flag: # duplicate matched ID
                        unmatched_data2.append((name, number, image_cnt))
                    else:
                        self.match_dict[matched_id] = (best_match, name, exact_name, number, image_cnt)
                        self.match_flag.add(matched_id)
            else:
                unmatched_data2.append((name, number, image_cnt))
                
        # Update unmatched_data with entries that were not successfully matched
        self.unmatched_data = unmatched_data2[:]
    
    def extract_info_from_image(self, image_cnt, image_path):
        """
        Perform OCR on a given image and extract names with their corresponding numbers.

        Args:
            image_cnt (int): Index of the current image.
            image_path (str): File path to the image.

        Behavior:
            - Opens the image using PIL.
            - Uses pytesseract to extract text in multiple languages.
            - Cleans up lines (removes extra whitespace).
            - Extracts "name" and "number" pairs using regex.
            - Normalizes the name to remove special characters.
            - Appends successfully extracted data to self.unmatched_data.
            - Appends lines that fail to match the pattern to self.fail_extract_data.
        """
        try:
            with Image.open(image_path) as image:
                # OCR read text
                custom_config = r'--psm 6'
                text = pytesseract.image_to_string(
                    image, lang="eng+chi_tra+chi_sim+jpn+kor", config=custom_config
                )
        except Exception as e:
            print(f"Error reading image {image_path}: {e}")
            return
        
        # Split text into lines and clean whitespace
        lines = text.split("\n")
        cleaned_lines = [re.sub(r"\s+", " ", line).strip() for line in lines if line.strip()]  
        
        for line in cleaned_lines:
            # Match "name + number" at the end of line
            match = re.match(r"(.+?)\s+([\d,\.]+)$", line, re.UNICODE)
            if match:
                raw_name = match.group(1).strip()
                norm_name = self._normalize_text(raw_name)
                try:
                    number = int(match.group(2).replace(".", "").replace(",", ""))
                    self.unmatched_data.append((norm_name, number, image_cnt))
                except ValueError:
                    continue
            else:
                self.fail_extract_data.append((line, None, image_cnt))                
    
    def match_data(self, image_folder, name_list, progress_callback=None, finish_callback=None):
        """
        Perform OCR extraction and match extracted names to a provided name list.
        
        Args:
            image_folder (str): Path to the folder containing images to process.
            name_list (list): List of dictionaries containing 'ID', 'Name', and 'Exact Name'.
            progress_callback (callable, optional): Function called with progress value (0~1) to update UI.
            finish_callback (callable, optional): Function called when matching is complete with result data.

        Behavior:
            1. Iterates over all PNG images in the folder, sorted numerically by filename.
            2. Extracts name/number pairs from each image and stores them in `self.unmatched_data`.
            3. Performs multi-threshold matching (_match_with_namelist) to assign unmatched names to the official name list.
            4. Appends matched results into `self.data` corresponding to each image index.
            5. Appends unmatched OCR data with placeholder values.
            6. Reports failed OCR extractions via `self.fail_extract_data`.
            7. Calls progress_callback at key steps and finish_callback when complete.
        """
        # Initialize counters
        image_cnt = 0
        idx = 0
        
        # Ensure folder exists
        if not os.path.exists(image_folder):
            print(f"Image folder '{image_folder}' does not exist.")
            return
        
        # Count steps for progress bar (image count + 3 match phases)
        max_step = len(os.listdir(image_folder)) + 3
        
        # Sort images numerically by any number in the filename
        image_files = sorted(
            [f for f in os.listdir(image_folder) if f.endswith(".png")],
            key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0
        )
        
        # OCR extraction phase
        for file_name in image_files:
            if file_name.endswith(".png"):
                self.data.append([])
                image_path = os.path.join(image_folder, file_name)
                self.extract_info_from_image(image_cnt, image_path)
                image_cnt = image_cnt + 1
                
                # Update progress
                if progress_callback:
                    progress_callback((idx + 1) / max_step)
                idx = idx + 1 
            
        # Matching phase with different thresholds
        for threshold in [90, 70, 60]:
            self._match_with_namelist(name_list, threshold)
            if progress_callback:
                progress_callback((idx + 1) / max_step)
            idx += 1
        print("Successfully complete match.")

        # Append matched data to self.data
        for matched_id, (best_match, ocr_name, exact_name, number, image_cnt,) in self.match_dict.items():
            self.data[image_cnt].append({
                    "ID": matched_id,
                    "OCR Name": ocr_name,
                    "Match Name": best_match,
                    "Exact Name": exact_name,
                    "Number": number,
                    "Image Count": image_cnt,
            })
            
        # Append unmatched OCR data to self.data
        for item in self.unmatched_data:
            number = item[1] if item[1] else 0
            self.data[item[2]].append({
                    "ID": 0,
                    "OCR Name": item[0],
                    "Match Name": "unmatched",
                    "Exact Name": "unmatched",
                    "Number": number,
                    "Image Count": item[2],
            })
            
        print("Fail to extrace from image:")
        print(self.fail_extract_data)
        
        
        # Call finish callback with all result data
        if finish_callback:
            finish_callback(
                self.data, self.unmatched_data, self.match_dict, self.match_flag
            )

    def save_excel(self, file_name, data):
        """
        Save the matched data to an Excel file.

        Args:
            file_name (str): Path to save the Excel file.
            data (list of tuples): Each tuple should contain 
                (ID, OCR Name, Match Name, Exact Name, Number, Diff).

        Returns:
            None
        """
        # Validate input
        if not data:
            print("No data to save.")
            return
        if not file_name or not file_name.endswith(".xlsx"):
            print("Invalid file name. Must end with '.xlsx'.")
            return
        
        try:
            df = pd.DataFrame(
                data, columns=["ID", "OCR Name", "Match Name", "Exact Name", "Number", "Diff"]
            )
            df.to_excel(file_name, index=False)
            print("Successfully saved to excel file.")
        except Exception as e:
            print(f"Error saving Excel file: {e}")
