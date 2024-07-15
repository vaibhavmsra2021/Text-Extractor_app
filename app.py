import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import ImageGrab, Image
import pyautogui
import time
import json
import os
import numpy as np
import cv2
import pygetwindow as gw
import threading
import subprocess
import easyocr
import pandas as pd
import pickle
import sys


reader = easyocr.Reader(['en'])

class App:
    def __init__(self, root, reader):
        self.root = root
        self.root.title("Screen Pathway Selector")
        self.selected_images = []
        self.pathway_file = 'button_pathway.json'
        self.preferences_file = 'preferences.json'
        self.rois_file = 'rois.pkl'
        self.reader = reader
        self.parameters = []
        self.delay_time = None
        self.target_application_path = None
        self.target_application_name = None
        self.running_event = threading.Event()


        
        if self.load_preferences():
            self.initialize_app()
        else:
            self.show_preferences_window()

    def load_preferences(self):
        if os.path.exists(self.preferences_file):
            with open(self.preferences_file, 'r') as f:
                preferences = json.load(f)
                self.parameters = preferences['parameters']
                self.delay_time = preferences['delay_time']
                self.target_application_name = preferences['target_application_name']
                self.target_application_path = preferences['target_application_path']
                return True
        return False

    def save_preferences(self):
        preferences = {
            'parameters': self.parameters,
            'delay_time': self.delay_time,
            'target_application_name': self.target_application_name,
            'target_application_path': self.target_application_path
            
        }
        with open(self.preferences_file, 'w') as f:
            json.dump(preferences, f)

    def show_preferences_window(self):
        self.pref_window = tk.Toplevel(self.root)
        self.pref_window.title("Set Preferences")

        tk.Label(self.pref_window, text="Parameters (comma separated):").grid(row=0, column=0, padx=10, pady=10)
        self.parameters_entry = tk.Entry(self.pref_window)
        self.parameters_entry.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(self.pref_window, text="Delay Time (seconds):").grid(row=1, column=0, padx=10, pady=10)
        self.delay_entry = tk.Entry(self.pref_window)
        self.delay_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(self.pref_window, text="Target Application Name:").grid(row=2, column=0, padx=10, pady=10)
        self.name_entry = tk.Entry(self.pref_window)
        self.name_entry.grid(row=2, column=1, padx=10, pady=10)

        tk.Label(self.pref_window, text="Target Application Path:").grid(row=3, column=0, padx=10, pady=10)
        self.path_entry = tk.Entry(self.pref_window)
        self.path_entry.grid(row=3, column=1, padx=10, pady=10)

        browse_button = tk.Button(self.pref_window, text="Browse", command=self.browse_file)
        browse_button.grid(row=3, column=2, padx=10, pady=10)

        save_button = tk.Button(self.pref_window, text="Save", command=self.save_pref_and_initialize)
        save_button.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.path_entry.insert(0, file_path)

    def save_pref_and_initialize(self):
        self.parameters = self.parameters_entry.get().split(',')
        self.delay_time = int(self.delay_entry.get())
        self.target_application_path = self.path_entry.get()
        self.target_application_name = self.name_entry.get()
        self.save_preferences()
        self.pref_window.destroy()
        self.initialize_app()

    def initialize_app(self):
        if not self.check_application_open():
            self.open_application()

        if os.path.exists(self.pathway_file):
            self.load_pathway()
            if self.selected_images:
                self.running_event.set()
                threading.Thread(target=self.repeat_execution, daemon=True).start()

        else:
            newindow = gw.getWindowsWithTitle(self.target_application_name)
            newindow = newindow[0]
            newindow.activate()
            capture_button = tk.Button(self.root, text="Capture Button Area", command=self.capture_area)
            capture_button.pack(pady=20)
            save_pathway_button = tk.Button(self.root, text="Save Pathway", command=self.save_pathway)
            save_pathway_button.pack(pady=20)
            messagebox.showinfo("Button Pathway", "Select button pathways by clicking capture area button and save it by pressing the save pathway button")
        
        exit_button = tk.Button(self.root, text="Pause", command=self.pause_program)
        exit_button.pack(pady=20)

        resume_button = tk.Button(self.root, text="Resume", command=self.resume_program)
        resume_button.pack(pady=20)
        
        download_button = tk.Button(self.root, text="Download OCR Data", command=self.download_ocr_data)
        download_button.pack(pady=20)

        reset_button = tk.Button(self.root, text="Reset", command=self.reset_data)
        reset_button.pack(pady=20)

    def check_application_open(self):
        windows = gw.getWindowsWithTitle(self.target_application_name)
        return bool(windows)

    def open_application(self):
        try:
            subprocess.Popen([self.target_application_path])
            time.sleep(5)
        except Exception as e:
            print(f"Error opening Targetted Application: {e}")

    def capture_area(self):
        self.root.withdraw()
        time.sleep(0.5)
        screen = pyautogui.screenshot()
        selection_window = tk.Toplevel(self.root)
        selection_window.attributes('-fullscreen', True)
        selection_window.attributes('-alpha', 0.3)
        selection_window.attributes('-topmost', True)
        canvas = tk.Canvas(selection_window, cursor="cross", bg='grey11')
        canvas.pack(fill=tk.BOTH, expand=True)
        rect = None
        start_x = start_y = curX = curY = 0

        def on_button_press(event):
            nonlocal start_x, start_y, rect
            start_x = canvas.canvasx(event.x)
            start_y = canvas.canvasy(event.y)
            rect = canvas.create_rectangle(start_x, start_y, start_x, start_y, outline='red', width=2)

        def on_move_press(event):
            nonlocal curX, curY
            curX = canvas.canvasx(event.x)
            curY = canvas.canvasy(event.y)
            canvas.coords(rect, start_x, start_y, curX, curY)

        def on_button_release(event):
            nonlocal rect
            x1, y1, x2, y2 = (int(start_x), int(start_y), int(curX), int(curY))
            if x1 != x2 and y1 != y2:
                selected_image = screen.crop((x1, y1, x2, y2))
                self.selected_images.append(selected_image)
                selection_window.destroy()
                self.root.deiconify()
            else:
                print("Selection cancelled.")
                selection_window.destroy()
                self.root.deiconify()

        canvas.bind("<ButtonPress-1>", on_button_press)
        canvas.bind("<B1-Motion>", on_move_press)
        canvas.bind("<ButtonRelease-1>", on_button_release)
        selection_window.mainloop()

    def save_pathway(self):
        if not self.selected_images:
            print("No images captured.")
            return

        pathway = []
        for image in self.selected_images:
            image_data = np.array(image)
            encoded_image = cv2.imencode('.png', image_data)[1].tobytes()
            pathway.append(encoded_image.hex())

        with open(self.pathway_file, 'w') as f:
            json.dump(pathway, f)
        print("Pathway saved successfully.")

        time.sleep(2)
        self.running_event.set()
        threading.Thread(target=self.repeat_execution, daemon=True).start()

    def load_pathway(self):
        with open(self.pathway_file, 'r') as f:
            pathway = json.load(f)

        self.selected_images = []
        for hex_image in pathway:
            image_data = bytes.fromhex(hex_image)
            image = cv2.imdecode(np.frombuffer(image_data, np.uint8), cv2.IMREAD_COLOR)
            pil_image = Image.fromarray(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
            self.selected_images.append(pil_image)
        print("Pathway loaded successfully.")

    def execute_pathway(self):
        if not self.selected_images:
            print("No pathway loaded.")
            return

        original_window = gw.getActiveWindow()
        target_window = None
        windows = gw.getWindowsWithTitle(self.target_application_name)
        if windows:
            target_window = windows[0]

        if not target_window:
            self.open_application()
            windows = gw.getWindowsWithTitle(self.target_application_name)
            if windows:
                target_window = windows[0]
            else:
                print("Failed to open Targeted Application.")
                return

        try:
            target_window.activate()
            time.sleep(2)
            all_buttons_clicked = True
            for image in self.selected_images:
                open_cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
                try:
                    button_location = pyautogui.locateCenterOnScreen(open_cv_image, confidence=0.8)
                    if button_location is not None:
                        pyautogui.moveTo(button_location)
                        pyautogui.click()
                        print("Button clicked successfully.")
                    else:
                        print("Button not found on the screen.")
                        all_buttons_clicked = False
                        break
                    time.sleep(2)
                except pyautogui.ImageNotFoundException:
                    print("Image not found exception occurred.")
                    all_buttons_clicked = False
                    break
                except Exception as e:
                    print(f"An unexpected error occurred: {e}")
                    all_buttons_clicked = False
                    break

            if all_buttons_clicked:
                self.run_ocr_and_save()
            else:
                print("Not all buttons were clicked successfully. OCR operation aborted.")

            if original_window:
                original_window.activate()
                print("Switched back to the original window.")

        except Exception as e:
            print(f"Error activating window: {e}")

    def run_ocr_and_save(self):
        def capture_active_window():
            active_window = gw.getActiveWindow()
            if active_window is not None:
                bbox = (active_window.left, active_window.top, active_window.right, active_window.bottom)
                screenshot = np.array(ImageGrab.grab(bbox))
                screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)
                return screenshot
            else:
                raise Exception("No active window detected")

        image = capture_active_window()
        clone = image.copy()

        def preprocess_image(image):
            
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

            # Resize the image to ensure it's suitable for OCR (scaling by a factor of 2)
            scale_factor = 2
            width = int(gray.shape[1] * scale_factor)
            height = int(gray.shape[0] * scale_factor)
            resized = cv2.resize(gray, (width, height), interpolation=cv2.INTER_CUBIC)

            # # Apply histogram equalization to improve the contrast
            # equalized = cv2.equalizeHist(resized)

            # Apply a Gaussian blur to the image
            blurred = cv2.GaussianBlur(resized, (3, 3), 0)

            # Apply adaptive thresholding
            adaptive_thresh = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                                    cv2.THRESH_BINARY, 11, 2)

            # Use morphological operations to remove noise
            kernel = np.ones((2, 2), np.uint8)
            morph = cv2.morphologyEx(adaptive_thresh, cv2.MORPH_CLOSE, kernel)
            return morph

        try:
            with open(self.rois_file, 'rb') as f:
                rois = pickle.load(f)
            print("Loaded saved ROIs.")
        except FileNotFoundError:
            rois = {param: None for param in self.parameters}

        def select_rois(event, x, y, flags, param):
            nonlocal clone, current_param, rois
            if event == cv2.EVENT_LBUTTONDOWN:
                rois[current_param] = [(x, y)]
            elif event == cv2.EVENT_LBUTTONUP:
                rois[current_param].append((x, y))
                cv2.rectangle(clone, rois[current_param][0], rois[current_param][1], (0, 255, 0), 2)
                cv2.imshow("image", clone)
                cv2.waitKey(500)

        def ocr_region(image, roi):
            (x1, y1), (x2, y2) = roi
            region = image[y1:y2, x1:x2]
            preprocessed_region = preprocess_image(region)
            result = self.reader.readtext(preprocessed_region)
            text = ' '.join([res[1] for res in result])
            return text.strip()

        if None in rois.values():
            messagebox.showinfo("ROIs Selection", "Select the ROIs for your parameters")
            for param in self.parameters:
                current_param = param
                print(f"Select the region for {param}.")
                clone = image.copy()
                cv2.imshow("image", clone)
                cv2.setMouseCallback("image", select_rois)
                cv2.waitKey(0)

            cv2.destroyAllWindows()   

            with open(self.rois_file, 'wb') as f:
                pickle.dump(rois, f)
            print("ROIs saved.")

        extracted_values = {param: ocr_region(image, rois[param]) for param in self.parameters}

        current_date = pd.to_datetime('now').strftime('%Y-%m-%d')
        current_time = pd.to_datetime('now').strftime('%H:%M:%S')

        data = {"Date": [current_date], "Time": [current_time]}
        for param in self.parameters:
            data[param] = [extracted_values[param]]

        print(data)
        df_new = pd.DataFrame(data)
        filename = f'OCR_{current_date}.xlsx'

        if os.path.exists(filename):
            df_existing = pd.read_excel(filename)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_combined = df_new

        df_combined.to_excel(filename, index=False)
        print(f"Data saved to {filename}")

    def repeat_execution(self):
        while self.running_event.is_set():
            self.execute_pathway()
            time.sleep(self.delay_time)
            

    def pause_program(self):
        self.running_event.clear()
        messagebox.showinfo("Program Stopped", "Iteration stopped. You can now download OCR data.")

    def resume_program(self):
        if not self.running_event.is_set():
            self.running_event.set()
            threading.Thread(target=self.repeat_execution, daemon=True).start()
            messagebox.showinfo("Program Resumed", "OCR operations resumed")
            print("Program resumed.")



    def download_ocr_data(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            current_date = pd.to_datetime('now').strftime('%Y-%m-%d')
            filename = f'OCR_{current_date}.xlsx'
            if os.path.exists(filename):
                os.rename(filename, file_path)
                messagebox.showinfo("Download Complete", f"File saved as {file_path}")
            else:
                messagebox.showwarning("File Not Found", "No OCR data file found to download.")


    def reset_data(self):
        self.running_event.clear()
        if os.path.exists(self.pathway_file):
            os.remove(self.pathway_file)
        if os.path.exists(self.preferences_file):
            os.remove(self.preferences_file)
        if os.path.exists(self.rois_file):
            os.remove(self.rois_file)
        messagebox.showinfo("Reset", "Preferences, pathways, and ROIs have been reset. Download the excel file of the saved data")
        print("Files reset successfully")
        self.download_ocr_data()

        self.root.destroy()
        main()
        

def main():
    root = tk.Tk()
    app = App(root, reader)
    root.mainloop()

if __name__ == '__main__':
    main()
