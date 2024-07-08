import time
import pyautogui
import cv2
import numpy as np
import os
import psutil
import win32gui
import win32process
import csv
from datetime import datetime, timedelta
import shutil

class PopupChecker:
    def __init__(self, app_name, template_dir, csv_file, input_dir, cleanup_days=7):
        self.app_name = app_name
        self.template_dir = template_dir
        self.csv_file = csv_file
        self.input_dir = input_dir
        self.cleanup_days = cleanup_days
        self.ensure_input_directory()

    def ensure_input_directory(self):
        if not os.path.exists(self.input_dir):
            os.makedirs(self.input_dir)

    def get_app_windows(self):
        """
        Lấy danh sách các handle của cửa sổ thuộc về ứng dụng đang chạy với tên app_name.
        """
        hwnds = []
        def callback(hwnd, hwnds):
            if self.is_app_window(hwnd):
                hwnds.append(hwnd)
            return True

        win32gui.EnumWindows(callback, hwnds)
        return hwnds

    def is_app_window(self, hwnd):
        """
        Kiểm tra xem cửa sổ có hiển thị, được kích hoạt và thuộc về tiến trình của ứng dụng cần kiểm tra hay không.
        """
        if not win32gui.IsWindowVisible(hwnd) or not win32gui.IsWindowEnabled(hwnd):
            return False

        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            return os.path.basename(process.exe()) == self.app_name
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            return False

    def capture_popup(self, hwnd):
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"popup_{timestamp}.png"
        filepath = os.path.join(self.input_dir, filename)
        screenshot = pyautogui.screenshot(region=(left, top, right-left, bottom-top))
        screenshot.save(filepath)
        print(f"Đã lưu ảnh chụp popup: {filepath}")
        return screenshot, filepath, (left, top)

    def read_csv(self):
        try:
            with open(self.csv_file, 'r', encoding='utf-8') as file:
                reader = csv.reader(file)
                return [row for row in reader]
        except UnicodeDecodeError as e:
            print(f"Lỗi UnicodeDecodeError: {e}")
            return None

    def compare_with_csv(self, caption, button):
        data = self.read_csv()
        for row in data:
            if row[1] in caption:
            # if caption == row['Caption'] and button == row['Button']:
                print(f"Nội dung popup phù hợp với dữ liệu trong CSV: {row}")
                # return row['Action']
                return row[4]
        print("Nội dung popup không phù hợp với bất kỳ dữ liệu nào trong CSV")
        return None

    def click_button(self, hwnd, button_name):
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        button_x, button_y = self.find_button_coordinates(hwnd, button_name)

        if button_x is not None and button_y is not None:
            pyautogui.click(left + button_x, top + button_y)
            print(f"Đã click {button_name} trên popup tại vị trí: ({left + button_x}, {top + button_y})")
            return left + button_x, top + button_y
        else:
            print(f"Không tìm thấy nút {button_name}")
            return None

    def find_button_coordinates(self, hwnd, button_name):
        # Logic tìm tọa độ của nút theo tên nút
        # Giả định nút OK nằm ở vị trí cố định so với cửa sổ
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        if button_name == "OK":
            return (right - left) // 2, bottom - top - 30
        return None, None

    def draw_dashed_line(self, img, start, end, color, thickness=1, dash_length=10, gap_length=5):
        dist = np.linalg.norm(np.array(end) - np.array(start))
        dashes = int(dist / (dash_length + gap_length))
        for i in range(dashes):
            s = i * (dash_length + gap_length)
            e = s + dash_length
            if e > dist:
                e = dist
            start_point = tuple(map(int, start + (end - start) * s / dist))
            end_point = tuple(map(int, start + (end - start) * e / dist))
            cv2.line(img, start_point, end_point, color, thickness)

    def draw_dashed_rectangle(self, img, top_left, bottom_right, color, thickness=1, dash_length=10, gap_length=5):
        x1, y1 = top_left
        x2, y2 = bottom_right
        
        self.draw_dashed_line(img, np.array([x1, y1]), np.array([x2, y1]), color, thickness, dash_length, gap_length)  # Top
        self.draw_dashed_line(img, np.array([x2, y1]), np.array([x2, y2]), color, thickness, dash_length, gap_length)  # Right
        self.draw_dashed_line(img, np.array([x2, y2]), np.array([x1, y2]), color, thickness, dash_length, gap_length)  # Bottom
        self.draw_dashed_line(img, np.array([x1, y2]), np.array([x1, y1]), color, thickness, dash_length, gap_length)  # Left

    def draw_and_save_result(self, img_path, hwnd, button_pos, is_ok):
        img = cv2.imread(img_path)
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        top_left = (0, 0)
        bottom_right = (right - left, bottom - top)

        # Vẽ viền đỏ quanh cửa sổ
        cv2.rectangle(img, top_left, bottom_right, (0, 0, 255), 2)

        # Vẽ viền nét đứt quanh nút OK
        if button_pos:
            top_left_square = (button_pos[0] - 40, button_pos[1] - 40)
            bottom_right_square = (button_pos[0] + 40, button_pos[1] + 40)
            self.draw_dashed_rectangle(img, top_left_square, bottom_right_square, (0, 255, 255), 2)

        prefix = "OK_" if is_ok else "NG_"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{prefix}popup_{timestamp}.png"
        filepath = os.path.join(self.input_dir, filename)
        cv2.imwrite(filepath, img)
        print(f"Đã lưu ảnh kết quả popup: {filepath}")

    def cleanup_input_directory(self):
        now = datetime.now()
        for filename in os.listdir(self.input_dir):
            file_path = os.path.join(self.input_dir, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                    if now - file_time > timedelta(days=self.cleanup_days):
                        os.unlink(file_path)
                        print(f"Đã xóa file: {file_path}")
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
                    print(f"Đã xóa thư mục: {file_path}")
            except Exception as e:
                print(f"Lỗi khi xóa {file_path}. Lý do: {e}")

    def run(self):
        while True:
            self.cleanup_input_directory()
            for hwnd in self.get_app_windows():
                screenshot, filepath, window_pos = self.capture_popup(hwnd)
                caption = win32gui.GetWindowText(hwnd)
                action = self.compare_with_csv(caption, "OK")
                if action == "Click":
                    button_pos = self.click_button(hwnd, "OK")
                    is_ok = button_pos is not None
                else:
                    button_pos = None
                    is_ok = False
                self.draw_and_save_result(filepath, hwnd, button_pos, is_ok)
            time.sleep(1)

# # Sử dụng
# checker = PopupChecker(app_name='YourAppName.exe', template_dir='templates', csv_file='data.csv', input_dir='screenshots')
# checker.run()
