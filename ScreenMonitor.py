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

class ScreenMonitor:
    def __init__(self, template_dir, input_dir, interval=5, cleanup_days=7):
        self.template_dir = template_dir
        self.input_dir = input_dir
        self.interval = interval
        self.cleanup_days = cleanup_days
        self.ensure_input_directory()

    def ensure_input_directory(self):
        if not os.path.exists(self.input_dir):
            os.makedirs(self.input_dir)

    def capture_and_process_screen(self):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot = pyautogui.screenshot()
        img = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        return img, timestamp

    def compare_with_template(self, screen):
        for template_file in os.listdir(self.template_dir):
            template_path = os.path.join(self.template_dir, template_file)
            template = cv2.imread(template_path, 0)
            screen_gray = cv2.cvtColor(screen, cv2.COLOR_BGR2GRAY)
            
            res = cv2.matchTemplate(screen_gray, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
            
            if max_val > 0.8:
                print(f"Tìm thấy kết quả phù hợp với template: {template_file}")
                return max_loc, template.shape[::-1]  # Trả về vị trí và kích thước của template
        print("Không tìm thấy kết quả phù hợp với bất kỳ template nào")
        return None, None
    def click_ok(self, location):
        x, y = location
        click_x, click_y = x + 50, y + 50

        # Chụp ảnh nhỏ 80x80 px với tâm là điểm click
        screenshot = pyautogui.screenshot(region=(click_x - 40, click_y - 40, 80, 80))
        img = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

        # Vẽ đường tròn màu đỏ
        cv2.circle(img, (40, 40), 10, (0, 0, 255), 2)

        # Lưu ảnh
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"click_{timestamp}.png"
        filepath = os.path.join(self.input_dir, filename)
        cv2.imwrite(filepath, img)

        # Thực hiện click
        pyautogui.click(click_x, click_y)
        print(f"Đã click tại vị trí: ({click_x}, {click_y})")
        print(f"Đã lưu ảnh click: {filepath}")

   
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

    def draw_and_save_result(self, img, location, size, timestamp, found):
        click_pos = None
        if location and size:
            # Vẽ viền đỏ xung quanh đối tượng được tìm thấy
            top_left = location
            bottom_right = (top_left[0] + size[0], top_left[1] + size[1])
            cv2.rectangle(img, top_left, bottom_right, (0, 0, 255), 2)

            #---------------TEST----------------
            # Tính toán vị trí điểm mới
            new_point_x = bottom_right[0] - 70  # Cách cạnh phải 60px
            new_point_y = top_left[1] + 70      # Cách cạnh trên 60px
            new_point = (new_point_x, new_point_y)

            # Vẽ điểm mới (ví dụ: một vòng tròn màu xanh lá)
            cv2.circle(img, new_point, 5, (0, 255, 0), -1)
            #---------------END TEST----------------

            # Tính toán vị trí trung tâm để click
            click_x, click_y = top_left[0] + size[0] // 2, top_left[1] + size[1] // 2

            # Vẽ hình vuông 80x80 px với nét đứt màu vàng
            top_left_square = (click_x - 40, click_y - 40)
            bottom_right_square = (click_x + 40, click_y + 40)
            self.draw_dashed_rectangle(img, top_left_square, bottom_right_square, (0, 255, 255), 2)

            # Vẽ điểm click
            cv2.circle(img, (click_x, click_y), 5, (255, 0, 0), -1)
            
            click_pos = (click_x, click_y)

        # Lưu ảnh
        prefix = "OK_" if found else "NG_"
        filename = f"{prefix}screen_{timestamp}.png"
        filepath = os.path.join(self.input_dir, filename)
        cv2.imwrite(filepath, img)
        print(f"Đã lưu ảnh kết quả: {filepath}")

        return click_pos
    
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
            screen, timestamp = self.capture_and_process_screen()
            location, size = self.compare_with_template(screen)
            click_pos = self.draw_and_save_result(screen, location, size, timestamp, location is not None)
            if click_pos:
                pyautogui.click(click_pos)
                print(f"Đã click tại vị trí: {click_pos}")
            time.sleep(self.interval)


    