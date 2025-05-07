# falconCommand.py
#
import argparse
import datetime
import io
import os
import cv2
import numpy as np
import sys
import time
from pathlib import Path
import pyperclip
import pyautogui
from PIL import Image

# 確保控制台輸出使用 UTF-8
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

COMMAND_VERSION = "1.0.3"  # Add version number here


class AutoGUIController:
    def __init__(self):
        # Set pyautogui's security settings
        pyautogui.FAILSAFE = True
        self.version = COMMAND_VERSION
        self.parser = self._create_parser()

    def _create_parser(self):
        parser = argparse.ArgumentParser(
            description=f"Falcon UI Automation Tool v{self.version}",  # Show version in help
            formatter_class=argparse.RawDescriptionHelpFormatter,
        )
        # Add version argument
        parser.add_argument(
            "--version", action="version", version=f"falconCommand v{self.version}"
        )
        # Modify various click command parameters to make them optional
        parser.add_argument(
            "--click",
            nargs="*",
            type=int,
            metavar="X Y",
            help="Click at the specified coordinates (x, y) or at current position if no coordinates provided",
        )
        parser.add_argument(
            "--fast-click",
            nargs="*",
            type=float,
            metavar="X Y COUNT DELAY",
            help="Fast click at coordinates (X, Y) or current position for COUNT times with minimal DELAY",
        )
        parser.add_argument(
            "--double-click",
            nargs="*",
            type=int,
            metavar="X Y",
            help="Double click at the specified coordinates (x, y) or at current position if no coordinates provided",
        )
        parser.add_argument(
            "--right-click",
            nargs="*",
            type=int,
            metavar="X Y",
            help="Right click at the specified coordinates (x, y) or at current position if no coordinates provided",
        )
        parser.add_argument(
            "--moveto",
            nargs="+",
            type=float,
            metavar="N",
            help="Move mouse to the specified coordinates (X, Y) over an optional DURATION seconds",
        )
        parser.add_argument(
            "--scroll",
            nargs="+",
            type=int,
            metavar="N",
            help="Scroll N clicks (positive for up, negative for down) at optional X Y coordinates",
        )

        parser.add_argument(
            "--type", type=str, metavar="TEXT", help="Type the specified text"
        )
        parser.add_argument(
            "--press",
            type=str,
            metavar="KEY",
            help="Press a key or key combination (e.g., enter, tab, ctrl+c)",
        )
        parser.add_argument(
            "--clipboard-copy",
            action="store_true",
            help="Copy current selection to clipboard",
        )
        parser.add_argument(
            "--clipboard-paste", action="store_true", help="Paste from clipboard"
        )
        parser.add_argument(
            "--clipboard-set", type=str, metavar="TEXT", help="Set clipboard content"
        )
        parser.add_argument(
            "--clipboard-get", action="store_true", help="Get clipboard content"
        )

        parser.add_argument(
            "--screenshot",
            type=str,
            metavar="FILENAME",
            help="Take a screenshot and save it to the specified filename",
        )
        parser.add_argument(
            "--position", action="store_true", help="Get current mouse position"
        )
        parser.add_argument(
            "--delay",
            type=float,
            default=0.1,
            help="Delay between actions in seconds (default: 0.1)",
        )
        parser.add_argument(
            "--timeout",
            type=float,
            default=None,
            help="Wait timeout in seconds (default: None)",
        )
        parser.add_argument(
            "--wait-time",
            type=float,
            default=None,
            help="Alternative wait timeout in seconds for wait-until functions (default: None)",
        )
        parser.add_argument(
            "--screen-size", action="store_true", help="Get screen size"
        )
        parser.add_argument(
            "--window-info",
            type=str,
            metavar="TITLE",
            help="Get position and size of a window with the specified title",
        )
        parser.add_argument(
            "--track-position",
            type=float,
            metavar="DURATION",
            help="Track and display mouse position for specified duration in seconds",
        )
        parser.add_argument(
            "--relative-move",
            nargs=2,
            type=int,
            metavar=("X", "Y"),
            help="Move mouse relative to current position",
        )
        parser.add_argument(
            "--center-on-screen",
            action="store_true",
            help="Move mouse to the center of the screen",
        )
        parser.add_argument(
            "--position-to-clipboard",
            action="store_true",
            help="Copy current mouse position to clipboard",
        )

        parser.add_argument(
            "--drag-to",
            nargs="+",
            type=float,
            metavar="N",
            help="Drag from current position to specified coordinates (X, Y) over optional DURATION seconds",
        )
        parser.add_argument(
            "--search-image",
            type=str,
            metavar="IMAGE_PATH",
            help="Search the specified image on screen",
        )
        parser.add_argument(
            "--click-image",
            type=str,
            metavar="IMAGE_PATH",
            help="Click the center of the specified image on screen",
        )
        parser.add_argument(
            "--right-click-image",
            type=str,
            metavar="IMAGE_PATH",
            help="Right click the center of the specified image on screen",
        )
        parser.add_argument(
            "--double-click-image",
            type=str,
            metavar="IMAGE_PATH",
            help="Locate and double-click the center of the specified image on screen",
        )
        parser.add_argument(
            "--image-exists",
            type=str,
            metavar="IMAGE_PATH",
            help="Check if an image exists on screen",
        )
        parser.add_argument(
            "--command-file",
            type=str,
            metavar="FILE_PATH",
            help="Execute commands from a text file",
        )
        parser.add_argument(
            "--launch", type=str, metavar="EXE_PATH", help="Launch a Application"
        )
        parser.add_argument(
            "--check-software",
            type=str,
            metavar="SOFTWARE_NAME",
            help="Check if specified software is installed on this computer",
        )
        parser.add_argument(
            "--run",
            nargs=argparse.REMAINDER,
            metavar="COMMAND",
            help="Run a specific command or sequence followed by optional repeat count (e.g., --run click 100 200 --repeat 5)",
        )
        parser.add_argument(
            "--sleep",
            type=float,
            metavar="SECONDS",
            help="Pause execution for specified number of seconds",
        )
        parser.add_argument("--repeat", type=int, help=argparse.SUPPRESS)
        parser.add_argument(
            "--stop-on-error",
            action="store_true",
            help="Stop script execution when any command returns an error",
        )
        parser.add_argument(
            "--wait-until-exist", type=str, help="Wait until file/image/folder exists"
        )
        parser.add_argument(
            "--wait-until-process",
            type=str,
            help="Wait until specific process is running",
        )
        parser.add_argument(
            "--wait-until-installed",
            type=str,
            metavar="SOFTWARE_NAME",
            help="Wait until specified software is installed on this computer"
        )
        parser.add_argument(
            "--check-interval",
            type=float,
            default=1.0,
            help="Interval in seconds between checks",
        )

        return parser

    def save_log_to_file(self, log_output, script_path=None):
        try:
            # Create Falcon_Log directory if it doesn't exist

            today = datetime.datetime.now().strftime("%Y-%m-%d")
            log_dir = Path(f"C:/Falcon_Log/{today}")
            log_dir.mkdir(parents=True, exist_ok=True)
            # log_dir = Path("C:/Falcon_Log")
            # log_dir.mkdir(exist_ok=True)

            # Generate the log filename
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

            if script_path:
                # Extract only the filename without path and extension
                script_name = os.path.basename(script_path)
                script_name = os.path.splitext(script_name)[0]
                log_filename = f"falconCommand_{script_name}_{timestamp}.log"
            else:
                # If no script path (e.g., direct command), use a generic name
                log_filename = f"command_{timestamp}.log"

            # Full path for the log file
            log_filepath = log_dir / log_filename

            # Write log content to file
            with open(log_filepath, "w", encoding="utf-8") as log_file:
                log_file.write(log_output)

            print(f"\nLog saved to: {log_filepath}")
            return str(log_filepath)

        except Exception as e:
            print(f"Error saving log: {str(e)}")
            return None

    def check_software(self, software_name):
        """
        Check if specified software is installed on the computer
        
        :param software_name: Name of the software to check (case-insensitive)
        :return: True if installed, False otherwise
        """
        try:
            import winreg
            import os
            
            software_found = False
            search_paths = [
                # Check uninstall registry keys (common installation location)
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                # Check Program Files directories
                None,  # Special case for Program Files check
            ]
            
            # Normalize search name
            software_name_lower = software_name.lower()
            
            for reg_path in search_paths:
                if reg_path is None:
                    # Check Program Files directories
                    program_files = [
                        os.environ.get('ProgramFiles', 'C:\\Program Files'),
                        os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)')
                    ]
                    
                    for program_dir in program_files:
                        if not os.path.exists(program_dir):
                            continue
                            
                        for item in os.listdir(program_dir):
                            item_path = os.path.join(program_dir, item)
                            if os.path.isdir(item_path) and software_name_lower in item.lower():
                                return True
                    continue
                    
                # Registry search
                hkey, subkey = reg_path
                try:
                    with winreg.OpenKey(hkey, subkey) as key:
                        # Enumerate subkeys
                        i = 0
                        while True:
                            try:
                                subkey_name = winreg.EnumKey(key, i)
                                with winreg.OpenKey(key, subkey_name) as subkey_handle:
                                    try:
                                        display_name = winreg.QueryValueEx(subkey_handle, "DisplayName")[0]
                                        if software_name_lower in display_name.lower():
                                            return True
                                    except (WindowsError, FileNotFoundError):
                                        pass  # No DisplayName value
                                i += 1
                            except WindowsError:
                                break  # No more subkeys
                except (WindowsError, FileNotFoundError):
                    pass  # Registry key not found
                    
            # Extra check for common executable patterns
            extensions = ['.exe', '.msi', '.app']
            for path in os.environ.get('PATH', '').split(os.pathsep):
                if path and os.path.exists(path):
                    for filename in os.listdir(path):
                        name, ext = os.path.splitext(filename.lower())
                        if ext in extensions and software_name_lower in name:
                            return True
            
            return False
        except Exception as e:
            print(f"Error checking for software: {str(e)}")
            return False

    def wait_until_exist(self, target_path, timeout=30, interval=1, confidence=0.8):
        """
        Wait until a target file, folder, or image exists
        
        This function checks for:
        1. Files or folders on disk (any file type: .txt, .pdf, .html, etc.)
        2. Images that appear on screen (.png, .jpg, .jpeg, .bmp, .gif)
        
        :param target_path: Path to the file, folder, or image
        :param timeout: Maximum waiting time in seconds
        :param interval: Interval between checks in seconds
        :param confidence: Image match confidence (0-1) - only for image detection on screen
        :return: True if found, False if timed out
        """
        # Define recognized image extensions for screen detection
        image_exts = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}
        target_path = target_path.strip('"\'')
        start = time.time()
        target = Path(target_path)
        target_ext = target.suffix.lower()
        
        # Print info about what we're checking for
        if target_ext in image_exts:
            print(f"Waiting for file or image: {target_path}")
            print(f"Will check both file system AND screen detection")
        else:
            print(f"Waiting for file: {target_path}")
        
        # Start waiting loop
        while time.time() - start < timeout:
            # Check if the file exists on disk first (for ALL file types)
            if target.exists():
                print(f"File found: {target_path}")
                return True
                
            # For image files, also try to find them on screen
            if target_ext in image_exts:
                try:
                    result = self.locate_image_multi_scale(
                        str(target), confidence=confidence
                    )
                    if result:
                        print(f"Image detected on screen: {target_path}")
                        return True
                except Exception as e:
                    print(f"[Image Detection Warning] {e}")
                    
            # Wait before next check
            time.sleep(interval)
            
            # Show progress indication
            elapsed = time.time() - start
            remaining = timeout - elapsed
            if int(elapsed) % 5 == 0 and elapsed > 0:  # Show progress every 5 seconds
                print(f"Still waiting... {int(remaining)}s remaining")
        
        # If we get here, we timed out
        print(f"[X] Timeout: File not found: {target_path}")
        
        return False

    def wait_until_process(self, process_name, timeout=30, interval=1):
        try:
            import psutil
        except ImportError:
            print("[X] psutil 模組未安裝，請先 pip install psutil")
            return False

        print(f"Waiting for process '{process_name}' to appear (timeout: {timeout}s)...")
        start = time.time()

        while time.time() - start < timeout:
            for proc in psutil.process_iter(attrs=["name"]):
                if proc.info["name"].lower() == process_name.lower():
                    print(f"[V] Process '{process_name}' is now running!")
                    return True
            time.sleep(interval)

        print(f"[X] Timeout: Process '{process_name}' not found within {timeout}s")
        return False


    def locate_image_multi_scale(
        self,
        template_path,
        scale_range=(0.6, 1.4),
        step=0.05,
        confidence=0.8,
        grayscale=True,
    ):
        # 擷取螢幕並轉為灰階圖
        screenshot = pyautogui.screenshot()
        screenshot_np = np.array(screenshot)

        if grayscale:
            screenshot_np = cv2.cvtColor(screenshot_np, cv2.COLOR_RGB2GRAY)

        # 讀取範本圖
        template = cv2.imread(template_path, 0 if grayscale else 1)
        if template is None:
            print(f"[X] 無法讀取圖片: {template_path}")
            return None

        best_match = None
        best_confidence = 0
        best_scale = 1.0
        best_position = None

        found_good_match = False
        last_confidence = 0
        drop_last_time = False

        for scale in np.arange(scale_range[0], scale_range[1], step):
            resized_template = cv2.resize(
                template, None, fx=scale, fy=scale, interpolation=cv2.INTER_LINEAR
            )

            # 若模板大於螢幕截圖則跳過
            if (
                resized_template.shape[0] > screenshot_np.shape[0]
                or resized_template.shape[1] > screenshot_np.shape[1]
            ):
                continue

            # 比對模板與螢幕截圖
            result = cv2.matchTemplate(
                screenshot_np, resized_template, cv2.TM_CCOEFF_NORMED
            )
            _, max_val, _, max_loc = cv2.minMaxLoc(result)

            # print(f"🔍 比例 {scale:.2f} → 信心值: {max_val:.3f}")

            # 判斷是否已找到過高信心值
            if max_val >= confidence:
                found_good_match = True

            if max_val > best_confidence:
                best_confidence = max_val
                best_position = max_loc
                best_match = resized_template
                best_scale = scale
                drop_last_time = False and drop_last_time
            else:
                drop_last_time = True or drop_last_time
                # 若已找到好結果，後續又掉下來，就停止
                if found_good_match and drop_last_time and (max_val <= confidence):
                    # print("連續信心值低於門檻，提早結束搜尋")
                    break

        if best_confidence >= confidence:
            h, w = best_match.shape
            # print(f"找到最佳匹配：scale={best_scale:.2f}, confidence={best_confidence:.3f}")
            return {
                "left": best_position[0],
                "top": best_position[1],
                "width": w,
                "height": h,
            }  # x, y, width, height
        else:
            # print(f"[X] 找不到符合門檻 ({confidence}) 的匹配，最高為 {best_confidence:.3f}")
            return None

    def locate_image(
        self, image_path, confidence=0.8, timeout=None, show_location=False
    ):
        """
        Position the image and click its center
        """
        try:
            # Check if the image exists
            if not Path(image_path).exists():
                raise FileNotFoundError(f"Image file not found: {image_path}")

            start_time = time.time()
            location = None

            while True:
                try:
                    # Try to locate the image
                    location = self.locate_image_multi_scale(
                        image_path, confidence=confidence, grayscale=True
                    )
                    # location = pyautogui.locateOnScreen(
                    #     image_path, confidence=confidence, grayscale=True
                    # )

                    if location:
                        break

                    if timeout and (time.time() - start_time > float(timeout)):
                        raise TimeoutError(
                            f"Could not find image within {timeout} seconds"
                        )

                    time.sleep(0.5)

                except TimeoutError:
                    raise
                except Exception as e:
                    if not timeout:
                        raise Exception(f"Error locating image: {str(e)}")
                    # If timeout is set, continue waiting until timeout
                    if time.time() - start_time > float(timeout):
                        raise TimeoutError(
                            f"Could not find image within {timeout} seconds"
                        )

            if location:
                # Calculate the center point
                center_x = location["left"] + location["width"] // 2
                center_y = location["top"] + location["height"] // 2
                # center_x = location.left + location.width // 2
                # center_y = location.top + location.height // 2

                if show_location:
                    try:
                        from PIL import ImageDraw, ImageGrab

                        # Screen capture
                        screen = ImageGrab.grab()
                        draw = ImageDraw.Draw(screen)
                        # Draw the box and center point
                        draw.rectangle(
                            [
                                location["left"],
                                location["top"],
                                location["left"] + location["width"],
                                location["top"] + location["height"],
                                # location.left,
                                # location.top,
                                # location.left + location.width,
                                # location.top + location.height,
                            ],
                            outline="red",
                        )
                        draw.line(
                            [center_x - 10, center_y, center_x + 10, center_y],
                            fill="red",
                        )
                        draw.line(
                            [center_x, center_y - 10, center_x, center_y + 10],
                            fill="red",
                        )
                        # Display the image
                        screen.show()
                    except ImportError:
                        print(
                            "[Warning] Pillow package not installed. Cannot show location."
                        )
                return center_x, center_y

            return None
        except Exception as e:
            raise Exception(f"Error processing image: {str(e)}")

    def locate_and_click_image(
        self, image_path, confidence=0.8, timeout=None, show_location=False
    ):
        """
        Position the image and click its center
        """
        try:
            image_laoc = self.locate_image(
                image_path, confidence, timeout, show_location
            )
            if image_laoc != None:
                pyautogui.click(image_laoc[0], image_laoc[1])
            return image_laoc
            """
            # Check if the image exists
            if not Path(image_path).exists():
                raise FileNotFoundError(f"Image file not found: {image_path}")

            start_time = time.time()
            location = None

            while True:
                try:
                    # Try to locate the image
                    location = self.locate_image_multi_scale(
                        image_path, confidence=confidence, grayscale=True
                    )
                    # location = pyautogui.locateOnScreen(
                    #     image_path, confidence=confidence, grayscale=True
                    # )

                    if location:
                        break

                    if timeout and (time.time() - start_time > float(timeout)):
                        raise TimeoutError(
                            f"Could not find image within {timeout} seconds"
                        )

                    time.sleep(0.5)

                except TimeoutError:
                    raise
                except Exception as e:
                    if not timeout:
                        raise Exception(f"Error locating image: {str(e)}")
                    # If timeout is set, continue waiting until timeout
                    if time.time() - start_time > float(timeout):
                        raise TimeoutError(
                            f"Could not find image within {timeout} seconds"
                        )

            if location:
                # Calculate the center point
                center_x = location["left"] + location["width"] // 2
                center_y = location["top"] + location["height"] // 2
                # center_x = location.left + location.width // 2
                # center_y = location.top + location.height // 2

                if show_location:
                    try:
                        from PIL import ImageDraw, ImageGrab

                        # Screen capture
                        screen = ImageGrab.grab()
                        draw = ImageDraw.Draw(screen)
                        # Draw the box and center point
                        draw.rectangle(
                            [
                                location["left"],
                                location["top"],
                                location["left"] + location["width"],
                                location["top"] + location["height"],
                                # location.left,
                                # location.top,
                                # location.left + location.width,
                                # location.top + location.height,
                            ],
                            outline="red",
                        )
                        draw.line(
                            [center_x - 10, center_y, center_x + 10, center_y],
                            fill="red",
                        )
                        draw.line(
                            [center_x, center_y - 10, center_x, center_y + 10],
                            fill="red",
                        )
                        # Display the image
                        screen.show()
                    except ImportError:
                        print(
                            "[Warning] Pillow package not installed. Cannot show location."
                        )

                # Move to the center point and click
                pyautogui.click(center_x, center_y)
                return center_x, center_y

            return None
            """

        except Exception as e:
            raise e

    def locate_and_right_click_image(
        self, image_path, confidence=0.8, timeout=None, show_location=False
    ):
        """
        Position the image and click its center
        """
        try:
            image_laoc = self.locate_image(
                image_path, confidence, timeout, show_location
            )
            if image_laoc != None:
                pyautogui.rightClick(x=image_laoc[0], y=image_laoc[1])
            return image_laoc
            """
            # Check if the image exists
            if not Path(image_path).exists():
                raise FileNotFoundError(f"Image file not found: {image_path}")

            start_time = time.time()
            location = None

            while True:
                try:
                    # Try to locate the image
                    location = self.locate_image_multi_scale(
                        image_path, confidence=confidence, grayscale=True
                    )
                    # location = pyautogui.locateOnScreen(
                    #     image_path, confidence=confidence, grayscale=True
                    # )

                    if location:
                        break

                    if timeout and (time.time() - start_time > float(timeout)):
                        raise TimeoutError(
                            f"Could not find image within {timeout} seconds"
                        )

                    time.sleep(0.5)

                except TimeoutError:
                    raise
                except Exception as e:
                    if not timeout:
                        raise Exception(f"Error locating image: {str(e)}")
                    # If timeout is set, continue waiting until timeout
                    if time.time() - start_time > float(timeout):
                        raise TimeoutError(
                            f"Could not find image within {timeout} seconds"
                        )

            if location:
                # Calculate the center point
                center_x = location["left"] + location["width"] // 2
                center_y = location["top"] + location["height"] // 2
                # center_x = location.left + location.width // 2
                # center_y = location.top + location.height // 2

                if show_location:
                    try:
                        from PIL import ImageDraw, ImageGrab

                        # Screen capture
                        screen = ImageGrab.grab()
                        draw = ImageDraw.Draw(screen)
                        # Draw the box and center point
                        draw.rectangle(
                            [
                                location["left"],
                                location["top"],
                                location["left"] + location["width"],
                                location["top"] + location["height"],
                                # location.left,
                                # location.top,
                                # location.left + location.width,
                                # location.top + location.height,
                            ],
                            outline="red",
                        )
                        draw.line(
                            [center_x - 10, center_y, center_x + 10, center_y],
                            fill="red",
                        )
                        draw.line(
                            [center_x, center_y - 10, center_x, center_y + 10],
                            fill="red",
                        )
                        # Display the image
                        screen.show()
                    except ImportError:
                        print(
                            "[Warning] Pillow package not installed. Cannot show location."
                        )

                # Move to the center point and click
                pyautogui.rightClick(x=center_x, y=center_y)
                # pyautogui.click(center_x, center_y)
                return center_x, center_y

            return None
            """
        except Exception as e:
            raise e

    def locate_and_double_click_image(
        self, image_path, confidence=0.8, timeout=None, show_location=False
    ):
        """
        Locate the image and double-click its center point

        Parameters:
        image_path (str): image file path
        confidence (float): confidence of the image match (0-1)
        timeout (float): The timeout for finding the image (in seconds). None means returning immediately.
        show_location (bool): whether to show the found location

        Returns:
        tuple: image center coordinates (x, y), returns None if not found
        """
        try:
            image_laoc = self.locate_image(
                image_path, confidence, timeout, show_location
            )
            if image_laoc != None:
                pyautogui.doubleClick(image_laoc[0], image_laoc[1], interval=0.1)
            return image_laoc
            """
            if not Path(image_path).exists():
                raise FileNotFoundError(f"Image file not found: {image_path}")

            start_time = time.time()
            location = None

            while True:
                try:
                    location = self.locate_image_multi_scale(
                        image_path, confidence=confidence, grayscale=True
                    )
                    # location = pyautogui.locateOnScreen(
                    #     image_path, confidence=confidence, grayscale=True
                    # )
                    if location:
                        break

                    if timeout and (time.time() - start_time > float(timeout)):
                        raise TimeoutError(
                            f"Could not find image within {timeout} seconds"
                        )

                    time.sleep(0.5)

                except TimeoutError:
                    raise
                except Exception as e:
                    if not timeout:
                        raise Exception(f"Error locating image: {str(e)}")
                    if time.time() - start_time > float(timeout):
                        raise TimeoutError(
                            f"Could not find image within {timeout} seconds"
                        )
            if location:
                center_x = location["left"] + location["width"] // 2
                center_y = location["top"] + location["height"] // 2
                # center_x = location.left + location.width // 2
                # center_y = location.top + location.height // 2
                if show_location:
                    try:
                        from PIL import ImageDraw, ImageGrab

                        screen = ImageGrab.grab()
                        draw = ImageDraw.Draw(screen)
                        draw.rectangle(
                            [
                                location["left"],
                                location["top"],
                                location["left"] + location["width"],
                                location["top"] + location["height"],
                                # location.left,
                                # location.top,
                                # location.left + location.width,
                                # location.top + location.height,
                            ],
                            outline="red",
                        )
                        draw.line(
                            [center_x - 10, center_y, center_x + 10, center_y],
                            fill="red",
                        )
                        draw.line(
                            [center_x, center_y - 10, center_x, center_y + 10],
                            fill="red",
                        )
                        screen.show()
                    except ImportError:
                        print(
                            "[Warning] Pillow package not installed. Cannot show location."
                        )
                pyautogui.doubleClick(center_x, center_y)
                return center_x, center_y

            return None
            """

        except Exception as e:
            raise e

    def execute_command_file(self, file_path, stop_on_error=True):
        """
        Execute the commands in the command file
        File format example:
        --click 100 200
        --delay 1
        --type "Hello World"
        """
        # Create a StringIO object to capture output
        log_buffer = io.StringIO()

        try:
            # Write header to log
            log_buffer.write(f"=== Falcon Command Execution Log ===\n")
            log_buffer.write(f"Script: {file_path}\n")
            log_buffer.write(
                f"Date/Time: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            )
            log_buffer.write(f"Stop on Error: {stop_on_error}\n\n")
            log_buffer.write("=== Command Execution ===\n\n")

            with open(file_path, "r", encoding="utf-8") as file:
                commands = []
                for line in file:
                    line_msg = f"Reading line: {line.strip()}"
                    log_buffer.write(line_msg + "\n")

                    # Ignore blank lines and comments
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    # try:
                    #     # Handling spaces between quotes
                    #     parts = []
                    #     current_part = ""
                    #     in_quotes = False
                    #     i = 0

                    #     while i < len(line):
                    #         char = line[i]
                    #         # Handle escaped quotes (\")
                    #         if (
                    #             char == "\\"
                    #             and i + 1 < len(line)
                    #             and line[i + 1] == '"'
                    #         ):
                    #             current_part += '"'
                    #             i += 2
                    #             continue
                    #         # Handling quotes
                    #         if char == '"':
                    #             in_quotes = not in_quotes
                    #             i += 1
                    #             continue
                    #         # Handling space
                    #         if char == " " and not in_quotes:
                    #             if current_part:
                    #                 parts.append(current_part)
                    #                 current_part = ""
                    #             i += 1
                    #             continue
                    #         # Add other characters to the current part
                    #         current_part += char
                    #         i += 1
                    #     # Add the last part
                    #     if current_part:
                    #         parts.append(current_part)
                    #     if parts:  # Make sure the command is parsed
                    #         commands.append(parts)
                    try:
                        # Handling spaces between quotes
                        parts = []
                        current_part = ""
                        in_quotes = False
                        i = 0

                        while i < len(line):
                            char = line[i]
                            # Handle escaped quotes (\")
                            if (
                                char == "\\"
                                and i + 1 < len(line)
                                and line[i + 1] == '"'
                            ):
                                current_part += '"'
                                i += 2
                                continue
                            # Handling quotes - 保持引號
                            if char == '"':
                                in_quotes = not in_quotes
                                current_part += char  # 把引號加入 current_part
                                i += 1
                                continue
                            # Handling space
                            if char == " " and not in_quotes:
                                if current_part:
                                    parts.append(current_part)
                                    current_part = ""
                                i += 1
                                continue
                            # Add other characters to the current part
                            current_part += char
                            i += 1
                        # Add the last part
                        if current_part:
                            parts.append(current_part)
                        if parts:  # Make sure the command is parsed
                            commands.append(parts)
                    except Exception as e:
                        error_msg = f"[Warning] Could not parse line: {line}"
                        print(error_msg)
                        log_buffer.write(error_msg + "\n")
                        error_msg = f"[Error] {str(e)}"
                        print(error_msg)
                        log_buffer.write(error_msg + "\n")
                        continue

                # execute all commands
                for cmd in commands:
                    try:
                        cmd_msg = f"[command] {' '.join(cmd)}"
                        print(cmd_msg)
                        log_buffer.write(cmd_msg + "\n")

                        # We need to modify the run method slightly for command file execution
                        # to prevent it from saving logs for each command
                        # Set a flag to indicate we're running from command file
                        self._running_from_command_file = True

                        # Execute the command without saving individual logs
                        result = self.run(cmd)

                        # Reset the flag
                        self._running_from_command_file = False

                        time.sleep(
                            self.parser.get_default("delay") or 0.1
                        )  # delay between commands

                        # check result
                        if result != 0 and stop_on_error:
                            error_msg = f"Command failed with exit code {result}, stopping execution."
                            print(error_msg)
                            log_buffer.write(error_msg + "\n")

                            # Only save the main log, not the intermediate logs
                            if file_path.endswith(".temp"):
                                # Skip saving log for .temp files
                                return result
                            else:
                                self.save_log_to_file(log_buffer.getvalue(), file_path)
                                return result

                    except Exception as e:
                        error_msg = f"Error executing command {cmd}: {str(e)}"
                        print(error_msg)
                        log_buffer.write(error_msg + "\n")

                        if stop_on_error:
                            error_msg = f"Command execution failed: {str(e)}"
                            log_buffer.write(error_msg + "\n")

                            # Only save log for the original file, not temp files
                            if not file_path.endswith(".temp"):
                                self.save_log_to_file(log_buffer.getvalue(), file_path)
                            return 1
                        continue

                # Write completion timestamp
                log_buffer.write(
                    f"\n=== Execution completed at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n"
                )

                # Only save log for the original file, not temp files
                if not file_path.endswith(".temp"):
                    self.save_log_to_file(log_buffer.getvalue(), file_path)
                return 0  # All commands completed successfully

        except FileNotFoundError:
            error_msg = f"Command file not found: {file_path}"
            log_buffer.write(error_msg + "\n")

            # Save log before exiting (only for original file)
            if not file_path.endswith(".temp"):
                self.save_log_to_file(log_buffer.getvalue(), file_path)
            raise FileNotFoundError(error_msg)
        except Exception as e:
            error_msg = f"Error reading command file: {str(e)}"
            log_buffer.write(error_msg + "\n")

            # Save log before exiting (only for original file)
            if not file_path.endswith(".temp"):
                self.save_log_to_file(log_buffer.getvalue(), file_path)
            raise Exception(error_msg)

    def execute_run_command(self, args_list):
        """
        Execute the run command to repeatedly perform a specific operation or a series of operations continuously.
        Format: --run command [command_args...] [--repeat N]
        For example: --run click 100 200 --repeat 5 # Click the coordinates (100, 200) 5 times
        """
        try:
            repeat_index = -1
            repeat_count = 1
            # Check if there is a --repeat parameter
            for i, arg in enumerate(args_list):
                if arg == "--repeat" and i + 1 < len(args_list):
                    try:
                        repeat_count = int(args_list[i + 1])
                        repeat_index = i
                        break
                    except ValueError:
                        print(
                            f"[Warning] Invalid repeat count '{args_list[i+1]}', using default 1"
                        )
            # Extract the command to be executed
            if repeat_index != -1:
                command_args = args_list[:repeat_index]
            else:
                command_args = args_list

            if not command_args:
                print("[Error] No command specified for --run")
                return 1

            # Convert the first argument to a full command format (e.g. 'click' -> '--click')
            command = command_args[0]
            if not command.startswith("--"):
                command = f"--{command}"

            command_args[0] = command

            print(f"Running command: {' '.join(command_args)} (repeat: {repeat_count})")

            # Repeat the specified command
            for i in range(repeat_count):
                print(f"Execution {i+1}/{repeat_count}")
                self.run(command_args)
                if i < repeat_count - 1:  # If it is not the last execution, wait
                    time.sleep(self.parser.get_default("delay") or 0.1)

            return 0

        except Exception as e:
            print(f"Error executing run command: {str(e)}")
            return 1

    def fast_click(self, x=None, y=None, count=1, delay=0.01):
        """
        Perform super fast clicks, skipping intermediate processing steps by calling pyautogui directly

        Parameters:
        x (int, optional): X coordinate, if not provided the current mouse position is used
        y (int, optional): Y coordinate, if not provided the current mouse position is used
        count (int): number of clicks
        delay (float): click interval time (seconds)
        """
        if x is None or y is None:
            current_pos = pyautogui.position()
            if x is None:
                x = current_pos.x
            if y is None:
                y = current_pos.y

        pyautogui.moveTo(x, y)
        print(f"Fast clicking at ({x}, {y}) {int(count)} times with {delay}s delay")

        original_pause = pyautogui.PAUSE

        try:

            pyautogui.PAUSE = delay

            for i in range(int(count)):
                pyautogui.click(x=x, y=y)
        finally:
            pyautogui.PAUSE = original_pause

        print(f"Completed {int(count)} fast clicks")
        return 0

    def launch_application(self, exe_path):
        """
        Open the specified application, support applications that require administrator privileges

        Parameters:
        exe_path (str): the path of the application

        Returns:
        int: returns 0 if successful, 1 if failed
        """
        try:
            import ctypes
            import os
            import subprocess
            import sys
            exe_path = exe_path.strip('"\'')
            if not os.path.exists(exe_path):
                print(f"[Error] Application not found at '{exe_path}'")
                return 1

            if os.name == "nt":
                try:
                    
                    # Try to start directly using os.startfile first
                    os.startfile(exe_path)
                    print(f"Launched application: {exe_path}")
                except Exception as e:
                    print(f"Standard launch failed: {str(e)}")
                    print("Attempting to launch with administrator privileges...")

                    # If direct startup fails, try to use the runas command to start with administrator privileges
                    try:
                        # Use ShellExecute to start the application with administrator privileges
                        if sys.version_info[0] >= 3:
                            ctypes.windll.shell32.ShellExecuteW(
                                None, "runas", exe_path, None, None, 1
                            )
                        else:
                            ctypes.windll.shell32.ShellExecuteW(
                                None, "runas", unicode(exe_path), None, None, 1
                            )
                        print(f"Launched application with admin privileges: {exe_path}")
                    except Exception as admin_error:
                        # If starting with administrator privileges also fails, try using subprocess
                        print(f"Admin launch failed: {str(admin_error)}")
                        print("Trying alternative launch method...")

                        # Try starting with subprocess.Popen and special parameters
                        startupinfo = subprocess.STARTUPINFO()
                        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                        subprocess.Popen(
                            [exe_path], startupinfo=startupinfo, shell=True
                        )
                        print(
                            f"Launched application using alternative method: {exe_path}"
                        )
            else:
                print(f"[Error] Unsupported operating system")
                return 1

            return 0
        except Exception as e:
            print(f"Error launching application: {str(e)}")
            return 1
    def wait_until_installed(self, software_name, timeout=120, interval=3):
        """
        Wait until specified software is installed on the system
        
        :param software_name: Name of the software to check for installation
        :param timeout: Maximum time to wait in seconds
        :param interval: Check interval in seconds
        :return: True if software is installed within timeout, False if timed out
        """
        print(f"Waiting for software '{software_name}' to be installed... (timeout: {timeout}s)")
        log_message = f"Waiting for software '{software_name}' to be installed... (timeout: {timeout}s)"
        
        start_time = time.time()
        last_status_time = 0
        
        while time.time() - start_time < timeout:
            # Check if software is installed
            if self.check_software(software_name):
                elapsed = time.time() - start_time
                success_msg = f"[V] Software '{software_name}' is now installed! (detected after {elapsed:.1f}s)"
                print(success_msg)
                return True
                
            # Show periodic status updates, but not too frequently
            current_time = time.time()
            elapsed = current_time - start_time
            remaining = timeout - elapsed
            
            # Display status message every 10 seconds
            if current_time - last_status_time >= 10:
                status_msg = f"...Still waiting for '{software_name}' to be installed... ({int(remaining)}s remaining)"
                print(status_msg)
                last_status_time = current_time
                
            # Wait before next check
            time.sleep(interval)
        
        # If we get here, we timed out
        timeout_msg = f"[✗] Timeout: Software '{software_name}' was not installed within {timeout}s"
        print(timeout_msg)
        return False
    def run(self, args=None):

        log_buffer = io.StringIO()
        script_path = None

        if args is None:
            args = self.parser.parse_args()
        else:
            if isinstance(args, list):
                args = self.parser.parse_args(args)

        if hasattr(args, "command_file") and args.command_file:
            script_path = args.command_file

        # Set the stop_on_error flag for command file execution
        self.stop_on_error = (
            args.stop_on_error if hasattr(args, "stop_on_error") else False
        )

        pyautogui.PAUSE = args.delay
        timeout_sec = args.timeout

        try:

            if hasattr(args, "run") and args.run:
                result = self.execute_run_command(args.run)
                # Save log for run command
                if script_path is None:  # Only save here if not part of a command file
                    pass
                return result

            if hasattr(args, "click") and args.click is not None:
                if len(args.click) == 2:
                    x, y = args.click
                    pyautogui.click(x=x, y=y)
                    msg = f"Clicked at position ({x}, {y})"
                else:
                    x, y = pyautogui.position()
                    pyautogui.click()
                    msg = f"Clicked at current position ({x}, {y})"
                print(msg)
                log_buffer.write(msg + "\n")

            if hasattr(args, "double_click") and args.double_click is not None:
                if len(args.double_click) == 2:
                    x, y = args.double_click
                    pyautogui.doubleClick(x=x, y=y,interval=0.1)
                    msg = f"Double clicked at position ({x}, {y})"
                else:
                    x, y = pyautogui.position()
                    pyautogui.doubleClick()
                    msg = f"Double clicked at current position ({x}, {y})"
                print(msg)
                log_buffer.write(msg + "\n")

            if hasattr(args, "right_click") and args.right_click is not None:
                if len(args.right_click) == 2:
                    x, y = args.right_click
                    pyautogui.rightClick(x=x, y=y)
                    msg = f"Right clicked at position ({x}, {y})"
                else:
                    x, y = pyautogui.position()
                    pyautogui.rightClick()
                    msg = f"Right clicked at current position ({x}, {y})"
                print(msg)
                log_buffer.write(msg + "\n")

            if hasattr(args, "fast_click") and args.fast_click is not None:
                if len(args.fast_click) == 4:
                    x, y, count, delay = args.fast_click
                    msg = f"Fast clicking at ({x}, {y}) {int(count)} times with {delay}s delay"
                    print(msg)
                    log_buffer.write(msg + "\n")
                    return self.fast_click(x, y, int(count), delay)
                elif len(args.fast_click) == 2:
                    # Only count and delay are provided
                    count, delay = args.fast_click
                    msg = f"Fast clicking at current position {int(count)} times with {delay}s delay"
                    print(msg)
                    log_buffer.write(msg + "\n")
                    return self.fast_click(count=int(count), delay=delay)
                elif len(args.fast_click) == 1:
                    # Only count
                    count = args.fast_click[0]
                    msg = f"Fast clicking at current position {int(count)} times"
                    print(msg)
                    log_buffer.write(msg + "\n")
                    return self.fast_click(count=int(count))
                else:
                    msg = "Fast clicking at current position 1 time"
                    print(msg)
                    log_buffer.write(msg + "\n")
                    return self.fast_click()

            if args.moveto:
                if len(args.moveto) not in [2, 3]:
                    msg = "[Error] --moveto requires exactly 2 (X, Y) or 3 (X, Y, DURATION) arguments"
                    print(msg)
                    log_buffer.write(msg + "\n")
                    sys.exit(1)

                x, y = args.moveto[:2]
                duration = (
                    args.moveto[2] if len(args.moveto) == 3 else 0
                )  # default duration = 0

                if duration > 0:
                    pyautogui.moveTo(x, y, duration=duration)
                    msg = f"Moved to ({x}, {y}) over {duration} seconds"
                else:
                    pyautogui.moveTo(x, y)
                    msg = f"Moved to ({x}, {y})"
                print(msg)
                log_buffer.write(msg + "\n")

            if args.scroll:
                if len(args.scroll) == 1:
                    # Just scroll at current position
                    pyautogui.scroll(args.scroll[0])
                    msg = f"Scrolled {args.scroll[0]} clicks"
                    print(msg)
                    log_buffer.write(msg + "\n")
                elif len(args.scroll) == 3:
                    # First move to the specified location, then scroll
                    pyautogui.moveTo(args.scroll[1], args.scroll[2])
                    pyautogui.scroll(args.scroll[0])
                    msg = f"Scrolled {args.scroll[0]} clicks at position ({args.scroll[1]}, {args.scroll[2]})"
                    print(msg)
                    log_buffer.write(msg + "\n")
                else:
                    msg = "[Error] --scroll requires either 1 (CLICKS) or 3 (CLICKS, X, Y) arguments"
                    print(msg)
                    log_buffer.write(msg + "\n")
                    return 1

            if args.type:
                try:
                    # Ensure text is properly encoded for clipboard
                    text = args.type
                    # For direct typing fallback if needed
                    # pyautogui.write(text, interval=0.05)
                    
                    # Preferred clipboard method for Chinese characters
                    pyperclip.copy(text)
                    pyautogui.hotkey("ctrl", "v")
                    msg = f"Typed :{text}"
                    print(msg)
                    log_buffer.write(msg + "\n")
                except Exception as e:
                    error_msg = f"Error typing text: {str(e)}"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1
            if args.press:
                keys = args.press.split("+")
                if len(keys) > 1:
                    pyautogui.hotkey(*keys)  # support `ctrl+c`, `shift+tab`
                else:
                    pyautogui.press(keys[0])
                msg = f"Pressed key: {args.press}"
                print(msg)
                log_buffer.write(msg + "\n")

            if args.clipboard_copy:
                pyautogui.hotkey("ctrl", "c")
                msg = "Copied selection to clipboard"
                print(msg)
                log_buffer.write(msg + "\n")

            if args.clipboard_paste:
                pyautogui.hotkey("ctrl", "v")
                msg = "Pasted from clipboard"
                print(msg)
                log_buffer.write(msg + "\n")

            if args.clipboard_set:
                try:
                    pyperclip.copy(args.clipboard_set)
                    msg = f"Set clipboard content to: {args.clipboard_set}"
                    print(msg)
                    log_buffer.write(msg + "\n")
                except ImportError:
                    err_msg = "[Error] pyperclip module not installed. Cannot set clipboard content."
                    print(err_msg)
                    log_buffer.write(err_msg + "\n")

            if args.clipboard_get:
                try:
                    content = pyperclip.paste()
                    msg = f"Clipboard content: {content}"
                    print(msg)
                    log_buffer.write(msg + "\n")
                except ImportError:
                    err_msg = "[Error] pyperclip module not installed. Cannot get clipboard content."
                    print(err_msg)
                    log_buffer.write(err_msg + "\n")

            if args.screenshot:
                pyautogui.screenshot(args.screenshot)
                msg = f"Screenshot saved as: {args.screenshot}"
                print(msg)
                log_buffer.write(msg + "\n")

            if args.position:
                x, y = pyautogui.position()
                msg = f"Current mouse position: ({x}, {y})"
                print(msg)
                log_buffer.write(msg + "\n")

            if args.screen_size:
                width, height = pyautogui.size()
                msg = f"Screen size: {width}x{height} pixels"
                print(msg)
                log_buffer.write(msg + "\n")

            if args.window_info:
                try:
                    window = pyautogui.getWindowsWithTitle(args.window_info)[0]
                    window_info_msg = f"Window '{args.window_info}':"
                    print(window_info_msg)
                    log_buffer.write(window_info_msg + "\n")

                    position_msg = f"  Position: ({window.left}, {window.top})"
                    print(position_msg)
                    log_buffer.write(position_msg + "\n")

                    size_msg = f"  Size: {window.width}x{window.height}"
                    print(size_msg)
                    log_buffer.write(size_msg + "\n")

                    bottom_right_msg = f"  Bottom-right: ({window.left + window.width}, {window.top + window.height})"
                    print(bottom_right_msg)
                    log_buffer.write(bottom_right_msg + "\n")
                except IndexError:
                    error_msg = (
                        f"[Error] Window with title '{args.window_info}' not found"
                    )
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                except AttributeError:
                    error_msg = (
                        "[Error] Window operations not supported on this platform"
                    )
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")

            if args.track_position:
                tracking_msg = "Tracking mouse position. Press Ctrl+C to stop..."
                print(tracking_msg)
                log_buffer.write(tracking_msg + "\n")

                start_time = time.time()
                try:
                    while time.time() - start_time < args.track_position:
                        x, y = pyautogui.position()
                        position_str = f"X: {str(x).rjust(4)} Y: {str(y).rjust(4)}"
                        print(position_str, end="\r")
                        # Only log positions occasionally to avoid huge log files
                        if (
                            int((time.time() - start_time) * 10) % 10 == 0
                        ):  # Log every ~1 second
                            log_buffer.write(f"Position: {position_str}\n")
                        time.sleep(0.1)
                except KeyboardInterrupt:
                    stop_msg = "\nMouse position tracking stopped."
                    print(stop_msg)
                    log_buffer.write(stop_msg + "\n")

            if args.relative_move:
                current_x, current_y = pyautogui.position()
                new_x = current_x + args.relative_move[0]
                new_y = current_y + args.relative_move[1]
                pyautogui.moveRel(args.relative_move[0], args.relative_move[1])
                move_msg = (
                    f"Moved from ({current_x}, {current_y}) to ({new_x}, {new_y})"
                )
                print(move_msg)
                log_buffer.write(move_msg + "\n")

            if args.center_on_screen:
                screen_width, screen_height = pyautogui.size()
                center_x = screen_width // 2
                center_y = screen_height // 2
                pyautogui.moveTo(center_x, center_y)
                center_msg = f"Moved to screen center: ({center_x}, {center_y})"
                print(center_msg)
                log_buffer.write(center_msg + "\n")

            if args.position_to_clipboard:
                x, y = pyautogui.position()
                position_str = f"({x}, {y})"
                try:

                    pyperclip.copy(position_str)
                    clipboard_msg = (
                        f"Current position {position_str} copied to clipboard"
                    )
                    print(clipboard_msg)
                    log_buffer.write(clipboard_msg + "\n")
                except ImportError:
                    pyautogui.write(position_str)
                    fallback_msg = f"Current position {position_str} written (pyperclip not available)"
                    print(fallback_msg)
                    log_buffer.write(fallback_msg + "\n")

            if args.drag_to:
                if len(args.drag_to) not in [2, 3]:
                    error_msg = "[Error] --drag-to requires 2 (X, Y) or 3 (X, Y, DURATION) arguments"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1

                x, y = args.drag_to[:2]
                duration = args.drag_to[2] if len(args.drag_to) == 3 else 0.5
                start_x, start_y = pyautogui.position()
                pyautogui.dragTo(x, y, duration=duration)
                drag_msg = f"Dragged from ({start_x}, {start_y}) to ({x}, {y}) over {duration} seconds"
                print(drag_msg)
                log_buffer.write(drag_msg + "\n")

            if args.search_image:
                try:
                    click_img_msg = f"Attempting to find image: {args.search_image}"
                    print(click_img_msg)
                    log_buffer.write(click_img_msg + "\n")

                    result = self.locate_image(
                        args.search_image, timeout=timeout_sec, show_location=False
                    )
                    if result:
                        center_x, center_y = result
                        success_msg = (
                            f"Found image at center point ({center_x}, {center_y})"
                        )
                        print(success_msg)
                        log_buffer.write(success_msg + "\n")
                    else:
                        not_found_msg = "Image not found on screen"
                        print(not_found_msg)
                        log_buffer.write(not_found_msg + "\n")
                except Exception as e:
                    error_msg = f"[Error] {str(e)}"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1

            if args.click_image:
                try:
                    click_img_msg = (
                        f"Attempting to find and click image: {args.click_image}"
                    )
                    print(click_img_msg)
                    log_buffer.write(click_img_msg + "\n")

                    result = self.locate_and_click_image(
                        args.click_image, timeout=timeout_sec, show_location=False
                    )
                    if result:
                        center_x, center_y = result
                        success_msg = f"Found and clicked image at center point ({center_x}, {center_y})"
                        print(success_msg)
                        log_buffer.write(success_msg + "\n")
                    else:
                        not_found_msg = "Image not found on screen"
                        print(not_found_msg)
                        log_buffer.write(not_found_msg + "\n")
                except Exception as e:
                    error_msg = f"[Error] {str(e)}"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1

            if args.right_click_image:
                try:
                    click_img_msg = f"Attempting to find and right click image: {args.right_click_image}"
                    print(click_img_msg)
                    log_buffer.write(click_img_msg + "\n")

                    result = self.locate_and_right_click_image(
                        args.right_click_image, timeout=timeout_sec, show_location=False
                    )
                    if result:
                        center_x, center_y = result
                        success_msg = f"Found and right clicked image at center point ({center_x}, {center_y})"
                        print(success_msg)
                        log_buffer.write(success_msg + "\n")
                    else:
                        not_found_msg = "Image not found on screen"
                        print(not_found_msg)
                        log_buffer.write(not_found_msg + "\n")
                except Exception as e:
                    error_msg = f"[Error] {str(e)}"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1

            if args.double_click_image:
                try:
                    dbl_click_img_msg = f"Attempting to find and double-click image: {args.double_click_image}"
                    print(dbl_click_img_msg)
                    log_buffer.write(dbl_click_img_msg + "\n")

                    result = self.locate_and_double_click_image(
                        args.double_click_image,
                        timeout=timeout_sec,
                        show_location=False,
                    )
                    if result:
                        center_x, center_y = result
                        success_msg = f"Found and double-clicked image at center point ({center_x}, {center_y})"
                        print(success_msg)
                        log_buffer.write(success_msg + "\n")
                    else:
                        not_found_msg = "Image not found on screen"
                        print(not_found_msg)
                        log_buffer.write(not_found_msg + "\n")
                except Exception as e:
                    error_msg = f"[Error] {str(e)}"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1

            if args.image_exists:
                try:
                    img_exists_msg = (
                        f"Checking if image exists on screen: {args.image_exists}"
                    )
                    print(img_exists_msg)
                    log_buffer.write(img_exists_msg + "\n")

                    location = pyautogui.locateOnScreen(
                        args.image_exists, confidence=0.8, grayscale=True
                    )
                    if location:
                        found_msg = f"Image found at location: ({location.left}, {location.top})"
                        print(found_msg)
                        log_buffer.write(found_msg + "\n")
                        return 0
                    else:
                        not_found_msg = "Image not found on screen"
                        print(not_found_msg)
                        log_buffer.write(not_found_msg + "\n")
                        return 1
                except Exception as e:
                    error_msg = f"Error checking for image: {str(e)}"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1

            if args.command_file:
                try:
                    cmd_file_msg = f"[command from file] {args.command_file}"
                    print(cmd_file_msg)
                    log_buffer.write(cmd_file_msg + "\n")

                    # Save the log before executing the command file, which has its own logging
                    # self.save_log_to_file(log_buffer.getvalue(), script_path)

                    # Pass stop_on_error flag to execute_command_file
                    return self.execute_command_file(
                        args.command_file, self.stop_on_error
                    )
                except Exception as e:
                    error_msg = f"Error executing command file: {str(e)}"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1

            if args.launch:
                try:
                    launch_msg = f"Launching application: {args.launch}"
                    print(launch_msg)
                    log_buffer.write(launch_msg + "\n")

                    result = self.launch_application(args.launch)
                    if result == 0:
                        success_msg = (
                            f"Application launched successfully: {args.launch}"
                        )
                        log_buffer.write(success_msg + "\n")
                    else:
                        failure_msg = f"Failed to launch application: {args.launch}"
                        log_buffer.write(failure_msg + "\n")
                    return result
                except Exception as e:
                    error_msg = f"Error launching application: {str(e)}"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1
            if args.check_software:
                try:
                    software_name = args.check_software
                    msg = f"Checking if '{software_name}' is installed..."
                    print(msg)
                    log_buffer.write(msg + "\n")
                    
                    is_installed = self.check_software(software_name)
                    
                    if is_installed:
                        result_msg = f"[V] Software '{software_name}' is installed on this computer."
                        print(result_msg)
                        log_buffer.write(result_msg + "\n")
                        return 0
                    else:
                        result_msg = f"[X] Software '{software_name}' is not found on this computer."
                        print(result_msg)
                        log_buffer.write(result_msg + "\n")
                        return 1
                except Exception as e:
                    error_msg = f"Error checking for software: {str(e)}"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1
            if args.wait_until_exist:
                effective_timeout = None
                if hasattr(args, 'wait_time') and args.wait_time is not None:
                    effective_timeout = args.wait_time
                else:
                    effective_timeout = args.timeout or 30
                if not self.wait_until_exist(
                    args.wait_until_exist,
                    timeout=args.timeout or 30,
                    interval=args.check_interval or 1,
                ):
                    print(f"[X] Timeout: {args.wait_until_exist} not found")
                    return 1

            if args.wait_until_process:
                # Use wait_time if provided, otherwise fall back to timeout or the default
                effective_timeout = None
                if hasattr(args, 'wait_time') and args.wait_time is not None:
                    effective_timeout = args.wait_time
                else:
                    effective_timeout = args.timeout or 30
                    
                if not self.wait_until_process(
                    args.wait_until_process,
                    timeout=effective_timeout,
                    interval=args.check_interval or 1,
                ):
                    print(f"[X] Timeout: Process '{args.wait_until_process}' not found")
                    return 1
            if args.wait_until_installed:
                software_name = args.wait_until_installed
                
                # Determine timeout and interval values
                timeout = args.wait_time if args.wait_time is not None else (args.timeout or 120)
                interval = args.check_interval or 3
                
                wait_msg = f"Waiting for software '{software_name}' to be installed..."
                print(wait_msg)
                log_buffer.write(wait_msg + "\n")
                
                if self.wait_until_installed(software_name, timeout=timeout, interval=interval):
                    success_msg = f"[V] Software '{software_name}' is now installed successfully."
                    print(success_msg)
                    log_buffer.write(success_msg + "\n")
                    return 0
                else:
                    timeout_msg = f"[✗] Timeout: Software '{software_name}' was not installed within {timeout}s"
                    print(timeout_msg)
                    log_buffer.write(timeout_msg + "\n")
                    return 1

            if args.sleep:
                sleep_msg = f"Waiting for {args.sleep} seconds..."
                print(sleep_msg)
                log_buffer.write(sleep_msg + "\n")

                time.sleep(args.sleep)

                continue_msg = f"Continue after {args.sleep} seconds..."
                print(continue_msg)
                log_buffer.write(continue_msg + "\n")

            # if script_path is None:
            #     self.save_log_to_file(log_buffer.getvalue())

            if (
                not hasattr(self, "_running_from_command_file")
                or not self._running_from_command_file
            ):
                if not hasattr(args, "command_file") or (
                    script_path and not script_path.endswith(".temp")
                ):
                    self.save_log_to_file(log_buffer.getvalue(), script_path)

        except KeyboardInterrupt:
            error_msg = "\nOperation cancelled by user. Exiting..."
            print(error_msg)
            log_buffer.write(error_msg + "\n")
            if script_path and not script_path.endswith(".temp"):
                self.save_log_to_file(log_buffer.getvalue(), script_path)
            sys.exit(1)
        except Exception as e:
            error_msg = f"[Error] {str(e)}"
            print(error_msg)
            log_buffer.write(error_msg + "\n")
            # Save log before exiting, but only for the original file
            if script_path and not script_path.endswith(".temp"):
                self.save_log_to_file(log_buffer.getvalue(), script_path)

        return 0


if __name__ == "__main__":
    controller = AutoGUIController()
    sys.exit(controller.run())
