# falconCommand.py
#
import argparse
import datetime
import io
import os
import sys
import time
from pathlib import Path
import pyperclip
import pyautogui

COMMAND_VERSION = "1.0.2"  # Add version number here


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
            "--click-image",
            type=str,
            metavar="IMAGE_PATH",
            help="Click the center of the specified image on screen",
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
        return parser
    
    def save_log_to_file(self, log_output, script_path=None):
        """
        Save execution log to C:\Falcon_Log directory
        
        Parameters:
        log_output (str): The log content to save
        script_path (str, optional): Path of the executed script file
        """
        try:
            # Create Falcon_Log directory if it doesn't exist
            log_dir = Path("C:/Falcon_Log")
            log_dir.mkdir(exist_ok=True)
            
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
                
            print(f"Log saved to: {log_filepath}")
            return str(log_filepath)
                
        except Exception as e:
            print(f"Error saving log: {str(e)}")
            return None
        
    def locate_and_click_image(
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
                    location = pyautogui.locateOnScreen(
                        image_path, confidence=confidence, grayscale=True
                    )

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
                center_x = location.left + location.width // 2
                center_y = location.top + location.height // 2

                if show_location:
                    try:
                        from PIL import ImageDraw, ImageGrab

                        # Screen capture
                        screen = ImageGrab.grab()
                        draw = ImageDraw.Draw(screen)
                        # Draw the box and center point
                        draw.rectangle(
                            [
                                location.left,
                                location.top,
                                location.left + location.width,
                                location.top + location.height,
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
                            "Warning: Pillow package not installed. Cannot show location."
                        )

                # Move to the center point and click
                pyautogui.click(center_x, center_y)
                return center_x, center_y

            return None

        except Exception as e:
            raise Exception(f"Error processing image: {str(e)}")

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
            if not Path(image_path).exists():
                raise FileNotFoundError(f"Image file not found: {image_path}")

            start_time = time.time()
            location = None

            while True:
                try:
                    location = pyautogui.locateOnScreen(
                        image_path, confidence=confidence, grayscale=True
                    )
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
                center_x = location.left + location.width // 2
                center_y = location.top + location.height // 2
                if show_location:
                    try:
                        from PIL import ImageDraw, ImageGrab

                        screen = ImageGrab.grab()
                        draw = ImageDraw.Draw(screen)
                        draw.rectangle(
                            [
                                location.left,
                                location.top,
                                location.left + location.width,
                                location.top + location.height,
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
                            "Warning: Pillow package not installed. Cannot show location."
                        )
                pyautogui.doubleClick(center_x, center_y)
                return center_x, center_y

            return None

        except Exception as e:
            raise Exception(f"Error processing image: {str(e)}")

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
            log_buffer.write(f"Date/Time: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
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
                            # Handling quotes
                            if char == '"':
                                in_quotes = not in_quotes
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
                        error_msg = f"Warning: Could not parse line: {line}"
                        print(error_msg)
                        log_buffer.write(error_msg + "\n")
                        error_msg = f"Error: {str(e)}"
                        print(error_msg)
                        log_buffer.write(error_msg + "\n")
                        continue

                # execute all commands
                for cmd in commands:
                    try:
                        cmd_msg = f"Executing command: {' '.join(cmd)}"
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
                log_buffer.write(f"\n=== Execution completed at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n")
                
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
                            f"Warning: Invalid repeat count '{args_list[i+1]}', using default 1"
                        )
            # Extract the command to be executed
            if repeat_index != -1:
                command_args = args_list[:repeat_index]
            else:
                command_args = args_list

            if not command_args:
                print("Error: No command specified for --run")
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

            if not os.path.exists(exe_path):
                print(f"Error: Application not found at '{exe_path}'")
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
                print(f"Error: Unsupported operating system")
                return 1

            return 0
        except Exception as e:
            print(f"Error launching application: {str(e)}")
            return 1

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
                    pyautogui.doubleClick(x=x, y=y)
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
                    msg = "Error: --moveto requires exactly 2 (X, Y) or 3 (X, Y, DURATION) arguments"
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
                    msg = "Error: --scroll requires either 1 (CLICKS) or 3 (CLICKS, X, Y) arguments"
                    print(msg)
                    log_buffer.write(msg + "\n")
                    return 1

            if args.type:
                pyautogui.write(
                    args.type, interval=0.05
                )  # Increase the interval to avoid typing too fast
                msg = f"Typed: {args.type}"
                print(msg)
                log_buffer.write(msg + "\n")
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
                    err_msg = "Error: pyperclip module not installed. Cannot set clipboard content."
                    print(err_msg)
                    log_buffer.write(err_msg + "\n")

            if args.clipboard_get:
                try:
                    content = pyperclip.paste()
                    msg = f"Clipboard content: {content}"
                    print(msg)
                    log_buffer.write(msg + "\n")
                except ImportError:
                    err_msg = "Error: pyperclip module not installed. Cannot get clipboard content."
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
                    error_msg = f"Error: Window with title '{args.window_info}' not found"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                except AttributeError:
                    error_msg = "Error: Window operations not supported on this platform"
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
                        if int((time.time() - start_time) * 10) % 10 == 0:  # Log every ~1 second
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
                move_msg = f"Moved from ({current_x}, {current_y}) to ({new_x}, {new_y})"
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
                    clipboard_msg = f"Current position {position_str} copied to clipboard"
                    print(clipboard_msg)
                    log_buffer.write(clipboard_msg + "\n")
                except ImportError:
                    pyautogui.write(position_str)
                    fallback_msg = f"Current position {position_str} written (pyperclip not available)"
                    print(fallback_msg)
                    log_buffer.write(fallback_msg + "\n")

            if args.drag_to:
                if len(args.drag_to) not in [2, 3]:
                    error_msg = "Error: --drag-to requires 2 (X, Y) or 3 (X, Y, DURATION) arguments"
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

            if args.click_image:
                try:
                    click_img_msg = f"Attempting to find and click image: {args.click_image}"
                    print(click_img_msg)
                    log_buffer.write(click_img_msg + "\n")
                    
                    result = self.locate_and_click_image(
                        args.click_image, show_location=False
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
                    error_msg = f"Error: {str(e)}"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1

            if args.double_click_image:
                try:
                    dbl_click_img_msg = f"Attempting to find and double-click image: {args.double_click_image}"
                    print(dbl_click_img_msg)
                    log_buffer.write(dbl_click_img_msg + "\n")
                    
                    result = self.locate_and_double_click_image(
                        args.double_click_image, show_location=False
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
                    error_msg = f"Error: {str(e)}"
                    print(error_msg)
                    log_buffer.write(error_msg + "\n")
                    return 1

            if args.image_exists:
                try:
                    img_exists_msg = f"Checking if image exists on screen: {args.image_exists}"
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
                    cmd_file_msg = f"Executing commands from file: {args.command_file}"
                    print(cmd_file_msg)
                    log_buffer.write(cmd_file_msg + "\n")
                    
                    # Save the log before executing the command file, which has its own logging
                    #self.save_log_to_file(log_buffer.getvalue(), script_path)
                    
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
                        success_msg = f"Application launched successfully: {args.launch}"
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

            if not hasattr(self, '_running_from_command_file') or not self._running_from_command_file:
                if not hasattr(args, "command_file") or (script_path and not script_path.endswith(".temp")):
                    self.save_log_to_file(log_buffer.getvalue(), script_path) 

        except KeyboardInterrupt:
            error_msg = "\nOperation cancelled by user. Exiting..."
            print(error_msg)
            log_buffer.write(error_msg + "\n")
            if script_path and not script_path.endswith(".temp"):
                self.save_log_to_file(log_buffer.getvalue(), script_path)
            sys.exit(1)
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            print(error_msg)
            log_buffer.write(error_msg + "\n")
            # Save log before exiting, but only for the original file
            if script_path and not script_path.endswith(".temp"):
                self.save_log_to_file(log_buffer.getvalue(), script_path)
        
        return 0


if __name__ == "__main__":
    controller = AutoGUIController()
    sys.exit(controller.run())
