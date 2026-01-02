"""
Background worker thread for GUI operations.
Prevents UI from freezing during long-running tasks.
"""

import threading
import sys
from io import StringIO


class WorkerThread(threading.Thread):
    """
    Executes automation tasks in a background thread.
    Captures console output and reports progress back to GUI.
    """

    def __init__(self, task_func, callback, error_callback, progress_callback=None):
        """
        Initialize worker thread.

        Args:
            task_func: Function to execute (e.g. automation.fill_images)
            callback: Function to call on success with output text
            error_callback: Function to call on error with error message
            progress_callback: Function to call with progress updates (current, total, message)
        """

        super().__init__(daemon=True)
        self.task_func = task_func
        self.callback = callback
        self.error_callback = error_callback
        self.progress_callback = progress_callback
        self.output = StringIO()

    def run(self):
        """
        Execute the task and capture all console output.
        Called automatically when thread starts
        """
        # Redirect stdout to capture print statements
        old_stdout = sys.stdout
        sys.stdout = self.output

        try:
            # Execute the task function
            result = self.task_func()

            # Restore stdout
            sys.stdout = old_stdout

            # get captured output
            output_text = self.output.getvalue()

            # Report success back to GUI
            if result:
                self.callback(output_text + "\n✅ Operation completed successfully!")
            else:
                self.error_callback("Operation failed. Check the log for details.\n" + output_text)

        except Exception as e:
            # Restore stdout
            sys.stdout = old_stdout

            # Get captured output
            output_text = self.output.getvalue()

            # report error back to GUI
            error_msg = f"ERROR: {str(e)}\n\n{output_text}"
            self.error_callback(error_msg)