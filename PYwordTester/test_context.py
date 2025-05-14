import io
import sys

class TestContext:
    """
    A context to execute Python code dynamically with specified libraries or global imports.
    """
    def __init__(self, DUT_SerialNumber="/", globals_to_import=None):
        """
        Initialize the test context with access to specified libraries or global imports.
        
        Args:
            DUT_SerialNumber (str): Serial number of the Device Under Test (optional).
            globals_to_import (dict): A dictionary of global variables to make available
                                      in the context, typically from globals().
        """
        self.DUT_SerialNumber = DUT_SerialNumber
        self.context = {
            "DUT_SerialNumber": DUT_SerialNumber,
        }

        # Add all provided global variables to the context
        if globals_to_import:
            self.context.update({
                name: obj for name, obj in globals_to_import.items()
                if not name.startswith("__")  # Ignore private/builtin variables
            })

    def execute(self, code):
        """
        Execute the provided Python code within the test context, capturing standard output (stdout)
        and returning the last expression's value for REPL-like behavior.

        Args:
            code (str): The Python code to execute as a string.

        Returns:
            tuple:
                - result (any): The value of the last executed expression or statement in the code.
                - exec_successful (bool): True if code executed without exceptions, False otherwise.
                - printed_output (str): The content printed to stdout during the code execution.
        """
        stdout_capture = io.StringIO()
        original_stdout = sys.stdout
        try:
            sys.stdout = stdout_capture  # Redirect stdout to capture printed output

            # Split code into lines to evaluate the last one
            lines = code.strip().splitlines()

            # Execute all lines except the last one
            for line in lines[:-1]:
                exec(line, {}, self.context)

            # Evaluate the last line explicitly to get the result
            last_line = lines[-1]
            try:
                result = eval(last_line, {}, self.context)
            except SyntaxError:
                # If last line is not an expression, execute it
                exec(last_line, {}, self.context)
                result = None

            exec_successful = True
        except Exception as e:
            result = f"Error: {e}"
            exec_successful = False
        finally:
            sys.stdout = original_stdout  # Restore original stdout

        return result, stdout_capture.getvalue(), exec_successful

    def destroy(self):
        """Clean up the context."""
        self.context.clear()
        print("Test context destroyed.")
