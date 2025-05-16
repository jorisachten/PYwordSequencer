import sys
import io

class TeeIO(io.StringIO):
    def __init__(self, *streams):
        super().__init__()
        self.streams = streams

    def write(self, s):
        for stream in self.streams:
            stream.write(s)
            stream.flush()  # Make sure output appears immediately
        return len(s)

    def flush(self):
        for stream in self.streams:
            stream.flush()

class TestContext:
    def __init__(self, DUT_SerialNumber="/", globals_to_import=None, debug=False):
        self.DUT_SerialNumber = DUT_SerialNumber
        self.debug = debug
        self.context = {
            "DUT_SerialNumber": DUT_SerialNumber,
        }

        if globals_to_import:
            self.context.update({
                name: obj for name, obj in globals_to_import.items()
                if not name.startswith("__")
            })

    def execute(self, code):
        stdout_capture = io.StringIO()
        if self.debug:
            tee_stdout = TeeIO(stdout_capture, sys.__stdout__)
        else:
            tee_stdout = stdout_capture

        original_stdout = sys.stdout
        sys.stdout = tee_stdout

        try:
            lines = code.strip().splitlines()
            for line in lines[:-1]:
                exec(line, {}, self.context)

            last_line = lines[-1]
            try:
                result = eval(last_line, {}, self.context)
            except SyntaxError:
                exec(last_line, {}, self.context)
                result = None

            exec_successful = True
        except Exception as e:
            result = f"Error: {e}"
            exec_successful = False
        finally:
            sys.stdout = original_stdout

        printed_output = stdout_capture.getvalue()
        return result, printed_output, exec_successful


    def destroy(self):
        """Clean up the context."""
        self.context.clear()
        print("Test context destroyed.")
