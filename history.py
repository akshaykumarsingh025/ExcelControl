# history.py


class CommandHistory:
    def __init__(self):
        self.log = []  # List of (command, code, status)

    def add(self, command: str, code: str, success: bool):
        self.log.append({"command": command, "code": code, "success": success})

    def get_last_code(self):
        if self.log:
            return self.log[-1]["code"]
        return None

    def get_all(self):
        return self.log

    def clear(self):
        self.log = []
