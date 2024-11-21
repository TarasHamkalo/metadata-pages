import json
import os
import subprocess
from pathlib import Path


class SimpleExifTool(object):
    sentinel = "{ready}\n"

    # windows_sentinel = "{ready}\r\n"

    def __init__(self, executable="/usr/bin/exiftool"):
        self.executable = executable

    def __enter__(self):
        self.process = subprocess.Popen(
            [self.executable, "-stay_open", "True", "-@", "-"],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE
        )
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.process.stdin.write("-stay_open\nFalse\n".encode())
        self.process.stdin.flush()

    def executable_exists(self):
        return Path(self.executable).exists()

    def execute(self, *args):
        args = args + ("-execute\n",)
        args = str.join("\n", args)
        self.process.stdin.write(args.encode())
        self.process.stdin.flush()
        output = b""
        fd = self.process.stdout.fileno()
        while not output.endswith(self.sentinel.encode()):
            output += os.read(fd, 4096)
        return output.decode()[:-len(self.sentinel)]

    def get_metadata(self, path: str):
        a = self.execute("-G1", "-j", "-n", path)
        return json.loads(a)
