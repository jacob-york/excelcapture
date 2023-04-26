"""ExcelCapture
Converts all Excel Spreadsheets in a given directory into PNGS, outputting them in a separate file.

excel2img module is free-to-use under the following License:

# -*- coding: utf-8 -*-
#  Copyright 2016 Alexey Gaydyukov
#
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.
"""

__author__ = "Jacob York"


import os

import excel2img


def capture(name, in_path, out_path):
    excel2img.export_img(f"{in_path}{name}.xlsx", f"{out_path}{name}.png", "Sheet1", None)


def clear_folder(folder):
    """Deletes every file in folder."""
    file_list = os.listdir(folder)
    for file in file_list:
        os.remove(f"{folder}{file}")


def get_paths_content() -> tuple[str, str]:
    """Gets path strings from Paths.cfg, properly formatted."""
    with open("Paths.cfg") as path_file:
        paths: str = path_file.read()
    if paths.endswith("\n"):
        paths = paths[:-1]

    read_from, write_to = paths.split(",")

    if write_to.startswith(" "):
        write_to = write_to[1:]

    if not (read_from.endswith("/") or read_from.endswith("\\")):
        read_from += "/" if read_from.find("\\") == -1 else "\\"
    if not (write_to.endswith("/") or write_to.endswith("\\")):
        write_to += "/" if write_to.find("\\") == -1 else "\\"

    return read_from, write_to


def main():
    IN, OUT = get_paths_content()

    print(f"Running ExcelCaptureApp.exe ({IN} -> {OUT}).")

    for path in IN, OUT:
        if not os.path.exists(path):
            print(f"Directory not Found: '{path}'.")
            return

    clear_folder(OUT)
    file_list = os.listdir(IN)

    for file in file_list:
        name = file[:file.index(".")]

        # Also ignores temporary hidden files that Excel creates when the user opens a spreadsheet.
        if file.endswith(".xlsx") and file[:2] != "~$":
            capture(name, IN, OUT)
        else:
            print(
                f"During runtime, the program detected a file in {IN} that was not an excel workbook and ignored it. "
                f"(File name: {file})\n..."
            )

    if not file_list:
        print("Input directory is empty; no action taken.")

    input("Press <ENTER> to close.")


if __name__ == "__main__":
    main()
