# Excel Capture


### Paths.cfg:
Paths.cfg stores your input and output paths, ordered /input/, /output/.
Default paths are provided in the project's directory, but if you change these (using absolute paths), ExcelCapture will update where it reads/writes from accordingly.

### Misc Note:
Don't make any files with names that begin with `~$`, otherwise ExcelCapture will ignore them.
This is because when you open an excel (e.g. `fileName.xlsx`) Excel makes a hidden file in the same directory named `~$fileName.xlsx` and removes it when you close `fileName.xlsx`.
I wrote ExcelCapture to ignore these files. However, if for some strange reason, you decide to start a file with `~$`, It'll be ignored.
