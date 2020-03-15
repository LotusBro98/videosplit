import os
import sys
sys.path.insert(0, "venv/Lib/site-packages")

import ffmpeg
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def split_multithread(in_filename, out_filename, start, end):
    os.makedirs(os.path.dirname(out_filename), exist_ok=True)

    p = (
        ffmpeg
        .input(in_filename)
        .trim(start=start, end=end)
        .output(out_filename)
        .run_async(pipe_stdin=True, pipe_stdout=True, pipe_stderr=True)
    )

    print("{}({}s : {}s) ==> {}".format(in_filename, start, end, out_filename))
    p.communicate(input="y".encode())
    p.wait()

if __name__ == '__main__':

    root = Tk()
    root.withdraw()  # hide the window
    filename = askopenfilename(
        parent=root,
        title='Select Excel table with split defs',
        filetypes=[('Excel table', '.xlsx'), ('Excel table', '.xls')])
    root.destroy()

    dfs = pd.read_excel(filename)
    for index, row in dfs.iterrows():
        input, output, start, end = row["Input"], row["Output"], row["Start"], row["End"]
        split_multithread(input, output, start, end)