import os
import sys
sys.path.insert(0, "venv/Lib/site-packages")

import ffmpeg
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def split_multithread(in_filename, out_filename, n_splits):
    os.makedirs(os.path.dirname(out_filename), exist_ok=True)

    probe = ffmpeg.probe(in_filename)
    duration = float(probe["format"]["duration"])

    for i in range(0, n_splits):
        start = i * duration / n_splits
        end = (i + 1) * duration / n_splits

        p = (
            ffmpeg
            .input(in_filename)
            .trim(start=start, end=end)
            .output(out_filename.format(i))
            .run_async(pipe_stdin=True)
        )

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
        input, output, n_splits = row["Input"], row["Output"], row["N splits"]
        split_multithread(in_filename=input, out_filename=output, n_splits=n_splits)