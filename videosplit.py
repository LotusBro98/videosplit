import datetime
import glob
import os
import sys
sys.path.insert(0, "venv/Lib/site-packages")

import ffmpeg
import xlrd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def extract_part(in_filename, out_filename, start, end):
    os.makedirs(os.path.dirname(out_filename), exist_ok=True)

    input = ffmpeg.input(in_filename)
    video = (
        input
        .trim(start=start, end=end)
        .setpts("PTS-STARTPTS")
    )
    output = ffmpeg.output(video, out_filename)

    proc = output.run_async(pipe_stdin=True, pipe_stdout=True, pipe_stderr=True)

    print("{}({}s : {}s) ==> {}".format(in_filename, start, end, out_filename))
    proc.communicate(input="y".encode())
    proc.wait()

def parse_file(filename, in_dir="in", out_dir="out"):
    wb = xlrd.open_workbook(filename)
    sheet = wb.sheet_by_index(0)

    scenes = sheet.row_values(2, 4)
    scene_names = sheet.row_values(0, 4)[::2]
    n_scenes = len(scene_names)

    t0 = datetime.datetime(1899, 12, 31, 0, 0)
    scenes = [(
        scene_names[i],
        (xlrd.xldate_as_datetime(scenes[2*i], datemode=0) - t0).total_seconds() / 60,
        (xlrd.xldate_as_datetime(scenes[2*i+1], datemode=0) - t0).total_seconds() / 60
    ) for i in range(n_scenes)]

    starts = sheet.col_slice(3, 2)
    starts = [(xlrd.xldate_as_datetime(t.value, datemode=0) - t0).total_seconds() / 60 for t in starts]

    names = sheet.col_values(1, 2)

    splits = []

    for start_zero, name in zip(starts, names):
        name_in = glob.glob(os.path.join(in_dir, name + ".*"))
        if len(name_in) == 0:
            raise FileNotFoundError("Input video \"{}.*\" is not found in directory \"{}\"".format(name, in_dir))
        name_in = name_in[0]

        for scene_name, start, end in scenes:
            name_out = "{} - {}.mp4".format(name, scene_name)
            name_out = os.path.join(out_dir, name, name_out)
            start += start_zero
            end += start_zero

            splits.append((name_in, name_out, start, end))

    return splits


if __name__ == '__main__':

    root = Tk()
    root.withdraw()  # hide the window
    filename = askopenfilename(
        parent=root,
        title='Выберите файл с диапазонами для нарезки',
        filetypes=[('Excel table', '.xlsx'), ('Excel table', '.xls')])
    root.destroy()

    splits = parse_file(filename)

    for input, output, start, end in splits:
        extract_part(input, output, start, end)