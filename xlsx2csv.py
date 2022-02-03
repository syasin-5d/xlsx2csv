import os
import sys
import csv
import openpyxl
from natsort import natsorted


def main():
    source_dir = sys.argv[1]
    dist_dir = sys.argv[2]

    if not os.path.exists(dist_dir):
        os.makedirs(dist_dir)


# get xlsx files order in natural sorted order
    files = natsorted(os.listdir(source_dir))

    for filename in files:
        filepath = os.path.join(source_dir, filename)

        wb = openpyxl.load_workbook(filepath)
        ws_name = wb.sheetnames[0]
        ws = wb[ws_name]

        savecsv_path = os.path.join(dist_dir,
                                    filename.rstrip(".xlsx") + ".csv")
        with open(savecsv_path, 'w', newline="", encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            for row in ws.rows:
                writer.writerow([cell.value for cell in row])

if __name__ == "__main__":
    main()
