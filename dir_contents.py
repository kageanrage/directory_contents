import os, logging, openpyxl
from config import Config
from pprint import pprint

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')  # turns on logging

cfg = Config()

direc = cfg.direc


os.chdir(direc)

dict_list = []  # create empty list of per-movie dictionaries
for folderName, subfolders, filenames in os.walk(direc):
    for filename in filenames:
        movie_dict = {}  # create empty per-movie dictionary
        fname_inc_path = folderName + "\\" + filename  # generate the full filename path for use in getsize method
        size_in_bytes = os.path.getsize(fname_inc_path)
        if size_in_bytes > 50000000:  # if it's large enough to be a movie file
            size_in_gb = round(size_in_bytes / 1073741824, 2)  # express size in GB, rounded to 2 d.p.
            # logging.debug(f"filename = {filename}")
            movie_dict.setdefault('name', filename)  # add movie name to the per-movie dict
            # logging.debug(f"file size = {size_in_gb} GB")
            movie_dict.setdefault('size', size_in_gb)  # add movie size to the per-movie dict
            dict_list.append(movie_dict)  # add the per-movie dict to the dict_list


os.chdir(cfg.repo_dir)  # change the current working dir
xls_name = 'Movies.xlsx'  # define xlsx filename
wb = openpyxl.Workbook()  # instantiate a new excel workbook
sheet = wb.active  # define the first sheet of the workbook to be the one to work with

sheet.cell(1, 1).value = "Movie"  # set heading cell content
sheet.cell(1, 2).value = "Size (GB)"  # set heading cell content

# cycle through the movies and insert them into the excel file
for num, movie_dict in enumerate(dict_list):
    # print(f"num = {num}, name = {movie_dict['name']}, size = {movie_dict['size']}")
    row_of_interest = num + 2  # offset by 2, as enumerate defaults to start from zero and I didn't change it
    first_cell = sheet.cell(row=row_of_interest, column=1)  # select the first cell in the row
    first_cell_value = movie_dict['name']  # grab the name of the movie from the movie_dict
    first_cell.value = first_cell_value  # assign that value to the cell
    second_cell = sheet.cell(row=row_of_interest, column=2)
    second_cell_value = movie_dict['size']
    second_cell.value = second_cell_value

wb.save(xls_name)  # save the excel workbook
