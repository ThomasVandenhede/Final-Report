# Python code that
# - creates a new INM study from a reference INM study for an FDR file
# - names the new study after the FDR file
# - replaces the .dbf input files located in the new INM study with the .dbf
# input files from the flight folder (eg CS-TNN_2016-01-05_TP842).
#
# NB: INM Studies need to be in the same directory as where the code is run.

import glob
import os
import shutil
from distutils.dir_util import copy_tree

__author__ = 'Thomas Vandenhede'


def get_immediate_subdirectories(a_dir):
    """
    Returns a list of all immediate subdirectories of a folder.

    :param a_dir:
    :return:
    """
    result = [name for name in os.listdir(a_dir)
              if os.path.isdir(os.path.join(a_dir, name))]
    return result


def create_inm_study_directories(data_path, studies_path, dir_list):
    """
    Creates INM study directories with all the required files to perform a
    complete INM study.

    :param studies_path: the folder where all study directories will be created
    :param dir_list: list of names of all directories to be created
    :return:
    """
    print("Creating study directories...")
    for d in dir_list:
        source_dir = os.path.join(studies_path, 'Reference')
        dest_dir = os.path.join(studies_path, d)
        data_dir = os.path.join(data_path, d)

        # create study directory if doesn't already exists
        if not os.path.exists(dest_dir):
            os.makedirs(dest_dir)

        # copy Reference study to new study
        copy_tree(source_dir, dest_dir)

        # copy .dbf files to new study
        files = glob.iglob(os.path.join(data_dir, '*.dbf'))
        for f in files:
            if os.path.isfile(f):
                shutil.copy2(f, dest_dir)

        print("Study directory created for %s" % d)


def main():
    # Set path to input data and INM study folders (.dbf files)
    data_path = os.path.join('INM Files', 'MCDP Flight Trials')
    studies_path = 'INM Studies'

    # Get
    dir_list = get_immediate_subdirectories(data_path)
    create_inm_study_directories(data_path, studies_path, dir_list)


if __name__ == '__main__':
    main()
