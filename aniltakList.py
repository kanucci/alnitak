import os, os.path
import pandas as pd
#from openpyx import load_worbook


def xl_List():

    DIR = 'C:/Users/RESIDENCIA/Documents/Dataset/'
    plan_list = [DIR + name for name in os.listdir(DIR) if os.path.isfile(os.path.join(DIR, name)) and name.split(".")[1] == "xlsx"]

    return plan_list
