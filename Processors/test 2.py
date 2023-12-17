#Einlesen Module
import pandas as pd
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog

##Pfad bestimmen von File
base_file_path_1UG_ver = "/Users/leameier/Desktop/Programmierung 1/0_Mein Projekt/Vorlagen/CSV neu/Wände/02_Grundriss 1.UG_ver.csv"
#base_file_path_0EG_ver = "/Users/leameier/Desktop/Programmierung 1/0_Mein Projekt/Vorlagen/CSV neu/Wände/03_Grundriss 0.EG_ver.csv"
#base_file_path_1OG_ver = "/Users/leameier/Desktop/Programmierung 1/0_Mein Projekt/Vorlagen/CSV neu/Wände/04_Grundriss 1.OG_ver.csv"


base_file_path_2UG_hor = "/Users/leameier/Desktop/Programmierung 1/0_Mein Projekt/Vorlagen/CSV neu/Decken/01_Grundriss 2.UG_hor.csv"
#base_file_path_1UG_hor = "/Users/leameier/Desktop/Programmierung 1/0_Mein Projekt/Vorlagen/CSV neu/Decken/02_Grundriss 1.UG_hor.csv"


##CSV File einlesen
base_df_1UG_ver = pd.read_csv(base_file_path_1UG_ver,sep=";")
#base_df_0EG_ver = pd.read_csv(base_file_path_0EG_ver,sep=";")
#base_df_1OG_ver = pd.read_csv(base_file_path_1OG_ver,sep=";")


base_df_2UG_hor = pd.read_csv(base_file_path_2UG_hor,sep=";")
#base_df_1UG_hor = pd.read_csv(base_file_path_1UG_hor,sep=";")

##Eingelesene CSV zeigen
print("Ergebnis eingelesenes CSV 1UG_ver: {}".format(str(base_df_1UG_ver)))
#print("Ergebnis eingelesenes CSV 0EG_ver: {}".format(str(base_df_0EG_ver)))
#print("Ergebnis eingelesenes CSV 1OG_ver: {}".format(str(base_df_1OG_ver)))