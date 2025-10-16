import tkinter as tk
from tkinter import filedialog
import pandas as pd
import time

from Funktioner.Hent_dataframe import import_deltagerliste, import_postoversigt
from Funktioner.Lav_pdf import lav_pdf
from Funktioner.Hent_pdf import hent_pdf