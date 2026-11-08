import importlib, sys
import os; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from schema import LESSONS
for m in ['d01','d02','d03','d04','d05','d06','d07']:
    importlib.import_module(m)
LESSONS.sort(key=lambda l: l['n'])
