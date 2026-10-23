# -*- coding: utf-8 -*-
"""
make-starters.py - rebuild every Excel practice workbook.

Run from the excel/ folder:  python3 tools/make-starters.py
Needs XlsxWriter and lxml (pip install xlsxwriter lxml).
Each week's workbooks are built by tools/xl_w1.py ... xl_w5.py; the same
builders make the answer keys in tools/test-docs (see make-test-docs.py).
"""
import os, sys, importlib
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
STARTERS = os.path.join(HERE, "..", "starters")

for week in range(1, 6):
    name = "xl_w%d" % week
    if not os.path.exists(os.path.join(HERE, name + ".py")):
        continue
    mod = importlib.import_module(name)
    mod.build_all(STARTERS)
    print("Built week", week)
