# -*- coding: utf-8 -*-
"""
make-starters.py - rebuild every PowerPoint practice presentation.

Run from the powerpoint/ folder:  python3 tools/make-starters.py
Needs python-pptx, lxml, Pillow and python-docx (pip install python-pptx lxml pillow python-docx),
and ffmpeg for the practice video and music. Each week's files are built by
tools/pp_w1.py ... pp_w5.py; the same builders make the answer keys in
tools/test-docs (see make-test-docs.py).
"""
import os, sys, importlib
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
STARTERS = os.path.join(HERE, "..", "starters")

for week in range(1, 6):
    name = "pp_w%d" % week
    if not os.path.exists(os.path.join(HERE, name + ".py")):
        continue
    mod = importlib.import_module(name)
    mod.build_all(STARTERS)
    print("Built week", week)
