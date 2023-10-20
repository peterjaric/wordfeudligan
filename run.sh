#!/bin/bash
curl -L https://docs.google.com/spreadsheets/d/1lsB-p2w0te_Ui4QYDLp49wB-_7vYrkSPVCt4fhkjz-I/export\?format\=xlsx -o WordFeud.xlsx
python convert_to_matches.py
rayter wordfeudligan.txt
