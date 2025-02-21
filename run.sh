#!/bin/bash

FILENAME="WordFeud.xlsx"

# Check if file exists before downloading
if [ -f "$FILENAME" ]; then
  echo "File exists, skipping download. Remove $FILENAME to redownload."
else
  echo "File does not exist, downloading"
  curl -L https://docs.google.com/spreadsheets/d/1lsB-p2w0te_Ui4QYDLp49wB-_7vYrkSPVCt4fhkjz-I/export\?format\=xlsx -o "$FILENAME"
fi
python convert_to_matches.py
rayter wordfeudligan.txt