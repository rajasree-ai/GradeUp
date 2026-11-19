#!/usr/bin/env bash
set -e

apt-get update
apt-get install -y --no-install-recommends tesseract-ocr tesseract-ocr-eng
python -m pip install -r requirements.txt
