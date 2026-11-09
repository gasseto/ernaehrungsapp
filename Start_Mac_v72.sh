#!/bin/bash
cd "$(dirname "$0")"
echo "Starting Ernährungsberatungs-Suite v72..."
python3 -m streamlit run ernaehrungsberater_app_v72.py
