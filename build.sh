#!/bin/bash
pip install -r requirements.txt
mkdir -p ~/.fonts
curl -L "https://github.com/anthonyfok/fonts-wqy-zenhei/raw/master/wqy-zenhei.ttc" -o ~/.fonts/wqy-zenhei.ttc 2>/dev/null || true
fc-cache -fv 2>/dev/null || true
