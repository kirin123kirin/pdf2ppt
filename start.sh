#!/bin/sh
python3 -m http.server 8000 &
sleep 0.3
open "http://localhost:8000/pdf2ppt.html" 2>/dev/null \
  || xdg-open "http://localhost:8000/pdf2ppt.html" 2>/dev/null \
  || echo "ブラウザで http://localhost:8000/pdf2ppt.html を開いてください"
wait
