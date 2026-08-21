#!/bin/bash
PORT=5175
DIR="/Users/apple/Desktop/日报文件夹"

# Check if static server is running
if ! python3 -c "import socket; s=socket.socket(); s.settimeout(2); s.connect(('127.0.0.1',$PORT)); s.close(); print('ok')" 2>/dev/null | grep -q ok; then
    kill $(lsof -ti:$PORT 2>/dev/null) 2>/dev/null
    kill $(lsof -ti:5176 2>/dev/null) 2>/dev/null
    sleep 1
    cd "$DIR" && nohup python3 start_services.py > /tmp/yinshan-dashboard.log 2>&1 &
fi
