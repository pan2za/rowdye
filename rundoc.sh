# 启动 LibreOffice 并在后台监听 2002 端口
# --headless: 无界面模式
# --norestore: 不恢复上次会话
# --accept: 接受 socket 连接
# --invisible: 不可见 (可选，防止弹出窗口)

libreoffice --headless --norestore --invisible --accept="socket,host=localhost,port=2002;urp;" &

#MUST NOT USE conda python for python version collapsing
/usr/bin/python3 colorful.py /home/coconet/xunteng/input.docx /home/coconet/xunteng/output.docx

