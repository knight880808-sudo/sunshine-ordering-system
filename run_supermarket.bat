@echo off
:: 切换到 C 盘，防止路径跨盘符报错
c:
:: 进入项目目录
cd C:\Users\Administrator\Desktop\Supermarket
:: 启动 Streamlit
streamlit run app.py --server.port 8502 --server.address 0.0.0.0
pause