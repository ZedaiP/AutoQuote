@echo off
cd D:\pythonProject\Quotation\.venv\Scripts
activate
cd D:\pythonProject\Quotation
python -m nuitka main.py --standalone --enable-plugin=pyqt5 --plugin-enable=numpy --show-progress --lto=yes --disable-console
pause