怎么运行和打包？

运行：直接运行这段 Python 代码即可。

打包为 EXE：

    打开命令行，输入：
        pip install pyinstaller

    输入：
        pyinstaller --onefile --noconsole --add-data "C:/Windows/Fonts/simsun.ttc;." 程序名.py

这样打包出来的 EXE 就可以分发给其他人直接使用了。