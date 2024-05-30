# excel_to_markdown

# 前言
自从我把博客从wordpress转到hexo上就踩了很多坑，其中就包括一个很棘手的问题，就是Markdown语法的表格不支持单元格合并，即使网上有很多excel转Markdown的在线工具可以用，但是都不支持单元格合并。于是我就自己用python写了简单的转换工具出来，希望可以帮助到各位。

# 特点
 - 支持单元格合并

# 使用方法

1. 先安装py库

        pip install pandas

        pip install openpyxl


2. 直接双击py脚本文件，然后会弹出一个cmd窗口，直接把excel表格文件拖进cmd窗口，然后回车即可。完成后会在脚本同级目录下生成一个名为`output.md`的文件