# CHwxanalysis

[![Build Status](https://travis-ci.org/ChiHoc/CHwxanalysis.svg?branch=master)](https://travis-ci.org/ChiHoc/CHwxanalysis)

微信公众平台文章统计数据爬虫  

因为微信的文章详细统计只能选择7天的范围  

因此写了一个程序来获取指定日期范围的统计数据，并对数据进行格式化  

同时提供了Python2和Python3两个版本，和windows的打包版。  

## Install

**Windows用户注意**  

由于pycookiecheat原作者没有支持windows  

请到这下载支持win的版本：https://github.com/ChiHoc/pycookiecheat  

**事前准备：**

`pip install requests xlwt pycookiecheat pyinstaller`

**Windows打包**

`pyinstaller -F -w wxanalysis.py`

**Mac打包**

`pyinstaller -w -F --hidden-import appdirs --hidden-import packaging.requirements --hidden-import packaging.specifiers --hidden-import packaging.version --hidden-import packaging wxanalysis_py2.py`

## Command Line

**python2执行：**  

`python2 wxanalysis_py2.py`  

**python3执行：**  

`python3 wxanalysis_py3.py`

## Usage

打开程序，输入要统计的日期范围 

使用**Chrome浏览器**登录微信公众平台  

复制地址栏上的token到程序token输入框，然后运行

excel文件将会输出到程序文件夹
