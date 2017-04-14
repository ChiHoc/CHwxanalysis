# CHwxanalysis

微信公众平台文章统计数据爬虫  

因为微信的文章详细统计只能选择7天的范围  

因此写了一个程序来获取指定日期范围的统计数据  

同时对数据进行格式化  

同时提供了Python2和Python3两个版本，和windows的打包版。  

（使用py2app和pyinstaller在mac上打包都失败了，尝试打包成功可以PR）  

## Usage

pip install requests xlwt pycookiecheat

**Windows用户注意**  

由于pycookiecheat原作者没有支持windows  

请到这下载支持win的版本：https://github.com/ChiHoc/pycookiecheat  

---

首先使用Chrome登录微信公众平台  

复制地址栏上的token  

**python2执行**  

python2 wxanalysis_py2.py  

**python3执行**  

python3 wxanalysis_py3.py

然后输入时间范围和token然后执行
