# AutoMaster
支持WIN10的自动操作

## 效果
1.启动效果
页面实现录制脚本、启动脚本、热键设置、循环设置、操作提示等功能
![请添加图片描述](https://img-blog.csdnimg.cn/29e0561994204296b0e2706542196282.gif)

2.录制效果
可以录制鼠标操作、键盘输入、组合键（复制黏贴）。
![请添加图片描述](https://img-blog.csdnimg.cn/95b1c8c505634c46b858940e394d19fa.gif)


3.执行效果
点击启动，则按照脚本重复执行刚刚的操作
![请添加图片描述](https://img-blog.csdnimg.cn/167e54f0589b4385b5c409627888e9e0.gif)

[立即下载免安装绿色版](https://download.csdn.net/download/weixin_42236404/23384988)


## 代码说明
开发环境：win10
开发语言：python3.8
开发IDE：pycharm
使用包：

| wxPython | 4.1.1 |

| pyWinhook| 1.6.2 |

| pywin32| 301|

pyWinhook如果电脑没有C++编译工具的话，会安装失败，可以通过WHL文件进行安装，详见[安装whl](https://editor.csdn.net/md/?articleId=120347684)

注释都有，自行阅读

### 使用指南
1.安装依赖
`pip install requestments.txt`

2.下载pyWinhook
按照这个走
https://blog.csdn.net/weixin_42236404/article/details/120347684?spm=1001.2014.3001.5501

3.运行即可
