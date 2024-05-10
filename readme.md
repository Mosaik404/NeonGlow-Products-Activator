# 简介

一种使用内嵌浏览器和服务器端文件路径进行软件授权的实现方法。<br>
整体使用 Visual Basic 6.0 写成 (~~有点老，但是能用~~)<br>

# 思路

客户端：输入认证代码 - 合成URL - 访问网页 - 下载文件<br>
服务端：建立认证代码对应路径 - 应答请求<br>

举例：<br>
认证代码：0x - 114514 - 1919810<br>
产品代码：sodayo  *(用于区分不同产品，即一个激活器可用于多个产品)*<br>
合成URL：https://ql.example.com/sodayo/0x-114514-1919810/index.html<br>
点击“授权”，软件内嵌浏览器访问对应URL，出现下载页面。<br>
【注意】使用内嵌浏览器是防止用户看到具体域名和URL，也可以直接使用系统默认浏览器打开对应URL。(o゜▽゜)o☆<br>

# 实现

1.设定认证代码，部署服务器端文件。<br>
2.公布“认证助手”应用程序下载链接，例如：[下载产品认证助手 (mosaik404.github.io)](https://mosaik404.github.io/products-quali/demo/download/quali/qualiapp.html)。<br>
3.随产品发放认证代码，例如：0x - 114514 - 1919810。<br>

4.用户在“认证助手”应用程序内输入认证代码，跳转至对应网页下载产品。<br>
