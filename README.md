这个是市面上最成功的免费 PPT/PPTX 转 图片的技术方案。 解决了 apache poi 工具 转换失真严重的问题 ，主要依赖 LibreOffice 开源包+ pdfbox 方案。
LibreOffice 支持 Windows/Mac/Linux各个操作系统，需要自行安装。
pdfbox 使用需要注入依赖 ,参考 pom.xml,不然会报错。
该方案的主要思路是 根据 LibreOffice 工具，先把 ppt pptx 文件 转成 pdf 文件，再根据 pdfbox 把 pdf 转成 图片
