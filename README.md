# WordChecker

主要函数
 1. FontsChecker(filePath) 按照样式以修订模式检查字体
 2. CaptionNumCheckerByFeature(filePath) 对编号错误图表进行定位
 3. int GetSectionCount(filePath) 获取节的数量
 4. XmlFormatter.FormatAndSaveXml() 在窗口中输入xml内容，会进行格式化输出，便于观察，**注意在 `XmlFormatter.cs` 中设置输出文件**

可能出现的问题
 1. 以修订模式检查字体时，请确保文档处于无修订状态，否则会导致修订id出现重复
    ~~不过暂时没有发现这个会导致什么后果，还是能够在word中正常使用~~
