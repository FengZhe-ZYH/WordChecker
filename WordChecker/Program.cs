using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Xml.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2016.Presentation.Command;
using System.Text.RegularExpressions;
using System.Security.Policy;
using DocumentFormat.OpenXml.Bibliography;
class Program
{
    static void Main()
    {
        string filePath = @"C:\Users\FengZhe\Desktop\Study\Gproject\WordChecker\assest\RevisionTest.docx";

        FontsChecker(filePath);//按照样式以题注检查字体

        //Console.WriteLine(GetSectionCount(filePath));//获取节的数量


        //CaptionNumCheckerByFeature(filePath);


        //DePaintTextHelper(filePath);

        //XmlFormatter.FormatAndSaveXml();
    }


    static int GetSectionCount(string filePath) 
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false)) 
        { 
            var body = wordDoc.MainDocumentPart.Document.Body; 
            var sections = body.Descendants<SectionProperties>();
            return sections.Count(); 
        } 
    }
    static string ExtractTextByPara(Paragraph paragraph)
    {
        string text = "";
        if(paragraph == null)
        {
            return text;
        }
        else
        {
            foreach (Run run in paragraph.Descendants<Run>())
            {
                text += run.InnerText;
            }
            return text;
        }
    }
    static void CaptionNumCheckerByFeature(string filePath)
    {
        void handle(string DocNum, int CounterNum, Paragraph paragraph)
        {
            string CounterNum_string = CounterNum.ToString();
            if(DocNum != CounterNum_string)
            {
                //暂时高亮
                Run firstRun = paragraph.GetFirstChild<Run>();
                if(firstRun.RunProperties == null)
                {
                    firstRun.RunProperties = new()
                    {
                        Highlight = new() { Val = HighlightColorValues.Red }
                    };
                }
                else
                {
                    firstRun.RunProperties.Highlight = new()
                    {
                        Val = HighlightColorValues.Red
                    };
                }
            }
        }
        using(WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            Body body = mainPart.Document.Body;
            int TableNum = 0, GraphNum = 0;//全局计数器

            foreach(Paragraph paragraph in body.Descendants<Paragraph>())
            {
                //先进行样式提取和判断
                ParagraphProperties paragraphProperties = paragraph.ParagraphProperties;
                ParagraphStyleId paragraphStyleId = null;
                bool notCaption = true;

                string paraStyleName = null;
                if (paragraphProperties != null)
                {
                    paragraphStyleId = paragraphProperties.ParagraphStyleId;
                    if(paragraphStyleId != null)
                    {
                        string styleId = paragraphStyleId.Val;
                        paraStyleName = getNameByStyle(getStyleById(mainPart, styleId));
                    }
                }
                //之后，进入到para的run中
                Run[] runs = paragraph.Descendants<Run>().ToArray();
                foreach (Run run in runs)
                {
                    //首先，获取run的样式，同样，可能为空，此时返回null即可
                    RunProperties runProperties = run.RunProperties;
                    RunStyle runStyle = null;
                    string runStyleName = null;

                    if (runProperties != null)
                    {
                        runStyle = runProperties.RunStyle;

                        if (runStyle != null)
                        {
                            string runStyleId = runStyle.Val;
                            runStyleName = getNameByStyle(getStyleById(mainPart, runStyleId));
                        }
                    }
                    //这里已经获取了一个区块的样式链，现在进行判断
                    string styleName = null;
                    if(runStyleName == null)
                    {
                        styleName = paraStyleName;
                    }
                    else
                    {
                        styleName = runStyleName;
                    }
                    if(styleName == "题注")
                    {
                        notCaption = false;
                        break;
                    }
                    
                }

                //若有Run的样式为题注，我们进行字符特征判定
                if (false)
                {
                    break;
                }
                else
                {
                    string paraText = ExtractTextByPara(paragraph);
                    string firstNonSpaceChar = paraText.TrimStart().FirstOrDefault().ToString();
                    if (firstNonSpaceChar == "图")
                    {   //此处，首字为图，我们只匹配第一个，所以用正则表达式，要么提取到开头，要么不进行提取
                        string pattern = @"图(\d+)";
                        Match match = Regex.Match(paraText, pattern);
                        if (match.Success)
                        {
                            string num = match.Groups[1].Value;
                            GraphNum++;
                            handle(num, GraphNum, paragraph);
                        }
                    }
                    else if(firstNonSpaceChar == "表")
                    {
                        string pattern = @"表(\d+)";
                        Match match = Regex.Match(paraText, pattern);
                        if (match.Success)
                        {
                            string num = match.Groups[1].Value;
                            TableNum++;
                            handle(num, TableNum, paragraph);
                        }
                    }
                }
            }
        }
    }
    static void CaptionNumChecker(string filePath)
    {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {   
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

            // 字典用于存储每个序列名称和对应的编号列表
            Dictionary<string, int> seqNumbers = new Dictionary<string, int>();
                foreach (Paragraph paragraph in mainPart.Document.Body.Descendants<Paragraph>())
                {
                bool hasSEQ = false;
                string seqName = null;

                foreach (Run run in paragraph.Descendants<Run>())
                {
                    FieldCode code1 =  run.Descendants<FieldCode>().FirstOrDefault();
                    string code = null;
                    if (code1 != null){
                        code = code1.InnerText;
                    }
                    

                    if (hasSEQ)
                    {
                        if (seqNumbers.ContainsKey(code))
                        {
                            seqName = code;
                            seqNumbers[code] += 1;
                            hasSEQ = false;
                        }
                        else
                        {
                            seqName = code;
                            seqNumbers.Add(code, 1);
                            hasSEQ = false;
                        }
                    }

                    if (code != null && code.Contains("SEQ"))
                    {
                        hasSEQ = true;
                    }

                    
                    if(run.RunProperties != null && run.RunProperties.NoProof != null)
                    {
                        if (run.InnerText != null && seqName != null && run.InnerText.ToString() != seqNumbers[seqName].ToString())
                        {
                            run.RunProperties.Highlight = new Highlight() { Val = HighlightColorValues.Red };
                        }
                        
                    }

                }
                }
            mainPart.Document.Save();
            }
    }
    static void DePaintTextHelper(string filePath)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            Paragraph[] paragraphs = mainPart.Document.Body.Descendants<Paragraph>().ToArray();
            foreach (var paragraph in paragraphs)
            {
                Run[] runs = paragraph.Descendants<Run>().ToArray();
                foreach (Run run in runs)
                {
                    DePaintText(run);
                }
            }
            mainPart.Document.Save();
        }
    }
    static void PageNumChecker()
    {
        string filepath = @"C:\Users\FengZhe\Desktop\Study\Gproject\WordChecker\assest\21376336-周亚辉.docx";

        int expectedPageNumber = 1;


        using (WordprocessingDocument doc = WordprocessingDocument.Open(filepath, true))
        {
            foreach (var footerPart in doc.MainDocumentPart.FooterParts)
            {
                foreach(Paragraph paragraph in footerPart.Footer.Descendants<Paragraph>())
                {
                    foreach (Run run in paragraph.Descendants<Run>())
                    {
                        var fieldChar = run.Descendants<FieldChar>().FirstOrDefault(fc => fc.FieldCharType == FieldCharValues.Begin);
                        
                    }
                }
            }
        }
    }
   
    static int RevisionCounter = 0;
    static void FontsChecker(string filePath)
    {
        //处理函数
        
        void Handle(string font, Run run)
        {
            if (run.RunProperties == null)
            {
                run.RunProperties = new RunProperties();
            }

            if(font != null)//若字体不为空，说明样式和区块产生了字体冲突
            {
                RunProperties runProperties = run.RunProperties;

                //接下来我们使用修订模式删除这个run区块内的字体
                string author = Environment.UserName;
                DateTime date = DateTime.Now;

                var newRunProperties = (RunProperties)runProperties.CloneNode(true);
                newRunProperties.RunFonts = null;

                PreviousRunProperties previousRunProperties = new PreviousRunProperties(runProperties.CloneNode(true));

                RunPropertiesChange revision = new()
                {
                    Author = author,
                    Id = RevisionCounter.ToString(),
                    Date = date,
                    DateUtc = DateTime.UtcNow,
                    PreviousRunProperties = previousRunProperties
                };

                run.RunProperties = newRunProperties;
                run.RunProperties.AddChild(revision);
               
                RevisionCounter++;
            }
        }

        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            Dictionary<string, string> FontsDict = new Dictionary<string, string>();
            FontsDict.Add("majorEastAsia", null);
            FontsDict.Add("minorEastAsia", null);
            FontsDict.Add("majorHAnsi", null);
            FontsDict.Add("minorHAnsi", null);

            FontsDict = getFontsDict(mainPart, FontsDict);

            //根据段落id查询段落样式 并结合 FontsDict 返回英文和中文字体信息

            //首先，先从文档的默认格式开始
            RunFonts runFontsOfDocDefault = getDocDefaultRunFonts(mainPart);

            //然后，获取默认段落的默认样式的RunFonts
            Style styleOfDefaultPara = getStyleById(mainPart, "a");
            RunFonts runFontsOfDefaultParaStyle = getFontsByStyle(styleOfDefaultPara);

            //之后开始进入文章
            Paragraph[] paragraphs = mainPart.Document.Body.Descendants<Paragraph>().ToArray();
            foreach(Paragraph paragraph in paragraphs)
            {
                //首先，获取段落样式，注意，段落样式可能为空，此时返回null即可
                ParagraphProperties paragraphProperties = paragraph.ParagraphProperties;
                ParagraphStyleId paragraphStyleId = null;
                RunFonts runFontsOfParaStyle = null;

                if (paragraphProperties != null)
                {
                    paragraphStyleId = paragraphProperties.ParagraphStyleId;

                    if (paragraphStyleId != null)
                    {
                        string styleId = paragraphStyleId.Val;
                        runFontsOfParaStyle = getFontsByStyle(getStyleById(mainPart, styleId));
                    }
                }
                //之后，进入到para的run中
                Run[] runs = paragraph.Descendants<Run>().ToArray();
                foreach (Run run in runs)
                {
                    //首先，获取run的样式，同样，可能为空，此时返回null即可
                    RunFonts runFontsOfRunStyle = null;
                    RunProperties runProperties = run.RunProperties;
                    RunStyle runStyle = null;

                    if (runProperties != null)
                    {
                        runStyle = runProperties.RunStyle;

                        if (runStyle != null)
                        {
                            string runStyleId = runStyle.Val;
                            runFontsOfRunStyle = getFontsByStyle(getStyleById(mainPart, runStyleId));
                        }
                    }
                    //最后，检查run中是否有明确的字体设置，如果有，返回设置的字体，如果没有，返回null
                    RunFonts runFontsOfRun = null;
                    if (runProperties != null)
                    {
                        runFontsOfRun = runProperties.RunFonts;
                    }

                    //通过五个rFonts，判断中文是什么字体
                    Dictionary<string, RunFonts> allRunFonts = new Dictionary<string, RunFonts>();
                    allRunFonts.Add("runFontsOfDocDefault", runFontsOfDocDefault);
                    allRunFonts.Add("runFontsOfDefaultParaStyle", runFontsOfDefaultParaStyle);
                    allRunFonts.Add("runFontsOfParaStyle", runFontsOfParaStyle);
                    allRunFonts.Add("runFontsOfRunStyle", runFontsOfRunStyle);
                    allRunFonts.Add("runFontsOfRun", runFontsOfRun);

                    string ChineseFont = null;

                    //ChineseFont = JudgeFontsByRFonts(FontsDict, allRunFonts);

                    ChineseFont = FontCheckerByStyleHelper(FontsDict, allRunFonts);

                    //处理函数，用于判断出后进行善后
                    Handle(ChineseFont, run);


                    //Console.WriteLine($"{run.InnerText}");
                    //Console.WriteLine("其中文字体为：" + ChineseFont);

                    //PaintText(run, ChineseFont);

                    //DePaintText(run);
                }
            }
            mainPart.Document.Save();
        }
    }
    static Dictionary<string, string> Dict_RunFonts2StringAndTransform(Dictionary<string, RunFonts> allRunFonts, Dictionary<string, string> FontsDict)
    {
        Dictionary<string, string> allRunFonts_string = new();
        //首先，我们要将字体字典进行转换
        foreach (KeyValuePair<string, RunFonts> kvp in allRunFonts)
        {
            if(kvp.Value != null)
            {
                string fonts = ExtractFontsByRunFonts(kvp.Value);
                allRunFonts_string.Add(kvp.Key, fonts);
            }
            else
            {
                allRunFonts_string.Add(kvp.Key, null);
            }
        }

        Dictionary<string, string> allRunFonts_string_2 = new();

        //其次，我们将所有的eastAsia进行查字典更换
        foreach (KeyValuePair<string, string> kvp in allRunFonts_string)
        {
            allRunFonts_string_2.Add(kvp.Key, kvp.Value);

            if (kvp.Value != null && FontsDict.ContainsKey(kvp.Value))
            {
                allRunFonts_string_2[kvp.Key] = FontsDict[kvp.Value];
            }
        }
        return allRunFonts_string_2;

    }
    static string FontCheckerByStyleHelper(Dictionary<string, string> FontsDict, Dictionary<string, RunFonts> allRunFonts)
    {
        //首先，我们将字体字典转换为查字典替换后的字符串字典
        Dictionary<string, string> allRunFonts_string = Dict_RunFonts2StringAndTransform(allRunFonts, FontsDict);

        //之后，我们判断这段字体需不需要进行更换
        if (allRunFonts_string["runFontsOfRun"] == null)
        {
            return null;
        }
        else
        {
            //查找runstyle和pstyle值是否为空
            string StyleFont = null;
            if (allRunFonts_string["runFontsOfRunStyle"] != null)
            {
                StyleFont = allRunFonts_string["runFontsOfRunStyle"];
            }
            else
            {
                StyleFont = allRunFonts_string["runFontsOfParaStyle"];
            }

            //最终判断
            if(StyleFont == allRunFonts_string["runFontsOfRun"])
            {
                return null;
            }
            else
            { 
                return "StyleFont";
            }
        }
    }
    static void DePaintText(Run run)
    {
        if (run.RunProperties != null)
        {
            run.RunProperties.Highlight = null;
        }
    }
    static void PaintText(Run run, string ChineseFont)
    {   
        if (run.RunProperties == null) {
            run.RunProperties = new RunProperties();
        }

        if (ChineseFont == "宋体")
        {
            run.RunProperties.Highlight = new Highlight() { Val = HighlightColorValues.Green };
        }
        else if (ChineseFont == "黑体")
        {
            run.RunProperties.Highlight = new Highlight() { Val = HighlightColorValues.Yellow };
        }
        else
        {
            run.RunProperties.Highlight = new Highlight() { Val = HighlightColorValues.Red };
        }
    }
    static string ExtractFontsByRunFonts(RunFonts runFonts)
    {
        string Chinese = null;

        if (runFonts != null)
        {
            string eastAsiaFont = runFonts.EastAsia;
            string eastAsiaThemeFont = runFonts.EastAsiaTheme;

            if (eastAsiaFont != null)
            {
                Chinese = eastAsiaFont;
            }
            else if (eastAsiaThemeFont != null)
            {
                Chinese = eastAsiaThemeFont;
            }
        }
       
        return Chinese;
    }
    static string JudgeFontsByRFonts(Dictionary<string, string> FontsDict, Dictionary<string, RunFonts> allRunFonts)
    {   
        string ChineseFont = null;

        if (ChineseFont == null)
        {
            ChineseFont = ExtractFontsByRunFonts(allRunFonts["runFontsOfRun"]);

            if (ChineseFont == null)
            {
                ChineseFont = ExtractFontsByRunFonts(allRunFonts["runFontsOfRunStyle"]);

                if (ChineseFont == null)
                {
                    ChineseFont = ExtractFontsByRunFonts(allRunFonts["runFontsOfParaStyle"]);

                    if (ChineseFont == null)
                    {
                        ChineseFont = ExtractFontsByRunFonts(allRunFonts["runFontsOfDefaultParaStyle"]);

                        if (ChineseFont == null)
                        {
                            ChineseFont = ExtractFontsByRunFonts(allRunFonts["runFontsOfDocDefault"]);
                        }
                    }
                }
            }
        }

        if (FontsDict.ContainsKey(ChineseFont))
        {
            ChineseFont = FontsDict[ChineseFont];
        }

        return ChineseFont;
    }
    static RunFonts getFontsByStyle(Style style)
    {
        if (style == null)
        {
            return null;
        }
        else
        {
            if (style.StyleRunProperties == null)
            {
                return null;
            }
            else
            {
                return style.StyleRunProperties.RunFonts;
            }
        }
    }
    static string getNameByStyle(Style style)
    {
        if (style == null)
        {
            return null;
        }
        else
        {
            return style.StyleName.ToString();
        }
    }
    static Style getStyleById(MainDocumentPart mainPart, string id)
    {
        StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;
        Styles styles = stylePart.Styles;
        Style style = styles.Descendants<Style>().Where(s => s.StyleId == id).FirstOrDefault();//如果应用default段落样式，就输入"a"

        return style;
    }
    static RunFonts getDocDefaultRunFonts(MainDocumentPart mainPart)
    {
        StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;
        DocDefaults docDefaults = stylePart.Styles.Descendants<DocDefaults>().FirstOrDefault();
        RunPropertiesDefault runPropertiesDefault = docDefaults.RunPropertiesDefault;
        RunPropertiesBaseStyle runProperties = runPropertiesDefault.RunPropertiesBaseStyle;
        RunFonts runFonts = runProperties.RunFonts;

        return runFonts;
    }

    static Dictionary<string, string> getFontsDict(MainDocumentPart mainPart, Dictionary<string, string> FontsDict)
    {
        ThemePart themePart = mainPart.ThemePart;
        DocumentFormat.OpenXml.Drawing.Theme theme = themePart.Theme;
        DocumentFormat.OpenXml.Drawing.FontScheme fontScheme = theme.Descendants<DocumentFormat.OpenXml.Drawing.FontScheme>().FirstOrDefault();
        if (fontScheme != null)
        {
            DocumentFormat.OpenXml.Drawing.MajorFont majorFont = fontScheme.MajorFont;
            DocumentFormat.OpenXml.Drawing.MinorFont minorFont = fontScheme.MinorFont;

            DocumentFormat.OpenXml.Drawing.LatinFont latinFont_major = majorFont.GetFirstChild<DocumentFormat.OpenXml.Drawing.LatinFont>();
            FontsDict["majorHAnsi"] = latinFont_major.Typeface.Value;

            DocumentFormat.OpenXml.Drawing.LatinFont latinFont_minor = minorFont.GetFirstChild<DocumentFormat.OpenXml.Drawing.LatinFont>();
            FontsDict["minorHAnsi"] = latinFont_minor.Typeface.Value;
            //获取latin的major与minor字体


            IEnumerable<DocumentFormat.OpenXml.Drawing.SupplementalFont> fonts_major = majorFont.Descendants<DocumentFormat.OpenXml.Drawing.SupplementalFont>();
            foreach (var font in fonts_major)
            {
                if (font.Script == "Hans")
                {
                    FontsDict["majorEastAsia"] = font.Typeface.Value;
                    break;
                }
            }

            IEnumerable<DocumentFormat.OpenXml.Drawing.SupplementalFont> fonts_minor = minorFont.Descendants<DocumentFormat.OpenXml.Drawing.SupplementalFont>();
            foreach (var font in fonts_minor)
            {
                if (font.Script == "Hans")
                {
                    FontsDict["minorEastAsia"] = font.Typeface.Value;
                    break;
                }
            }

            //IEnumerable<OpenXmlElement> fonts_major = majorFont.Descendants<OpenXmlElement>();
            //foreach (OpenXmlElement font in fonts_major)
            //{
            //    Dictionary<string, string> tempFontDict = new Dictionary<string, string>();
            //    tempFontDict.Add("Script", null);
            //    tempFontDict.Add("Typeface", null);

            //    var attributes = font.GetAttributes();

            //    foreach (var attribute in attributes)
            //    {
            //        if (attribute.LocalName == "latin")
            //        {
            //            tempFontDict["Script"] = attribute.LocalName;
            //        }

            //        if (attribute.LocalName == "script" && attribute.Value == "Hans")
            //        {
            //            tempFontDict["Script"] = attribute.Value;
            //        }

            //        if (attribute.LocalName == "typeface")
            //        {
            //            tempFontDict["Typeface"] = attribute.Value;
            //        }
            //    }

            //    if (tempFontDict["Script"] == "Hans")
            //    {
            //        FontsDict["majorEastAsia"] = tempFontDict["Typeface"];
            //    }

            //    if (FontsDict["majorEastAsia"] != null && FontsDict["majorHAnsi"] != null)
            //    {
            //        break;
            //    }

            //}

            //IEnumerable<OpenXmlElement> fonts_minor = minorFont.Descendants<OpenXmlElement>();
            //foreach (OpenXmlElement font in fonts_major)
            //{
            //    Dictionary<string, string> tempFontDict = new Dictionary<string, string>();
            //    tempFontDict.Add("Script", null);
            //    tempFontDict.Add("Typeface", null);

            //    var attributes = font.GetAttributes();

            //    foreach (var attribute in attributes)
            //    {
            //        if (attribute.LocalName == "latin")
            //        {
            //            FontsDict["minorHAnsi"] = attribute.Value;
            //        }

            //        if (attribute.LocalName == "script" && attribute.Value == "Hans")
            //        {
            //            tempFontDict["Script"] = attribute.Value;
            //        }

            //        if (attribute.LocalName == "typeface")
            //        {
            //            tempFontDict["Typeface"] = attribute.Value;
            //        }
            //    }

            //    if (tempFontDict["Script"] == "Hans")
            //    {
            //        FontsDict["minorEastAsia"] = tempFontDict["Typeface"];
            //    }

            //    if (FontsDict["minorEastAsia"] != null && FontsDict["minorHAnsi"] != null)
            //    {
            //        break;
            //    }

            //}

        }
        else
        {
            Console.WriteLine("Error : FontScheme is null");
            System.Environment.Exit(-1);
        }
        return FontsDict;
    }

    static void testFonts(string filePath)
    {
        Dictionary<string, string> FontsDict = new Dictionary<string, string>();
        FontsDict.Add("DefaultThemeFont", null);
        FontsDict.Add("DefaultParaStyle", null);
        FontsDict.Add("ParaStyle", null);
        FontsDict.Add("RunStyle", null);
        FontsDict.Add("RunFont", null);

        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;

            //文档样式定义

            Paragraph[] paragraphs = mainPart.Document.Body.Descendants<Paragraph>().ToArray();

            foreach(var paragraph in paragraphs)
            {
                ParagraphProperties paragraphProperties = paragraph.ParagraphProperties;

                Run[] runs = paragraph.Descendants<Run>().ToArray();

                foreach(Run run in runs)
                {
                    RunProperties runProperties = run.RunProperties;

                }

            }


        }
    }

    static void checkFont(string filePath)
    {
        Dictionary<string, List<string>> fontDictionary = new Dictionary<string, List<string>>();

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
        {
            Body body = wordDoc.MainDocumentPart.Document.Body;

            foreach (Paragraph paragraph in body.Elements<Paragraph>())
            {
                foreach (Run run in paragraph.Elements<Run>())
                {
                    if (run.RunProperties != null && run.RunProperties.RunFonts != null)
                    {
                        string fontName = run.RunProperties.RunFonts.Ascii;

                        // 如果字体名称为null，则设置为"NULL"
                        if (string.IsNullOrEmpty(fontName))
                        {
                            fontName = "NULL";
                        }

                        string text = run.InnerText.Trim();

                        if (!string.IsNullOrEmpty(text))
                        {
                            if (!fontDictionary.ContainsKey(fontName))
                            {
                                fontDictionary[fontName] = new List<string>();
                            }
                            fontDictionary[fontName].Add(text);
                        }
                    }
                }
            }
        }

        Console.WriteLine("字体检查结果：");
        foreach (var kvp in fontDictionary)
        {
            Console.WriteLine($"字体: {kvp.Key}");
            Console.WriteLine("文本内容：");
            foreach (var text in kvp.Value)
            {
                Console.WriteLine(text);
            }
            Console.WriteLine();
        }
    }

    static XDocument ExtractStylesPart(string fileName, string getStylesWithEffectsPart = "true")
    {
        // Declare a variable to hold the XDocument.
        XDocument? styles = null;

        // Open the document for read access and get a reference.
        using (var document = WordprocessingDocument.Open(fileName, false))
        {
            if (document.MainDocumentPart is null || document.MainDocumentPart.StyleDefinitionsPart is null || document.MainDocumentPart.StylesWithEffectsPart is null)
            {
                throw new ArgumentNullException("MainDocumentPart and/or one or both of the Styles parts is null.");
            }

            // Get a reference to the main document part.
            var docPart = document.MainDocumentPart;

            // Assign a reference to the appropriate part to the
            // stylesPart variable.
            StylesPart? stylesPart = null;

            if (getStylesWithEffectsPart.ToLower() == "true")
                stylesPart = docPart.StylesWithEffectsPart;
            else
                stylesPart = docPart.StyleDefinitionsPart;

            using var reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read));

            // Create the XDocument.
            styles = XDocument.Load(reader);
        }
        // Return the XDocument instance.
        return styles;
    }

    static void GetRunsWithStyle(string filePath, string styleId)
    {
        Dictionary<string, List<Run>> runsByStyle = new();

        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;

            //IEnumerable<Paragraph> paragraphs = mainPart.Document.Descendants<Paragraph>();
            Paragraph[] paragraphs = mainPart.Document.Body.Descendants<Paragraph>().ToArray(); 

            foreach (var para in paragraphs)
            {
                //IEnumerable<Run> runs = para.Descendants<Run>();
                Run[] runs = para.Descendants<Run>().ToArray();   

                foreach (var run in runs)
                {
                    RunProperties runProperties = run.RunProperties;

                    if (runProperties != null)
                    {
                        RunFonts runFonts = runProperties.RunFonts;

                        string asciiFont = runFonts.Ascii;
                        string highAnsiFont = runFonts.HighAnsi; 
                        string eastAsiaFont = runFonts.EastAsia; 
                        string complexScriptFont = runFonts.ComplexScript;

                        
                        Highlight highcolor = runProperties.GetFirstChild<Highlight>();

                      
                      

                        runProperties.Highlight = new Highlight() { Val = HighlightColorValues.Green };


                        runFonts.Ascii = "Times New Roman";
                        runFonts.HighAnsi = "Times New Roman";
                        runFonts.EastAsia = "华光准圆_CNKI";

                        Text text = run.GetFirstChild<Text>();
                        Console.WriteLine("修改前的文本为: " + text.Text);
                        //text.Text += $"（这段文字的字体信息为：Ascii：{asciiFont}，highAnsiFont：{highAnsiFont}，eastAsiaFont：{eastAsiaFont}，complexScriptFont：{complexScriptFont}";
                        //Console.WriteLine("修改后的文本为: " + text.Text);

                    }
                }
            }
            mainPart.Document.Save();
        }
       
    }

    static void setColorByFonts(string filePath)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;

           
            Paragraph[] paragraphs = mainPart.Document.Body.Descendants<Paragraph>().ToArray();
            //数组版本，用于Debug

            //IEnumerable<Paragraph> paragraphs = mainPart.Document.Body.Descendants<Paragraph>();
            //迭代器版本，用于Release

            foreach (var paragraph in paragraphs)
            {
                Run[] runs = paragraph.Descendants<Run>().ToArray();

                //IEnumerable<Run> runs = paragraph.Descendants<Run>();

                foreach (Run run in runs)
                {
                    string eastAsiaFont = run.RunProperties.RunFonts.EastAsia;

                    if (eastAsiaFont == null)//如果为空，说明没有设置字体，采用默认字体，这里需要插入一个获取默认字体的方法
                    {
                        run.RunProperties.Highlight = new Highlight() { Val = HighlightColorValues.Yellow };
                        //采用默认字体，设置高亮为黄色
                    }
                    else if (eastAsiaFont == "宋体")//如果为宋体，设置高亮为绿色
                    {
                        run.RunProperties.Highlight = new Highlight() { Val = HighlightColorValues.Green };
                    }
                    else
                    {
                        run.RunProperties.Highlight = new Highlight() { Val = HighlightColorValues.Red };
                        //其他字体设置为红色
                    }

                }
            }
            mainPart.Document.Save();//保存文档
        }
    }

    static Dictionary<string, string> getDefaultFonts(MainDocumentPart mainPart)
    {
        Dictionary<string, string> defaultFonts = new Dictionary<string, string>();

        return defaultFonts;
    }
}