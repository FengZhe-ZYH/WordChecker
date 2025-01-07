using System;
using System.IO;
using System.Text;

public class XmlFormatter
{
    public static void FormatAndSaveXml()
    {
        StringBuilder xmlContentBuilder = new StringBuilder();
        string inputLine; while ((inputLine = Console.ReadLine()) != "END")
        {
            xmlContentBuilder.AppendLine(inputLine);

        }
        string xmlContent = xmlContentBuilder.ToString();

        FormatAndSaveXml_helper(xmlContent, "C:\\Users\\FengZhe\\Desktop\\Study\\Gproject\\WordChecker\\assest\\Theme.txt");
    }
    public static void FormatAndSaveXml_helper(string xmlContent, string outputPath)
    {
        StringBuilder formattedXml = new StringBuilder();
        int indentLevel = 0;
        bool inTag = false;

        for (int i = 0; i < xmlContent.Length; i++)
        {
            char currentChar = xmlContent[i];

            if (currentChar == '<')
            {
                if (i < xmlContent.Length - 1 && xmlContent[i + 1] == '/')
                {
                    indentLevel--;
                    AppendNewLineAndIndent(formattedXml, indentLevel);
                }
                else
                {
                    if (inTag)
                    {
                        AppendNewLineAndIndent(formattedXml, indentLevel);
                    }
                    AppendNewLineAndIndent(formattedXml, indentLevel);
                    if (i < xmlContent.Length - 1 && xmlContent[i + 1] != '/')
                    {
                        indentLevel++;
                    }
                }

                inTag = true;
            }

            formattedXml.Append(currentChar);

            if (currentChar == '>')
            {
                inTag = false;
                if (i > 0 && xmlContent[i - 1] == '/')
                {
                    indentLevel--;
                }
            }
        }

        File.WriteAllText(outputPath, formattedXml.ToString());
        Console.WriteLine($"Formatted XML has been saved to {outputPath}");
    }

    private static void AppendNewLineAndIndent(StringBuilder sb, int indentLevel)
    {
        sb.Append(Environment.NewLine);
        sb.Append(new string(' ', indentLevel * 2)); // 每一级缩进5个空格
    }
}
