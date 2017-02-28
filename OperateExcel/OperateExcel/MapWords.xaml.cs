using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;

namespace OperateExcel
{
    /// <summary>
    /// MapWords.xaml 的交互逻辑
    /// </summary>
    public partial class MapWords : System.Windows.Window
    {
        public MapWords()
        {
            InitializeComponent();
        }

        private Workbook ewb;
        public MapWords(Workbook excelwb)
        {
            InitializeComponent();
            ewb = excelwb;
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            //所有sheet页循环
            for (int i = 1; i <= ewb.Sheets.Count; i++)
            {
                Worksheet subTableSheet = ewb.Sheets[i] as Worksheet;
                string tempSheetName = subTableSheet.Name;
                if (tempSheetName.Contains("所有单词RES") ||
                    tempSheetName.Contains("所有单词logic") ||
                    tempSheetName.Contains("读音唯一")
                )
                {
                    continue;
                }
                 Dictionary<String, ArrayList> subTableWholeWordList = new Dictionary<string, ArrayList>();
                //检查字体
                int subTableSheetValidCount = GetValidCellRowCount(subTableSheet, subTableWholeWordList);
                for (int j = 2; j < subTableSheetValidCount; j++)
                {
                    String x;
                    try
                    {
                        x = ((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 4])).Font.Name;
                    }
                    catch
                    {
                        x = "Kingsoft Phonetic Plain";
                    }
                    //映射字体
                    ((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 8])).Value = FontName(x);

                    try
                    {
                        //插入RES
                        WordRES wordres = new WordRES();
                        //wordres.num = j;
                        wordres.word = ((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 2])).Value;

                        if (null != (((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 3])).Value))
                        {
                            wordres.subject = (int)(((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 3])).Value);
                        }

                        wordres.phoneticSymbol = ((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 4])).Value;
                        wordres.wordMeaning = ((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 5])).Value;
                        if (null != (((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 6])).Value))
                        {

                            wordres.unit = (int)((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 6])).Value;
                        }

                        if (null != (((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 7])).Value))
                        {

                            wordres.book = (int)(((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 7])).Value) - 1;
                        }

                        if (null != (((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 8])).Value))
                        {

                            wordres.fontType = (int)((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 8])).Value;
                        }

                        wordres.CreateAndFlush();

                        //插入Logic
                        WordLogic wordlogic = new WordLogic();
                        //wordlogic.num = j - 1;
                        wordlogic.word = ((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 2])).Value;

                        if (null != (((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 3])).Value))
                        {

                            wordlogic.subject = (int)((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 3])).Value;
                        }
                        wordlogic.remCount = 1;
                        wordlogic.lastRemTime = new DateTime(2014, 7, 8, 18, 6, 0);
                        wordlogic.nextRemTime = new DateTime(2014, 7, 8, 18, 6, 0);

                        if (null != (((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 6])).Value))
                        {

                            wordlogic.unit = (int)((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 6])).Value;
                        }
                        if (null != (((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 7])).Value))
                        {
                            wordlogic.book = (int)(((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 7])).Value) - 1;
                        }

                        wordlogic.CreateAndFlush();

                    }
                    catch (Exception exc)
                    {
                        Console.WriteLine(exc);
                    }
                }
            }


        }

        private int GetValidCellRowCount(Worksheet sheet, Dictionary<String, ArrayList> wordDict)
        {
            int i = 0;
            for (i = 1; i < sheet.Cells.Height; i++)
            {

                if (!String.IsNullOrEmpty(((Microsoft.Office.Interop.Excel.Range)(sheet.Cells[i, 2])).Value as string))
                {
                    //Console.WriteLine(((Microsoft.Office.Interop.Excel.Range)(sheet.Cells[i, 2])).Value);
                    String word = ((Microsoft.Office.Interop.Excel.Range)(sheet.Cells[i, 2])).Value;
                    word = word.Trim();
                    ArrayList wordIndex = new ArrayList();
                    if (!wordDict.ContainsKey(word))
                    {
                        wordIndex.Add(i);
                        wordDict.Add(word, wordIndex);
                    }
                    else
                    {
                        //在key值对应的alueList里追加
                        wordDict[word].Add(i);
                    }
                    continue;
                }
                else
                {
                    if ((String.IsNullOrEmpty(((Microsoft.Office.Interop.Excel.Range)(sheet.Cells[i + 1, 2])).Value as string)) &&
                        (String.IsNullOrEmpty(((Microsoft.Office.Interop.Excel.Range)(sheet.Cells[i + 2, 2])).Value as string)) &&
                        (String.IsNullOrEmpty(((Microsoft.Office.Interop.Excel.Range)(sheet.Cells[i + 3, 2])).Value as string)))
                    {
                        break;
                    }
                }
            }

            return i;
        }

        private int FontName(string fontName)
        {
            switch (fontName)
            {
                case "Kingsoft Phonetic Plain":
                    return 1;
                case "宋体":
                    return 2;
                case "Times New Roman":
                    return 3;
                case "Lucida Sans Unicode":
                    return 4;
                case "Verdana":
                    return 5;
                case "Arial":
                    return 6;
                default:
                    return -1;
            }
        }

    }
}
