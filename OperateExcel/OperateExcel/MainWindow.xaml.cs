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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Xml.Linq;
using Castle.ActiveRecord.Framework;
using Castle.ActiveRecord;
using SpeechLib;
using System.Threading;


namespace OperateExcel
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        //m_Excel为excel应用程序对象
        private Microsoft.Office.Interop.Excel.Application m_Excel;
        //m_Workbooks为所有工作簿对象
        private Workbooks m_Workbooks;
        //parentFirstWordCount总表第一个单词数量
        private int parentFirstWordCount = -1;
        //parentSecondWordCount总表第二个单词数量
        private int parentSecondWordCount = -1;
        //parentWorkBook为工作簿对象
        private Workbook parentWorkBook;

        //选择地址按钮
        private void StartTOChangeButton_Click(object sender, RoutedEventArgs e)
        {
            //打开Excel
            if (null == m_Excel)
            {
                m_Excel = new Microsoft.Office.Interop.Excel.Application();
            }


            //地址栏如果为空则不执行
            if (String.IsNullOrEmpty(textBox1.Text) || String.IsNullOrEmpty(textBox2.Text))
            {
                return;
            }

            //path
            ////临时文件名称
            Regex pathrgx = new Regex(@".*\\(.*\..*)");
            Match subTableMth = pathrgx.Match(textBox1.Text);
            //Match parentTableMth = pathrgx.Match(textBox2.Text);
            String subWorkbookName = subTableMth.Groups[1].ToString();

            //分表单词表
            String subWordTable = textBox1.Text;


            //总表单词表
            //String parentWordTable = textBox2.Text;



            //执行遍历
            //打开Excel,同时打开两个excel

            m_Workbooks.Open(subWordTable);


            Workbook subWorkbook = m_Workbooks[subWorkbookName];
            string subWordTableName = subWorkbook.FullName;


            //在总表中未找到的单词，进入数组中
            ArrayList wordNotFound = new ArrayList();
            ArrayList phraseNotFound = new ArrayList();

            //依次选择子分表中的sheet页，遍历其中的单词，与总表中的单词对比
            for (int i = 1; i <= subWorkbook.Sheets.Count; i++)
            {

                Worksheet subTableSheet = subWorkbook.Sheets[i] as Worksheet;
                string tempSheetName = subTableSheet.Name;

                if (tempSheetName.Contains("所有单词RES") ||
                    tempSheetName.Contains("所有单词logic") ||
                    tempSheetName.Contains("读音唯一")
                )
                {
                    continue;
                }
                subTableWholeWordList.Clear();
                int subTableSheetValidCount = GetValidCellRowCount(subTableSheet, subTableWholeWordList);

                Console.WriteLine(parentFirstWordCount + "; " + parentSecondWordCount);
                for (int j = 2; j < subTableSheetValidCount; j++)
                {
                    String word = ((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 2])).Value;
                    string wordMeaning;
                    string phoneticSymbol;
                    //重新计算表中cell数，传入
                    Boolean isExist = checkIfExistInParent(word.Trim(), out wordMeaning, out phoneticSymbol, parentFirstWordCount, parentSecondWordCount);
                    Console.WriteLine(word + "  " + isExist);

                    String subWordMeaning = ((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 5])).Value;

                    if (String.IsNullOrEmpty(subWordMeaning))
                    {
                        if (isExist == true)
                        {
                            ((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 5])).Value = wordMeaning;
                            ((Microsoft.Office.Interop.Excel.Range)(subTableSheet.Cells[j, 4])).Value = phoneticSymbol;
                        }
                        else
                        {
                            if (word.Trim().Contains(" "))
                            {
                                phraseNotFound.Add(word);
                            }
                            else
                            {
                                wordNotFound.Add(word);
                            }
                        }
                    }


                }

            }

            AddUnfoundWordToParent(wordNotFound, phraseNotFound);


            //subWorkbook.SaveAs(subTablePath);
            //parentWorkBook.SaveAs(parentTablePath);

            subWorkbook.Save();
            parentWorkBook.Save();

            if (MessageBox.Show("单词查找完毕!", "查找结果", MessageBoxButton.OK) == MessageBoxResult.OK)
            {
                subWorkbook.Close();
                subWorkbook = null;
            }

        }
        //新建三个字典，键值对为字符串和数组，分别为第一个sheet所有单词列表，第二个sheet所有单词列表，子表格所有单词列表
        private Dictionary<String, ArrayList> firstSheetWholeWordList = new Dictionary<string, ArrayList>();
        private Dictionary<String, ArrayList> secondSheetWholeWordList = new Dictionary<string, ArrayList>();
        private Dictionary<String, ArrayList> subTableWholeWordList = new Dictionary<string, ArrayList>();
        //方法：获取有效的单元格行数
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
        //方法：把没有找到的单词加入总表
        private void AddUnfoundWordToParent(ArrayList wordNotFound, ArrayList phraseNotFound)
        {
            string parentWordTableName = parentWorkBook.FullName;
            Worksheet firstSheet = parentWorkBook.Sheets[1] as Worksheet;
            Worksheet secondSheet = parentWorkBook.Sheets[2] as Worksheet;

            int i = parentFirstWordCount;
            Console.WriteLine(i + "; ");

            foreach (String word in wordNotFound)
            {
                ((Microsoft.Office.Interop.Excel.Range)(firstSheet.Cells[i, 2])).Value = word;
                i++;
            }

            int j = parentSecondWordCount;
            Console.WriteLine(j + "; ");
            foreach (String word in phraseNotFound)
            {
                ((Microsoft.Office.Interop.Excel.Range)(secondSheet.Cells[j, 2])).Value = word;
                j++;
            }
        }
        //方法：检查是否在总表中存在
        private Boolean checkIfExistInParent(String word, out String meaning, out String phoneticSymbol, int validCellRowCountFirstSheet, int validCellRowCountSecondSheet)
        {
            //Workbook parentWorkBook = m_Workbooks.get_Item(2);
            string parentWordTableName = parentWorkBook.FullName;
            Worksheet firstSheet = parentWorkBook.Sheets[1] as Worksheet;
            Worksheet secondSheet = parentWorkBook.Sheets[2] as Worksheet;

            //检查有用的word
            if (firstSheetWholeWordList.ContainsKey(word))
            {
                #region 循环list,将所有的解释追加到一个单词解释中
                ////循环list,将所有的解释追加到一个单词解释中
                //String cellMeaning = "";
                //foreach (int i in firstSheetWholeWordList[word])
                //{
                //    cellMeaning += ((Microsoft.Office.Interop.Excel.Range)(firstSheet.Cells[i, 4])).Value + " ;  ";
                //}
                #endregion

                #region 只保留遇到的第一个形似的单词意思
                String cellMeaning = "";
                foreach (int i in firstSheetWholeWordList[word])
                {
                    cellMeaning += ((Microsoft.Office.Interop.Excel.Range)(firstSheet.Cells[i, 4])).Value + " ;  ";
                    break;
                }
                #endregion

                String cellPhoneticSymbol = ((Microsoft.Office.Interop.Excel.Range)(firstSheet.Cells[firstSheetWholeWordList[word][0], 3])).Value as String;
                meaning = cellMeaning;
                phoneticSymbol = cellPhoneticSymbol;
                return true;
            }

            if (secondSheetWholeWordList.ContainsKey(word))
            {
                //返回的值填入到分表单元格中
                //循环list,将所有的解释追加到一个单词解释中
                String cellMeaning = "";
                foreach (int i in secondSheetWholeWordList[word])
                {
                    cellMeaning += ((Microsoft.Office.Interop.Excel.Range)(secondSheet.Cells[i, 3])).Value + " ;\r\n";
                }

                meaning = cellMeaning;
                phoneticSymbol = "";
                return true;
            }

            meaning = "";
            phoneticSymbol = "";
            return false;

        }
        //当界面加载时，实例化两个对象
        private void onFormLoaded(object sender, RoutedEventArgs e)
        {
            //打开Excel
            m_Excel = new Microsoft.Office.Interop.Excel.Application();
            m_Workbooks = m_Excel.Application.Workbooks;
            //initDataBase();
        }
        //初始化数据库，生成各种字段
        private void initDataBase()
        {
            //IConfigurationSource source = new Castle.ActiveRecord.Framework.Config.XmlConfigurationSource("../TestCases/ActiveRecordConfig.xml");


            //// 载入程序集中所有 ActiveRecord 类。
            //ActiveRecordStarter.Initialize(source, typeof(WordRES), typeof(WordLogic));

            WordRES u = WordRES.Find(1);
            WordRES m = new WordRES();
            m.CreateAndFlush();

        }
        //界面关闭时，退出excel
        private void onFormClosed(object sender, EventArgs e)
        {

            m_Workbooks.Close();
            m_Excel.Quit();
            m_Excel = null;
        }
        //选择分表路径
        private void chooseSubTabelPath(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory = Directory.GetParent("./TestCases/").ToString();
            openFile.Filter = "Excel Documents (*.xlsx)|*.xlsx|Excel Documents (*.xls)|*.xls|All Files (*.*)|*.*";
            //openFile.Filter = "All Files (*.*)|*.*";
            if (openFile.ShowDialog() == true)
            {
                textBox1.Text = openFile.FileName;
            }


        }
        //选择总表路径
        private void chooseParentTablePath(object sender, RoutedEventArgs e)
        {
            if (null != parentWorkBook)
            {
                parentWorkBook.Close();
                parentWorkBook = null;
            }

            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory = Directory.GetParent("./TestCases/").ToString();
            openFile.Filter = "Excel Documents (*.xlsx)|*.xlsx|Excel Documents (*.xls)|*.xls|All Files (*.*)|*.*";
            //openFile.Filter = "All Files (*.*)|*.*";
            if (openFile.ShowDialog() == true)
            {
                textBox2.Text = openFile.FileName;
            }

            Regex pathrgx = new Regex(@".*\\(.*\..*)");
            Match parentTableMth = pathrgx.Match(textBox2.Text);
            String parentWorkbookName = parentTableMth.Groups[1].ToString();

            String parentWordTable = textBox2.Text;
            m_Workbooks.Open(parentWordTable);
            //parentWorkBook = m_Workbooks.get_Item(2);
            parentWorkBook = m_Workbooks[parentWorkbookName];
            string parentWordTableName = parentWorkBook.FullName;
            Worksheet firstSheet = parentWorkBook.Sheets[1] as Worksheet;
            Worksheet secondSheet = parentWorkBook.Sheets[2] as Worksheet;
            parentFirstWordCount = GetValidCellRowCount(firstSheet, firstSheetWholeWordList);
            parentSecondWordCount = GetValidCellRowCount(secondSheet, secondSheetWholeWordList);
        }

        //button4，将各个子sheet页的内容插入到RES和Logic表中，并填入其他所缺的内容
        private void insert_words_to_list(object sender, RoutedEventArgs e)
        {
            Regex pathrgx = new Regex(@".*\\(.*\..*)");

            if (String.IsNullOrEmpty(mappath.Text))
            {
                MessageBox.Show("检查文档不能为空！");
                return;
            }

            Match parentTableMth = pathrgx.Match(mappath.Text);
            m_Workbooks.Open(mappath.Text);
            String parentWorkbookName = parentTableMth.Groups[1].ToString();
            parentWorkBook = m_Workbooks[parentWorkbookName];
            string parentWordTableName = parentWorkBook.FullName;

            Regex versionRgx = new Regex(@"([0-9]+)_(.*)");


            //所有sheet页循环
            for (int i = 1; i <= parentWorkBook.Sheets.Count; i++)
            {
                Worksheet subTableSheet = parentWorkBook.Sheets[i] as Worksheet;
                //特殊名称的sheet页跳过
                string tempSheetName = subTableSheet.Name;
                if (tempSheetName.Contains("所有单词RES") ||
                    tempSheetName.Contains("所有单词logic") ||
                    tempSheetName.Contains("读音唯一") ||
                    tempSheetName.Contains("_未分节")
                )
                {
                    continue;
                }

                //解析版本序号名
                Match versionMatch = versionRgx.Match(tempSheetName);
                int bookId = -1;
                int.TryParse(versionMatch.Groups[1].ToString(), out bookId);

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

                        //书号
                        wordres.book = bookId - 1;

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
                        //书号
                        wordlogic.book = bookId - 1;

                        wordlogic.CreateAndFlush();

                    }
                    catch (Exception exc)
                    {
                        Console.WriteLine(exc);
                    }
                }
            }

            parentWorkBook.Save();
            MessageBox.Show("单词对应完毕！");
        }
        //字体
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

        private static SpObjectToken englishToken = null;
        private static SpObjectToken chineseToken = null;
        //button5生成读音
        private void button5_Click(object sender, RoutedEventArgs e)
        {
            //数据库全部读出
            WordLogic[] wordList = WordLogic.FindAll();

            SpFileStream SpFileStream = new SpFileStream();
            String fileName;
            SpeechVoiceSpeakFlags SpFlags = SpeechVoiceSpeakFlags.SVSFlagsAsync;
            SpeechStreamFileMode SpFileMode = SpeechStreamFileMode.SSFMCreateForWrite;
            SpVoice voice = new SpVoice();

            // prepare voice
            SpObjectTokenCategory aotc = new SpObjectTokenCategory();
            aotc.SetId(SpeechLib.SpeechStringConstants.SpeechCategoryVoices);

            foreach (ISpeechObjectToken token in aotc.EnumerateTokens())
            {
                if (token.GetDescription() == "VW Julie")
                {
                    englishToken = (SpObjectToken)token;
                }
                else if (token.GetDescription() == "VW Hui")
                {
                    chineseToken = (SpObjectToken)token;
                }
            }

            voice.Voice = englishToken;
            voice.Rate = -4;

            String outFolderPath = Directory.GetParent("../../TestCases/") + @"\单词音频\";

            if (!Directory.Exists(outFolderPath))
            {
                Directory.CreateDirectory(outFolderPath);
            }

            for (int i = 0; i < wordList.Length; i++)
            {
                String word = wordList[i].word;

                if (word != null)
                {
                    word = word.Trim();
                }

                if (String.IsNullOrEmpty(word))
                {
                    // 遇到无效内容，退出
                    continue;
                }
                word = convert(word);

                fileName = outFolderPath + word + ".wav";
                SpFileStream.Open(fileName, SpFileMode, false);
                voice.AudioOutputStream = SpFileStream;
                voice.Speak(word, SpFlags);
                voice.WaitUntilDone(Timeout.Infinite);
                SpFileStream.Close();
            }


            MessageBox.Show("音频生成完毕！");
        }

        private String convert(String input)
        {
            char[] leagleChars = { 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            ArrayList lcs = new ArrayList(leagleChars);

            char[] retChars = { };

            ArrayList retList = new ArrayList();

            foreach (char c in input.ToCharArray())
            {
                if (!lcs.Contains(c))
                {
                    retList.Add('_');
                }
                else
                {
                    retList.Add(c);
                }
            }

            return new String((char[])retList.ToArray(typeof(char)));
        }
    }


}

