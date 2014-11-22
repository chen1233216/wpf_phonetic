using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


//office
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

//phonetic
using System.Collections;
using System.Web;
using System.Net;
using System.Text.RegularExpressions;

//thread
using System.Threading;

namespace WpfApplication1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {

        private Word.Document doc;
        private static object Nothing = System.Reflection.Missing.Value;

        private SearchWord search = new SearchWord();
        private CodeConvert code = new CodeConvert();
        private string errorFile = "error.txt";
        private bool no_errors = true;
        private string cur_file;


        #region　-　插入分页符　-
        public void InsertBreak()
        {
            Word.Paragraph para;
            para = doc.Content.Paragraphs.Add(ref　Nothing);
            object pBreak = (int)Word.WdBreakType.wdSectionBreakNextPage;
            para.Range.InsertBreak(ref　pBreak);
        }
        #endregion

        public MainWindow()
        {
            InitializeComponent();
        }

        private void init()
        {
            Word.Application theApplication = new Word.Application();
            theApplication.Visible = true;
            object missing = Type.Missing;
            Word.Document theDocument = theApplication.Documents.Add(
                ref missing,
                ref missing,
                ref missing,
                ref missing);
            Word.Range range = theDocument.Range(ref missing, ref missing);
            int rowCount = 4, colCount = 4;
            Word.Table table = range.Tables.Add(
                range,
                rowCount,
                colCount,
                ref missing,
                ref missing);
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    Word.Cell cell = table.Cell(i + 1, j + 1);
                    cell.Range.Text = i.ToString();
                }
            }
        }

        private void combine_doc(string _path)
        {
            String path = _path;// @"D:\Data\wer\第二本书\第10章 J 6\注音版";
            String archiveDirectory = path + @"\和并版";

            if (!Directory.Exists(archiveDirectory))
            {
                Directory.CreateDirectory(archiveDirectory);
            }
            var docFiles = Directory.EnumerateFiles(path, "*.doc*");

            Word.Application theApplication = new Word.Application();
            theApplication.Visible = true;
            doc = theApplication.Documents.Add();
            Word.Range com_range = doc.Range();

            try
            {
                int count = 0;
                List<String> allFiles = new List<string>();
                foreach (string currentFile in docFiles)
                {
                    allFiles.Add(currentFile);
                }
                allFiles.Sort();
                foreach (string currentFile in allFiles)
                {
                    count++;
                    com_range.SetRange(doc.Range().End, doc.Range().End);
                    string fileName = currentFile.Substring(path.Length + 1);
                    Word.Document doc_temp = theApplication.Documents.Open(@currentFile);
                    Word.Range ran = doc_temp.Range();
                    ran.Select();
                    ran.Copy();
                    com_range.Paste();
                    InsertBreak();
                    doc_temp.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                //MessageBox.Show("size = " + count, "size");
            }
            catch (Exception e_)
            {
                MessageBox.Show("Exception.!!! ", "Some Exception ! ");
            }
            finally
            {
                doc.SaveAs2(archiveDirectory + @"\合并.doc");
                doc.Close();
                theApplication.Quit();
            }
        }

        //添加一条error信息
        private void addRrror(String error)
        {

            FileStream fs = new FileStream(this.errorFile, FileMode.Append);
            StreamWriter sw = new StreamWriter(fs);

            if (no_errors)
            {
                sw.Write("Error in File : " + cur_file + "\n");
                no_errors = false;
            }

            //开始写入
            sw.Write(error + "\n");
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();

        }


        //一段一段选中
        private void everyParagraph(Word.Document doc)
        {
            object missing = Type.Missing;

#if _Debug_Show_
            MessageBox.Show("paragraph count :" + doc.Paragraphs.Count);
#endif
            //遍历每一段
            for (int i = 1; i <= doc.Paragraphs.Count; i++)
            {
#if _Debug_Show_
                doc.Paragraphs[i].Range.Select();
                MessageBox.Show("p: " + i);
#endif
                //先寻找'/'
                Word.Range rng = doc.Paragraphs[i].Range;   //获取范围
                String lineText = doc.Paragraphs[i].Range.Text; //记录该行内容
                //在范围中搜索“/”
                rng.Find.ClearFormatting();
                rng.Find.Forward = true;
                rng.Find.Text = "[";
                rng.Find.Execute(
                     ref missing, ref missing, ref missing, ref missing, ref missing,
                     ref missing, ref missing, ref missing, ref missing, ref missing,
                     ref missing, ref missing, ref missing, ref missing, ref missing);
#if _Debug_Show_
                Word.Range debugRange = doc.Content;//为了方便debug所设置的范围变量，用来在debug时显示当先所选中的区域
#endif

                if (rng.Find.Found)//该行需要处理
                {
                    //找到“/”位置，并记录位置
                    object start1 = rng.Start;
                    object end1 = rng.End;
#if _Debug_Show_
                    debugRange.SetRange((int)start1, (int)end1);
                    debugRange.Select();
                    MessageBox.Show("Start1: " + start1 + " End1: " + end1, "Range Information : ");
#endif
                    //第二次执行搜索
                    rng.Find.Text = "]";
                    rng.SetRange((int)end1, doc.Paragraphs[i].Range.End);
                    rng.Find.Execute(
                         ref missing, ref missing, ref missing, ref missing, ref missing,
                         ref missing, ref missing, ref missing, ref missing, ref missing,
                         ref missing, ref missing, ref missing, ref missing, ref missing);
                    if (!rng.Find.Found)
                    {
                        //MessageBox.Show(@"error : second / not found.");
                        this.addRrror(@" second / not found : " + lineText);
                        continue;
                    }
                    //记录第二次搜索结果
                    object start2 = rng.Start;
                    object end2 = rng.End;
#if _Debug_Show_
                    debugRange.SetRange((int)start2, (int)end2);
                    debugRange.Select();
                    MessageBox.Show("Start2: " + start2 + " End2: " + end2, "Range Information : ");
#endif
                    //解析字符串lineText,获取需要注音的单词组
                    String[] str = lineText.Split('[');//先去掉音标及其之后的部分
                    String[] words = new String[10]; int wordCount = 0;
                    String wordPhons = "";
                    if (str == null || str.Length < 1)
                    {
                        //MessageBox.Show("Error", "Range Information : ");
                        this.addRrror(" : format error." + lineText);
                        continue;
                    }
                    Regex r = new Regex("([a-zA-Z]+)");//用来解析的正则
                    if (r.IsMatch(str[0]))
                    {
                        //首次匹配 
                        Match m = r.Match(str[0]);
                        while (m.Success && wordCount < 10)
                        {
                            words[wordCount++] = m.Value;
                            //下一个匹配 
                            m = m.NextMatch();
                        }
                    }
#if _Debug_Show_
                    String wordResult = "";
                    int count = wordCount;
                    while (--count >= 0)
                    {
                        wordResult += words[count] + "  ";
                    }
                    MessageBox.Show(wordResult, "Range Information : ");
#endif
                    if (wordCount <= 0)
                        continue;
                    //查询单词的音标
                    for (int iWord = 0; iWord < wordCount; iWord++)
                    {
                        if (search.getPhonetic(words[iWord]).Equals("wrong"))
                        {
                            this.addRrror(" : phonetic not found ." + lineText);
                        }
                        if (iWord > 0)
                            wordPhons += "-" + search.getPhonetic(words[iWord]);
                        else
                            wordPhons = search.getPhonetic(words[iWord]);
                    }
                    //测试字符映射
                    wordPhons = code.charConv(wordPhons);
#if _Debug_Show_
                    MessageBox.Show(wordPhons, "Phonetic Information : ");
#endif

                    //使用剪贴板的粘贴操作
                    Word.Range rngReplace = doc.Range(ref end1, ref start2);
#if _Debug_Show_
                    //rngReplace.Select();
                    //MessageBox.Show(rngReplace.Text, "Range Information : ");
#endif
                    String replaceText = wordPhons;
                    if (replaceText == null)
                        continue;
                    Clipboard.SetText(replaceText, TextDataFormat.Text);
                    rngReplace.Paste();

                    //设置粘贴后的字体
                    object endFont = (int)end1 + replaceText.Length;
                    Word.Range rngFont = doc.Range(ref end1, ref endFont);
                    rngFont.Font.Name = "Kingsoft Phonetic Plain";
#if _Debug_Show_
                    rngFont.Select();
                    MessageBox.Show("Paste Result.!!! ", "Range Information : ");
#endif
                }


            }
        }


        private void phonety_doc(object _path)
        {

            string temp = _path.ToString();
            String path =temp;
            String archiveDirectory = path + @"\注音版";
            if (!Directory.Exists(archiveDirectory))
            {
                Directory.CreateDirectory(archiveDirectory);
            }
            string completeDirectory = path + @"\已完成";
            if (!Directory.Exists(completeDirectory))
            {
                Directory.CreateDirectory(completeDirectory);
            }
            string exceptionDirectory = path + @"\异常";
            var docFiles = Directory.EnumerateFiles(path, "*.doc*");
            Word.Application theApplication = new Word.Application();
            theApplication.Visible = true;

            try
            {
                //progressBar
                //this.pbar_1.Visibility = System.Windows.Visibility.Visible;
                //this.pbar_1.Minimum = 0;
                //this.pbar_1.Maximum = 10;
                //this.pbar_1.Value = 0;
                
                foreach (string currentFile in docFiles)
                {
                    cur_file = currentFile;
                    no_errors = true;
                    string fileName = currentFile.Substring(path.Length + 1);
                    doc = theApplication.Documents.Open(@currentFile);//打开word文档
                    everyParagraph(doc);
                    if (no_errors)
                    {
                        string target = archiveDirectory + @"\" + fileName;
                        doc.SaveAs2(target);
                        doc.Close();
                        Directory.Move(currentFile, completeDirectory + @"\" + fileName);
                        //MessageBox.Show("source : " + currentFile + "\n target : " + target, "Inf.!!! ");
                    }
                    else
                    {
                        doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                        MessageBox.Show("Errors.!!! ", "Some Errors ! Save failed! ");
                    }
                    //this.pbar.Value += 1;
                }

                //if (this.pbar.Value >= this.pbar.Maximum)
                //{
                //    MessageBox.Show("完成 ", "Inf");
                //}
            }
            catch (Exception ee)
            {
                if (!Directory.Exists(exceptionDirectory))
                {
                    Directory.CreateDirectory(exceptionDirectory);
                }

                addRrror("Exception !!!!!!\n");
                if (doc != null)
                    doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                string fileName = cur_file.Substring(path.Length + 1);
                Directory.Move(cur_file, exceptionDirectory + @"\" + fileName);
                MessageBox.Show(ee.ToString(), "Exception");
            }
            finally
            {
                this.search.saveWord();
                //this.pbar.Visibility = System.Windows.Visibility.Hidden;
            }
        }

        private void pbStart_Click(object sender, RoutedEventArgs e)
        {
            string path = @combine_addr.Text;
            MessageBox.Show("path = " + path, "path");
            try
            {
                this.combine_doc(path);
            }
            catch(Exception ee){
                MessageBox.Show("Exception.!!! "+ee.ToString(), "Some Exception ! ");
            }
        }

        private void pbPhenetic_Click(object sender, RoutedEventArgs e)
        {
            string path = @phonetic_addr.Text;
            MessageBox.Show("path = " + path, "path");
            this.pbar_1.Maximum = 100;
            try
            {
                Thread phoThread = new Thread(phonety_doc);
                phoThread.SetApartmentState(ApartmentState.STA);
                phoThread.Start(path);
            }
            catch (Exception ee)
            {
                MessageBox.Show("Exception.!!! " + ee.ToString(), "Some Exception ! ");
            }
        }
    }

    //该类负责将单词的音标转换成Kingsoft Phonetic Plain字体所用的编码
    class CodeConvert
    {
        private bool existed = false;
        private string conFile = "charConv.txt";
        private string missingFile = "missWord.txt";
        private Hashtable map;

        public CodeConvert()
        {
            if (!existed)
            {
                if (!existed)//同步
                {
                    map = new Hashtable();
                    readFile();
                    existed = true;
                }
            }
        }

        //读取字符映射表
        private void readFile()
        {
            StreamReader objReader = new StreamReader(this.conFile);
            string sLine = "";
            String[] str;
            char[] para = { ',' };
            while (sLine != null)
            {
                sLine = objReader.ReadLine();
                if (sLine == null)
                {
                    //MessageBox.Show("No config file.");
                    break;
                }
                str = sLine.Split(para);
                if (str == null)
                {
                    //MessageBox.Show("No config file.");
                    break;
                }
                if (str.Length == 2)
                {
                    char a = (char)(int.Parse(str[0]));
                    this.map.Add(str[1], a);
                }
                else if (str.Length == 3)
                {
                    char a = (char)(int.Parse(str[0]));
                    this.map.Add(",", a);
                }
                else
                {
                    MessageBox.Show("read file error.!!!!!");
                }
            }
            objReader.Close();

        }

        //增加一个字符到丢失文件
        private void addMiss(String miss)
        {
            FileStream fs = new FileStream(this.missingFile, FileMode.Append);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入

            sw.Write(miss + "\n");
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();

        }

        //转换字符
        public String charConv(String word)
        {
            String result = word;
            for (int i = 0; i < word.Length; i++)
            {
                if (!map.ContainsKey(word[i].ToString()))
                {
                    addMiss(word[i].ToString());
                }
            }
            foreach (DictionaryEntry de in map)
            {
                if (result.Contains(de.Key.ToString()))
                {
                    result = result.Replace(de.Key.ToString(), de.Value.ToString());
                }
            }

            return result;
        }
    }

    //该类负责获取单词的音标
    class SearchWord
    {
        //单词音标缓存：减少网络查询次数
        private Hashtable table = null;    //映射表
        private bool existed = false;
        private String conFile = "wordConf.txt";    //配置文件名
        private long workCount = 0;     //执行总次数
        private long searchCount = 0;   //网络查询总次数

        //避免创建对象
        public SearchWord()
        {
            if (!existed)
            {
                if (!existed)//在这儿加同步：优化的懒汉式
                {
                    table = new Hashtable();
                    table.Clear();
                    readWord();
                    existed = true;
                }
            }
        }

        ~SearchWord()
        {

        }
        //加载单词记录
        private void readWord()
        {
            StreamReader objReader = new StreamReader(this.conFile);
            string sLine = "";
            String[] str;
            char[] para = { '#' };
            while (sLine != null)
            {
                sLine = objReader.ReadLine();
                if (sLine == null)
                {
                    //MessageBox.Show("No config file.");
                    break;
                }
                str = sLine.Split(para);
                if (str == null)
                {
                    //MessageBox.Show("No config file.");
                    break;
                }
                if (str.Length == 3) //第一行的统计信息
                {
                    this.searchCount = long.Parse(str[1]);
                    this.workCount = long.Parse(str[2]);
                }
                else if (str.Length == 2)
                {
                    this.table.Add(str[0], str[1]);
                }
                else if (str.Length == 0)
                {
                    //MessageBox.Show("read file complete.");
                }
                else
                {
                    MessageBox.Show("read file error.!!!!!");
                }
            }
            objReader.Close();
        }
        //保存单词记录
        public void saveWord()
        {
            FileStream fs = new FileStream(this.conFile, FileMode.OpenOrCreate);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入
            sw.Write("sum#" + this.searchCount + "#" + this.workCount + "\n");
            foreach (DictionaryEntry de in table)
            {
                sw.Write(de.Key + "#" + de.Value + "\n");
            }
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();
        }

        //查询:返回wrong，表示没查到
        public String getPhonetic(String word)
        {
            this.workCount++;
            if (table.ContainsKey(word))//已经存在了
            {
                return table[word].ToString();
            }

            this.searchCount++;
            //向网络查询单词音标
            string serverUrl = @"http://fanyi.youdao.com/openapi.do?keyfrom=chen1233216&key=1817341544&type=data&doctype=json&version=1.1&q="
                + HttpUtility.UrlEncode(word);
            WebRequest request = WebRequest.Create(serverUrl);
            WebResponse response = request.GetResponse();
            string resJson = new StreamReader(response.GetResponseStream(), Encoding.UTF8).ReadToEnd();
            Regex r = new Regex("phonetic\":\"([^\"]+)\"");//用来解析的正则
            String result;
            if (r.IsMatch(resJson))//匹配到了
            {
                Match m = r.Match(resJson);
                if (m.Groups.Count < 2)
                {
                    result = "wrong";
                }
                else//查到啦
                {
                    result = m.Groups[1].Value;
                    table.Add(word, result);
                }
            }
            else
            {
                result = "wrong";
            }
            return result;
        }
    }

}
