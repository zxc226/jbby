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
using System.IO;
using NPOI.XWPF;
using NPOI.XWPF.UserModel;
using NPOI.Util;
using System.Windows.Forms;
using MessageBox = Panuon.WPF.UI.MessageBoxX;
using System.Windows.Xps.Packaging;
using Path = System.IO.Path;
using Microsoft.Office.Interop.Word;
using Window = System.Windows.Window;
using System.Threading;
using System.Collections.ObjectModel;
using Panuon.WPF.UI;
using MessageBoxIcon = Panuon.WPF.UI.MessageBoxIcon;

namespace jbby
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : WindowX
    {
        ObservableCollection<Setting> settings = new ObservableCollection<Setting>();
        List<Setting> scnr = new List<Setting>();
        static int num = 0;
        public string filename = "";
        public string wjlxx = "";
        public string text = "";
        public MainWindow()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        /// <summary>
        /// 生成
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }
        /// <summary>
        /// 选择文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Button_Click_2(object sender, RoutedEventArgs e)
        {
            jdt.Value = 50;
            //创建＂打开文件＂对话框
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
            {
                //设置文件类型过滤
                Filter = "Word文档|*.docx|Word文档（95~2007）|*.doc|文本文档|*.txt;"
            };
            // 调用ShowDialog方法显示＂打开文件＂对话框
            Nullable<bool> result = dlg.ShowDialog();

           

            if (result == true)
            {
                //获取所选文件名并在FileNameTextBox中显示完整路径
                filename = dlg.FileName;
                wjdx.Content = dlg.OpenFile().Length;
                //int staerindex = dlg.SafeFileName.Trim().Length-1;
                int endindex = dlg.SafeFileName.Trim().IndexOf(".");
                wjlxx = dlg.SafeFileName.Substring(endindex);
                wjlx.Content = wjlxx;
            }
            else
            {
                jdt.Value = 0;
            }
            switch (wjlxx)
            {
                case ".docx":
                    try
                    {
                        Thread thread = new Thread(() => OpenWord(filename));
                        thread.Start();
                        var date= await OpenWord(filename);
                        for (int i = 0; i < date.Count; i++)
                        {
                            text += date[i] + "\t\n";
                        }
                        FileNR.Text = text;
                        jdt.Value = 100;
                    }
                    catch
                    {
                        MessageBox.Show(string.Format("文件{0}打开失败，错误：{1}", new string[] { filename, e.ToString() }));
                    }

                    break;
                case ".doc":
                    try
                    {
                        Thread thread = new Thread(() => OpenWord(filename));
                        thread.Start();
                        var date = await OpenWord(filename);
                        for (int i = 0; i < date.Count; i++)
                        {
                            text += date[i] + "\t\n";
                        }
                        FileNR.Text = text;
                        jdt.Value = 100;
                    }
                    catch
                    {
                        MessageBox.Show(string.Format("文件{0}打开失败，错误：{1}", new string[] { filename, e.ToString() }));
                    }

                    break;
                case ".txt":
                    text = File.ReadAllText(filename);
                    FileNR.Text = text;
                    jdt.Value = 100;
                    break;
                default:
                    break;
            }


        }
        /// <summary>
        /// 重置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            FileNR.Text = "";
            FileSC.Text = "";
            jdt.Value = 0;
            wjdx.Content = "";
            wjlx.Content = "";
            //Settinges.Columns.Clear();
        }
        /// <summary>
        /// 添加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            settings.Add(new Setting());
        }
        /// <summary>
        /// 应用设置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            var nrsc = "";
            var data = settings.AsQueryable();
            List<Setting> setting = data.ToList();
            Setting scnrs = new Setting();
            string path = filename;
            if (wjlxx != ".txt" || wjlxx != "txt")
            {
                var temp = System.AppDomain.CurrentDomain.BaseDirectory;
                temp = temp + "\\cs\\temp.txt";
                if (File.Exists(temp))
                {
                    File.Delete(temp);
                    FileStream aFile = new FileStream(temp, FileMode.OpenOrCreate);
                    StreamWriter sw = new StreamWriter(aFile);
                    sw.WriteLine(text);
                    sw.Close();
                    Encoding encoding = UTF8Encoding.UTF8;
                    ReadTxtFileLine ReadTxtFileTest1 = new ReadTxtFileLine(temp, encoding);
                    //while (ReadTxtFileTest1.IsReadEnd > 0)
                    //{
                    //    string str = ReadTxtFileTest1.GetLineStr();  //这里将读出来的1行赋值给str
                    //    MessageBox.Show(str, "文件读取内容", (MessageBoxButton)MessageBoxButtons.OK);
                    //}
                    while (ReadTxtFileTest1.IsReadEnd > 0)
                    {
                        string str = ReadTxtFileTest1.GetLineStr();  //这里将读出来的1行赋值给str
                        scnrs = JQstring(str);
                        if (str == "false")
                        {
                            continue;
                        }
                        else
                        {
                            scnr.Add(scnrs);
                        }
                    }

                    for (int i = 0; i < scnr.Count; i++)
                    {
                        try
                        {
                            if (scnr[i]==null)
                            {
                                continue;
                            }
                            else
                            {
                                nrsc = "Type:" + scnr[i].Type + ",\t\n";
                                nrsc += "ZTColore:" + scnr[i].ZTColore + ",\t\n";
                                nrsc += "JSName:" + scnr[i].JSName + ",\t\n";
                                nrsc += "Nr:"+scnr[i].Nr;
                                nrsc = "{\t\n" + nrsc + "\t\n}\t\n";
                            }
                            
                        }
                        catch (Exception ee)
                        {
                            Console.WriteLine(ee);
                            continue;
                        }
                        
                    }
                    FileSC.Text = nrsc;
                }
                else
                {
                    FileStream aFile = new FileStream(temp, FileMode.OpenOrCreate);
                    StreamWriter sw = new StreamWriter(aFile);
                    sw.WriteLine(text);
                    sw.Close();
                    Encoding encoding = UTF8Encoding.UTF8;
                    ReadTxtFileLine ReadTxtFileTest1 = new ReadTxtFileLine(temp, encoding);
                    //while (ReadTxtFileTest1.IsReadEnd > 0)
                    //{
                    //    string str = ReadTxtFileTest1.GetLineStr();  //这里将读出来的1行赋值给str
                    //    MessageBox.Show(str, "文件读取内容", (MessageBoxButton)MessageBoxButtons.OK);
                    //}
                    while (ReadTxtFileTest1.IsReadEnd > 0)
                    {
                        string str = ReadTxtFileTest1.GetLineStr();  //这里将读出来的1行赋值给str
                        scnrs = JQstring(str);
                        if (str == "false")
                        {
                            continue;
                        }
                        else
                        {
                            scnr.Add(scnrs);
                        }
                    }

                    for (int i = 0; i < scnr.Count; i++)
                    {
                        try
                        {
                            nrsc = scnr[i].Type + "\t\n";
                            nrsc += scnr[i].ZTColore + "\t\n";
                            nrsc += scnr[i].JSName + "\t\n";
                            nrsc += scnr[i].Nr + "\t\n";
                            nrsc = "{\t\n" + nrsc + "\t\n}\t\n";
                        }
                        catch (Exception ee)
                        {
                            Console.WriteLine(ee);
                            continue;
                        }
                    }
                    FileSC.Text = nrsc;
                }
               
            }
            else
            {
                
                Encoding encoding = UTF8Encoding.UTF8;
                ReadTxtFileLine ReadTxtFileTest1 = new ReadTxtFileLine(path, encoding);
                while (ReadTxtFileTest1.IsReadEnd > 0)
                {
                    string str = ReadTxtFileTest1.GetLineStr();  //这里将读出来的1行赋值给str
                    scnrs = JQstring(str);
                    if (str=="false")
                    {
                        continue;
                    }
                    else
                    {
                        scnr.Add(scnrs);
                    }
                }

                for (int i = 0; i < scnr.Count; i++)
                {
                    nrsc = scnr[i].Type+"\t\n";
                    nrsc += scnr[i].ZTColore+"\t\n";
                    nrsc += scnr[i].JSName+"\t\n";
                    nrsc += scnr[i].Nr+"\t\n";
                    nrsc= "{\t\n" + nrsc+ "\t\n}\t\n";
                }
                FileSC.Text = nrsc;
            }


        }
        /// <summary>
        /// 编辑设置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_6(object sender, RoutedEventArgs e)
        {

            Settinges.IsReadOnly = true;
            ++num;
            if (num % 2 == 0)
            {
                Settinges.IsReadOnly = false;
            }
            else
            {
                Settinges.IsReadOnly = true;
            }

        }
        /// <summary>
        /// 删除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            var ss = MessageBox.Show("是否删除？", "提示", MessageBoxButton.OKCancel, MessageBoxIcon.Question);
            if (ss == MessageBoxResult.OK)
            {
                if (Settinges.SelectedIndex < 0)
                {
                    MessageBox.Show("请选择要删除的数据！", "警告", MessageBoxButton.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    settings.RemoveAt(Settinges.SelectedIndex);
                }

            }


        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_8(object sender, RoutedEventArgs e)
        {

        }

        public static Task<List<string>> OpenWord(string fileName)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();//可以打开word
            Microsoft.Office.Interop.Word.Document doc = null;      //需要记录打开的word

            object missing = System.Reflection.Missing.Value;
            object File = fileName;
            object readOnly = false;//不是只读
            object isVisible = true;

            object unknow = Type.Missing;

            try
            {
                doc = app.Documents.Open(ref File, ref missing, ref readOnly,
                 ref missing, ref missing, ref missing, ref missing, ref missing,
                 ref missing, ref missing, ref missing, ref isVisible, ref missing,
                 ref missing, ref missing, ref missing);

                //doc.ActiveWindow.Selection.WholeStory();//全选word文档中的数据
                //doc.ActiveWindow.Selection.Copy();//复制数据到剪切板
                String read = string.Empty;
                List<string> data = new List<string>();
                for (int i = 0; i < doc.Paragraphs.Count; i++)
                {
                    string temp = doc.Paragraphs[i + 1].Range.Text+"\t\n";
                    if (temp != string.Empty)
                        data.Add(temp);
                }
                //var textes = doc.ActiveWindow.Selection.Text;
                return System.Threading.Tasks.Task.Run(() =>
                {
                    return data;
                });//richTextBox粘贴数据
                   //richTextBox1.Text = doc.Content.Text;//显示无格式数据
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(ref missing, ref missing, ref missing);
                    doc = null;
                }

                if (app != null)
                {
                    app.Quit(ref missing, ref missing, ref missing);
                    app = null;
                }
            }
        }

        private void Settinges_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void Settinges_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            var uie = e.OriginalSource as UIElement;
            if (e.Key == Key.Enter)
            {
                uie.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                e.Handled = true;
            }
        }

        private void Settinges_Loaded_1(object sender, RoutedEventArgs e)
        {
            settings.Add(new Setting());
            Settinges.ItemsSource = settings;
        }

        private void jdt_Loaded(object sender, RoutedEventArgs e)
        {
            jdt.Minimum = 0;
            jdt.Maximum = 100;

        }

        private void jdt_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }

        public Setting JQstring(string temp)
        {
            Setting setting = new Setting();
            if (temp==null||temp=="")
            {
                for (int i = 0; i < settings.Count; i++)
                {
                    try
                    {
                        var staerindex = settings[i].StaerFH;
                        var endindex = settings[i].EndFH;

                        int ks = temp.IndexOf(staerindex);
                        int js = temp.LastIndexOf(endindex);
                        var jg = temp.Substring(ks, js - ks);
                        
                        setting.Nr = jg;
                        setting.JSName = settings[i].JSName;
                        setting.ZTColore = settings[i].ZTColore;
                        setting.Type = settings[i].Type;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                        continue;
                    }


                }
                return setting;
            }
            else
            {
                return null;
            }
            
        }


    }
}
