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
using MessageBox = System.Windows.Forms.MessageBox;
using System.Windows.Xps.Packaging;
using Path = System.IO.Path;
using Microsoft.Office.Interop.Word;
using Window = System.Windows.Window;


namespace jbby
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
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
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            //创建＂打开文件＂对话框
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
            {
                //设置文件类型过滤
                Filter = "Word文档|*.docx|Word文档（95~2007）|*.doc|文本文档|*.txt;"
            };
            // 调用ShowDialog方法显示＂打开文件＂对话框
            Nullable<bool> result = dlg.ShowDialog();
            string filename = "";
            string text = "";
            string wjlxx = "";
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

            switch (wjlxx)
            {
                case ".docx":
                    try
                    {
                        
                        FileNR.Text = OpenWord(filename);
                    }
                    catch
                    {
                        MessageBox.Show(string.Format("文件{0}打开失败，错误：{1}", new string[] { filename, e.ToString() }));
                    }

                    break;
                case ".doc":
                    try
                    {

                        FileNR.Text = OpenWord(filename);
                    }
                    catch
                    {
                        MessageBox.Show(string.Format("文件{0}打开失败，错误：{1}", new string[] { filename, e.ToString() }));
                    }

                    break;
                case ".txt":
                    text = File.ReadAllText(filename);
                    FileNR.Text = text;
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
            //Settinges.Columns.Clear();
        }
        /// <summary>
        /// 添加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            this.Settinges.CanUserAddRows = true;//显示新行


        }
        /// <summary>
        /// 应用设置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
           var data= Settinges.Columns.AsQueryable();
            Setting setting = new Setting();
            //var date=Sett
            //setting.EndFH=
        }
       /// <summary>
       /// 编辑设置
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            Settinges.Columns.RemoveAt(Settinges.SelectedIndex);
            Settinges.Columns.Add(Settinges.Columns.FirstOrDefault());
        }
       /// <summary>
       /// 删除
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            var ss = MessageBox.Show("是否删除？", "提示", (MessageBoxButtons)MessageBoxButton.OKCancel, (MessageBoxIcon)MessageBoxImage.Question);
            if (ss == System.Windows.Forms.DialogResult.OK)
            {
                Settinges.Columns.Remove(Settinges.Columns.First());
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

        public string OpenWord(string fileName)
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

                doc.ActiveWindow.Selection.WholeStory();//全选word文档中的数据
                doc.ActiveWindow.Selection.Copy();//复制数据到剪切板
                return doc.ActiveWindow.Selection.Text;//richTextBox粘贴数据
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







    }
}
