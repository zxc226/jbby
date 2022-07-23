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
                int staerindex = dlg.SafeFileName.Trim().Length-1;
                int endindex = dlg.SafeFileName.IndexOf(".");
                wjlxx = dlg.SafeFileName.Substring(endindex, staerindex-1);
                wjlx.Content = wjlxx;
            }

            switch (wjlxx)
            {
                case ".docx":
                    
                    break;
                case ".doc":
                    Stream wordFile = File.OpenRead(filename);
                    XWPFDocument doc = new XWPFDocument(wordFile);
                    foreach (var para in doc.Paragraphs)
                    {
                        text = para.ParagraphText; //获得文本
                        if (text.Trim() != "")
                            FileNR.Text = text;
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

        }
        /// <summary>
        /// 添加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {

        }
       /// <summary>
       /// 应用设置
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {

        }
       /// <summary>
       /// 编辑设置
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
        private void Button_Click_6(object sender, RoutedEventArgs e)
        {

        }
       /// <summary>
       /// 删除
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("是否删除？", "提示", MessageBoxButton.OKCancel, MessageBoxImage.Question);
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
    }
}
