using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace jbby
{
    public class Setting
    {
        public string StaerFH { get; set; }
        public string EndFH { get; set; }
        public string ZTColore { get; set; }
        public string Type { get; set; }
        public string JSName { get; set; }
        public string Nr { get; set; } = null;
    }

    public class ReadTxtFileLine
    {    //本类用于使用StreamReader.Read()方法，实现逐行读取文本文件，
        int _IsReadEnd = 0;  //文件读取的状态，当为false时，代表未读完最后一行，true为读完了最后一行
        System.IO.StreamReader sr1;
        int _LoopRowNumNow = 0;
        //定义了一个是否读到最后的属性，数据类型为整数
        public int IsReadEnd { get => _IsReadEnd; }
        //构造函数
        public ReadTxtFileLine(string TxtFilePath, Encoding FileEncoding)
        {
            sr1 = new System.IO.StreamReader(TxtFilePath, FileEncoding);
            _IsReadEnd = 1;
        }
        //成员方法，执行一次，返回1行的结果，当全部读完，依然执行该方法，将返回空字符串""
        public string GetLineStr()
        {
            string strLine = "";
            int charCode = 0;
            while (sr1.Peek() > 0)
            {
                charCode = sr1.Read();
                if (charCode == 10)  //发现换行符char10就返回拼接字符串
                {
                    _LoopRowNumNow++;
                    return strLine;
                }
                else
                {
                    if (charCode != 13)
                    {    //将一行的数据重新拼接起来
                        strLine += ((char)charCode).ToString();
                    }
                }
            }
            _IsReadEnd = -1;
            sr1.Close();
            sr1.Dispose();
            return strLine;
        }
    }

}
