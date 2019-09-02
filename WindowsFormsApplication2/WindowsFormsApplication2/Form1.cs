using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApplication2.DAL;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Threading;


namespace WindowsFormsApplication2
{
    public delegate DataTable GetExcel(string path);
    
    public partial class Form1 : Form
    {
        public static GetExcel getExcel = new GetExcel(ReadExcelToTable);
        public static  DataTable dt = new DataTable();
        public Form1()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        public static DataTable ReadExcelToTable(string path)
        {
            DataTable set = new DataTable();
            DataTable table = new DataTable();
            string strFileType = System.IO.Path.GetExtension(path);
            //string connstring = null;
            //if (strFileType == ".xls")
            //{
            //    connstring = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            //}
            //else
            //{
            //    connstring = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            //}
            //try {
            //    using (OleDbConnection conn = new OleDbConnection(connstring))
            //    {

            //        conn.Open();
            //        DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
            //        string firstSheetName = sheetsName.Rows[0][2].ToString();
            //        string sql = string.Format("SELECT [Host Asset ID],[Publish URL] FROM [{0}] WHERE [Host Asset ID] not like '%toc.%'", firstSheetName);
            //        OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
            //        ada.Fill(set);
            string msg;
            int a = 0;
            int b = 0;
            set = Unilever.Common.NPOI_Office.NPOI_Excel.ImportExcel.ImportExceltoDt(path, 0, -1, out msg);
            List<Info> List = new List<Info>();
            for (int j = 0; j < set.Columns.Count; j++)
            {

                if (set.Rows[0][j].ToString() == "Host Asset ID")
                {
                    a = j;
                }
                if (set.Rows[0][j].ToString() == "Publish Url")
                {
                    b = j;
                }
            }
            if (a != 0 && b != 0)
            {
                for (int i = 1; i < set.Rows.Count; i++)
                {
                    Info info = new Info();
                    string str = set.Rows[i][b].ToString();
                    string strr = set.Rows[i][a].ToString();
                    if (strr.Contains("TOC.") || strr.Contains("toc."))
                    {

                    } else
                    {
                        if ((strr != null && strr !="")|| (str != null&& str !=""))
                        {
                            info.FileName = strr;
                            info.PublishURL = str.Replace("review.", "").Replace("?branch=live-sxs", "").Replace("&branch=live-sxs", "");
                            List.Add(info);
                        }
                    }

                }
            }
            else
            {
                ShowYesNoAndError("没有找到指定列");
            }
            table = ToDataTableTow(List);


            //    }
            //} catch (Exception e)
            //{

            //}
            return table;
        }

        public static DataTable ReadTxtToTable(string path1, string path)
        {
            DataTable set = new DataTable();
            DataTable table = new DataTable();
            string idpath = string.Format("{0}\\_Overall.log.page.error.rerun.txt", path);
            StreamReader str = new StreamReader(idpath, Encoding.GetEncoding("GB2312"));
            //string txt = sr.ReadToEnd().Replace("\r\n", "-");
            string txt = str.ReadToEnd();
            string result = System.Text.RegularExpressions.Regex.Replace(txt, @"[^0-9]+", "-");
            string[] nodes = result.Split('-');
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(path);
            FileInfo[] ff = di.GetFiles("ID Group*.txt");
            List<string> list = new List<string>();
            foreach (FileInfo temp in ff)
            {
                using (StreamReader sr = temp.OpenText())
                {
                    string srtxt = sr.ReadToEnd();
                    string res = System.Text.RegularExpressions.Regex.Replace(srtxt, @"[^0-9]+", "-");
                    string[] grid = res.Split('-');
                    for (int i = 1; i < grid.Length - 1; i++)
                    {
                        if (nodes.Contains(grid[i]))
                        {
                            string ur = Search_string(srtxt.Substring(srtxt.IndexOf(grid[i])), grid[i], "_").Trim();
                            list.Add(ur);

                        }
                    }
                }
            }
            //string strFileType = System.IO.Path.GetExtension(path1);
            //string connstring = null;
            //if (strFileType == ".xls")
            //{
            //    connstring = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + path1 + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            //}
            //else
            //{
            //    connstring = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path1 + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            //}
            //try
            //{
            //    using (OleDbConnection conn = new OleDbConnection(connstring))
            //    {

            //        conn.Open();
            //        DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
            //        string firstSheetName = sheetsName.Rows[0][2].ToString();
            //        string sql = string.Format("SELECT * FROM [{0}]", firstSheetName);
            //        OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
            //        ada.Fill(set);
            string msg;
            set = Unilever.Common.NPOI_Office.NPOI_Excel.ImportExcel.ImportExceltoDt(path1, 0, -1, out msg);
            List<Info> List = new List<Info>();

            for (int i = 1; i < set.Rows.Count; i++)
            {
                Info info = new Info();
                if (!set.Rows[i][1].ToString().Equals("Not Set")) {
                    string strr = set.Rows[i][1].ToString().Substring(33);
                    if (list.Contains(strr))
                    {
                        info.FileName = set.Rows[i][0].ToString();
                        info.PublishURL = set.Rows[i][1].ToString();
                        List.Add(info);
                    }
                }


            }
            table = ToDataTableTow(List);
            //}
            //}
            //catch (Exception e)
            //{

            //}   
            return table;
        }
        public static DataTable AddLanguege(string path)
        {
            string msg;
            DataTable set = new DataTable();
            DataTable table = new DataTable();
            set = Unilever.Common.NPOI_Office.NPOI_Excel.ImportExcel.ImportExceltoDt(path, 0, -1, out msg);
            List<Language> List = new List<Language>();
            for (int i = 1; i < set.Rows.Count; i++)
            {
                string language = set.Rows[i][1].ToString();
                string[] lan = language.Split(',');
                StringBuilder MyStringBuilder = new StringBuilder();
                if (set.Rows[i][0].ToString().Contains("review"))
                {
                    string Path1 = set.Rows[i][0].ToString().Substring(40);
                    for (int j = 0; j < lan.Length; j++)
                    {
                        Language lang = new Language();
                        string a = lan[j].Trim();
                        string val = string.Format("https://review.docs.microsoft.com/{0}/{1}", a, Path1);
                        MyStringBuilder.Append(val);
                        MyStringBuilder.Append("\r\n");
                        lang.Path = val;
                        lang.Languages = val;
                        List.Add(lang);

                    }
                }
                else
                {
                    string Path1 = set.Rows[i][0].ToString().Substring(33);


                    for (int j = 0; j < lan.Length; j++)
                    {
                        Language lang = new Language();
                        string a = lan[j].Trim();
                        string val = string.Format("https://docs.microsoft.com/{0}/{1}", a, Path1);
                        MyStringBuilder.Append(val);
                        MyStringBuilder.Append("\r\n");
                        lang.Path = val;
                        lang.Languages = val;
                        List.Add(lang);

                    }
                }
                
            }
            table = ToDataTableTow(List);
            return table;
        }
        public static DataTable SplitTable(string path,string ra)
        {
            double rate = 1;
            if (ra != "")
            {
                rate = double.Parse(ra);
            }
            string msg;
            DataTable set = new DataTable();
            DataTable table = new DataTable();
            set = Unilever.Common.NPOI_Office.NPOI_Excel.ImportExcel.ImportExceltoDt(path, 0, -1, out msg);
            List<Dic> list = new List<Dic>();
            List<LanTemp> list2 = new List<LanTemp>();
            List<Info> list3 = new List<Info>();
            for (int i = 1; i < set.Rows.Count; i++)
            {
                string str = set.Rows[i][1].ToString();
                string str1 = str.Substring(32).ToString();
                string str2 = str.Substring(27, 5).ToString();
                Dic dic = new Dic();

                dic.Lan = str2;
                dic.URL = str1;
                list.Add(dic);

            }
            int ListCount = list.Count();
            var re = list.GroupBy(p => new { p.URL }).Select(g => g.First()).ToList();
            var language = from a in list select a.Lan;
            var language2 = language.Distinct();
            int lanCount = language2.Count();
            var urls = from c in re select c.URL;
            foreach (var url in urls)
            {
                LanTemp lantemp = new LanTemp();
                var lang = from c in list where c.URL == url select c.Lan;
                string result = string.Join(",", lang.ToArray());
                lantemp.Lang = result;
                lantemp.PublishURL = url;
                list2.Add(lantemp);
            }
            foreach (var url in list2)
            {
                string[] lan=url.Lang.Split(',');
                int count = lan.Length;
                double a = count * rate;
                int lancount = Convert.ToInt32((lanCount * rate)-1);
                int b = Convert.ToInt32(a);
                if (b < 0)
                {
                    b = 1;
                }
                if (count > lancount)
                {
                    List<string> lanlist = new List<string>();
                    Random rand;
                    int i = 0;
                    while (i < b)
                    {
                        int index = 0;
                        for (int ii = 0; ii < count; ii++)
                        {
                            rand = new Random((unchecked((int)DateTime.Now.Ticks + ii)));
                            index = rand.Next(0, count);
                        }
                        Info info = new Info();
                        info.FileName=String.Format("{0}{1}{2}", "https://docs.microsoft.com/", lan[index], url.PublishURL);
                        info.PublishURL = String.Format("{0}{1}{2}", "https://docs.microsoft.com/", lan[index], url.PublishURL);
                        list3.Add(info);
                        List<string> lann = lan.ToList();
                        lann.Remove(lan[index]);
                        lan = lann.ToArray();
                        count = count - 1;
                        i++;
                    }
                }else
                {
                    for(int i = 0; i < count; i++)
                    {
                        Info info = new Info();
                        info.FileName = String.Format("{0}{1}{2}", "https://docs.microsoft.com/", lan[i], url.PublishURL);
                        info.PublishURL = String.Format("{0}{1}{2}", "https://docs.microsoft.com/", lan[i], url.PublishURL);
                        list3.Add(info);
                    }
                }
            }
            table = ToDataTableTow(list3);
            return table;

        }
        public static DataTable SplitTable2(string path, string ra)
        {
            double rate = 1;
            if (ra != "")
            {
                rate = double.Parse(ra);
            }
            string msg;
            DataTable set = new DataTable();
            DataTable table = new DataTable();
            set = Unilever.Common.NPOI_Office.NPOI_Excel.ImportExcel.ImportExceltoDt(path, 0, -1, out msg);
            List<Dic> list = new List<Dic>();
            List<LanTemp> list2 = new List<LanTemp>();
            List<Info> list3 = new List<Info>();
            for (int i = 1; i < set.Rows.Count; i++)
            {
                for(int j = 2; j < set.Columns.Count; j++)
                {
                    if (!set.Rows[i][j].ToString().Equals(""))
                    {
                        string str1 = set.Rows[i][1].ToString();
                        string str2 = set.Rows[i][j].ToString();
                        Dic dic = new Dic();
                        dic.Lan = str2;
                        dic.URL = str1;
                        list.Add(dic);
                    }
                   
                }
            }
            int ListCount = list.Count();
            var re = list.GroupBy(p => new { p.URL }).Select(g => g.First()).ToList();
            var language = from a in list select a.Lan;
            var language2 = language.Distinct();
            int lanCount = language2.Count();
            var urls = from c in re select c.URL;
            foreach (var url in urls)
            {
                LanTemp lantemp = new LanTemp();
                var lang = from c in list where c.URL == url select c.Lan;
                string result = string.Join(",", lang.ToArray());
                lantemp.Lang = result;
                lantemp.PublishURL = url;
                list2.Add(lantemp);
            }
            foreach (var url in list2)
            {
                string[] lan = url.Lang.Split(',');
                int count = lan.Length;
                double a = count * rate;
                int lancount = Convert.ToInt32((lanCount * rate) - 1);
                int b = Convert.ToInt32(a);
                if (b < 0)
                {
                    b = 1;
                }
                if (count > lancount)
                {
                    List<string> lanlist = new List<string>();
                    Random rand;
                    int i = 0;
                    while (i < b)
                    {
                        int index = 0;
                        for (int ii = 0; ii < count; ii++)
                        {
                            rand = new Random((unchecked((int)DateTime.Now.Ticks + ii)));
                            index = rand.Next(0, count);
                        }
                        Info info = new Info();
                        info.FileName = String.Format("{0}{1}/{2}", "https://docs.microsoft.com/", lan[index],url.PublishURL);
                        info.PublishURL = String.Format("{0}{1}/{2}", "https://docs.microsoft.com/", lan[index],url.PublishURL);
                        list3.Add(info);
                        List<string> lann = lan.ToList();
                        lann.Remove(lan[index]);
                        lan = lann.ToArray();
                        count = count - 1;
                        i++;
                    }
                }
                else
                {
                    for (int i = 0; i < count; i++)
                    {
                        Info info = new Info();
                        info.FileName = String.Format("{0}{1}/{2}", "https://docs.microsoft.com/", lan[i], url.PublishURL);
                        info.PublishURL = String.Format("{0}{1}/{2}", "https://docs.microsoft.com/", lan[i], url.PublishURL);
                        list3.Add(info);
                    }
                }
            }
            table = ToDataTableTow(list3);
            return table;

        }
        //public static IEnumerable<TSource> DistinctBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        //{
        //    HashSet<TKey> seenKeys = new HashSet<TKey>();
        //    foreach (TSource element in source)
        //    {
        //        if (seenKeys.Add(keySelector(element)))
        //        {
        //            yield return element;
        //        }
        //    }
        //} 
        public static DataTable GetOtherUrl(string path, string path1)
        {
            string msg;
            DataTable set = new DataTable();
            DataTable set2 = new DataTable();
            DataTable table = new DataTable();
            set = Unilever.Common.NPOI_Office.NPOI_Excel.ImportExcel.ImportExceltoDt(path, 0, -1, out msg);
            List<Info> List = new List<Info>();
            List<Info> List2 = new List<Info>();
            List<Dic> List3 = new List<Dic>();
            for (int i = 1; i < set.Rows.Count; i++)
            {
                string language = set.Rows[i][1].ToString();
                string name = set.Rows[i][0].ToString().ToLower();
                string url = set.Rows[i][1].ToString().ToLower();
                Info info = new Info();
                info.FileName = name;
                info.PublishURL = url;
                List.Add(info);
            }
            set2 = Unilever.Common.NPOI_Office.NPOI_Excel.ImportExcel.ImportExceltoDt(path1, 0, -1, out msg);
            for(int i = 1; i < set2.Rows.Count;i++)
            {
                
                string ID = set2.Rows[i][0].ToString().Substring(33).ToLower();
                string Lan = set2.Rows[i][0].ToString().Substring(27, 5).ToLower();
                Dic dic = new Dic();
                dic.Lan = Lan;
                dic.URL = ID;
                List3.Add(dic);
            }
            var list=List3.GroupBy(a => a.URL).Select(a => a.First());
            foreach( var li in list)
            {
                string url=li.URL;
                var la = from z in List3 where z.URL.Equals(url) select z.Lan;
                var other = from s in List where s.PublishURL.Substring(33).Equals(url) select s;
                foreach (var oth in other)
                {
                    if (!la.Contains(oth.PublishURL.Substring(27, 5)))
                    {
                        Info info2 = new Info();
                        info2.FileName = oth.FileName;
                        info2.PublishURL = oth.PublishURL;
                        List2.Add(info2);
                    }
                }
            }
           
            table = ToDataTableTow(List2);
            return table;
        }
        
        public static DataTable Sort (string path)
        {
            List<Info> List = new List<Info>();
            string msg;
            DataTable set = new DataTable();
            DataTable table = new DataTable();
            set = Unilever.Common.NPOI_Office.NPOI_Excel.ImportExcel.ImportExceltoDt(path, 0, -1, out msg);
            for (int i = 0; i < set.Rows.Count; i++)
            {
                for(int j=0; j < set.Columns.Count; j++)
                {
                    string str = set.Rows[i][j].ToString();
                    if (str != null&& str != "")
                    {
                        Info info = new Info();
                        info.FileName = str;
                        info.PublishURL = str;
                        List.Add(info);
                    }
                }

            }
            table = ToDataTableTow(List);
            return table;
        }

        public static DialogResult ShowYesNoAndError(string message)
        {
            return MessageBox.Show(message, "错误信息", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
        }
        public static string Search_string(string s, string s1, string s2)  //获取搜索到的数目  
        {
            int n1, n2;
            n1 = s.IndexOf(s1, 0) + s1.Length;   //开始位置  
            n2 = s.IndexOf(s2, n1);               //结束位置    
            return s.Substring(n1, n2 - n1);   //取搜索的条数，用结束的位置-开始的位置,并返回    
        }
        public static DataTable ToDataTableTow(IList list)
        {
            DataTable result = new DataTable();
            if (list.Count > 0)
            {
                PropertyInfo[] propertys = list[0].GetType().GetProperties();

                //foreach (PropertyInfo pi in propertys)
                //{
                //    result.Columns.Add(pi.Name, pi.PropertyType);
                //}
                result.Columns.Add("File Name");
                result.Columns.Add("Publish URL");
                foreach (object t in list)
                {
                    ArrayList tempList = new ArrayList();
                    foreach (PropertyInfo pi in propertys)
                    {
                        object obj = pi.GetValue(t, null);
                        tempList.Add(obj);
                    }
                    object[] array = tempList.ToArray();
                    result.LoadDataRow(array, true);
                }
            }
            return result;
        }

        void TaskFinished(IAsyncResult result)
        {
            DataTable da = getExcel.EndInvoke(result);
            dt = da;
            this.dataGridView1.DataSource = dt;
            label1.Text = dt.Rows.Count.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dt.Clear();
            dataGridView1.DataSource = null;
            string folder_path = this.textBox1.Text.Trim();
            //getExcel.BeginInvoke(folder_path, new AsyncCallback(TaskFinished), null);
            
            dt = ReadExcelToTable(folder_path);
            dataGridView1.DataSource = dt;
            label1.Text = dt.Rows.Count.ToString();

        }

        private void button2_Click(object sender, EventArgs e)
        {

            string folder_path = this.textBox1.Text.Trim();
            string pa = folder_path.Substring(0,folder_path.LastIndexOf('\\') + 1);
            string path = string.Format("{0}localization-status.xlsx", folder_path.Substring(0,folder_path.LastIndexOf('\\')+1));
            //DataChangeExcel.DataSetToExcel(dt, path);
            DataChangeExcel.TableToExcel(dt, path);


        }

        private void button3_Click(object sender, EventArgs e)
        {
            string value = this.textBox1.Text.Trim();
            string[] lan = null;
            if (value.Contains(",")) {
                lan = value.Split(',');
            }else
            {
                lan = value.Split('	');
            }
            
            StringBuilder MyStringBuilder = new StringBuilder();
            for (int i = 0; i < lan.Length; i++)
            {
                string a = lan[i].Trim();
                string val = string.Format("<target culture =\"{0}\"/>",a);
                MyStringBuilder.Append(val);
                MyStringBuilder.Append("\r\n");
            }
            this.textBox2.Text = MyStringBuilder.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Rerun代码
            //dt.Clear();
            //dataGridView1.DataSource = null;
            //string folder_path = this.textBox1.Text.Trim();
            //string folder_pathre = this.textBox3.Text.Trim();
            //dt = ReadTxtToTable(folder_path, folder_pathre);
            //dataGridView1.DataSource = dt;
            //dt.Clear();
            //dataGridView1.DataSource = null;
            //string folder_path = this.textBox1.Text.Trim();
            //dt = Sort(folder_path);
            //dataGridView1.DataSource = dt;
            //label1.Text = dt.Rows.Count.ToString();
            dt.Clear();
            dataGridView1.DataSource = null;
            string path1 = this.textBox1.Text.Trim().ToString();
            string path2 = this.textBox3.Text.Trim().ToString();
            dt = GetOtherUrl(path1,path2);
            dataGridView1.DataSource = dt;
            label1.Text = dt.Rows.Count.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dt.Clear();
            dataGridView1.DataSource = null;
            string path = this.textBox1.Text.Trim();
            dt=AddLanguege(path);
            dataGridView1.DataSource = dt;
            label1.Text = dt.Rows.Count.ToString();


        }

        private void button6_Click(object sender, EventArgs e)
        {
            dt.Clear();
            string path = this.textBox1.Text.Trim().ToString();
            string  rate = this.textBox3.Text.Trim().ToString();
            dt = SplitTable(path,rate);
            dataGridView1.DataSource = dt;
            label1.Text = dt.Rows.Count.ToString();
        }
    }
}
