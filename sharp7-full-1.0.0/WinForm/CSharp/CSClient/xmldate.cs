using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;//XML处理类专属的头文件
using Sharp7;
using System.Data.SQLite;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections;
using System.Windows.Forms.DataVisualization.Charting;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Management;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Drawing.Drawing2D;

namespace CSClient
{
    public delegate void SetDateTime();
    public delegate void ToolTipInfo1();
    public delegate void ToolTipInfo2();
    public delegate void ToolTipInfo3();
    public delegate void ChartDelegate();
    public partial class xmldate : Form
    {
        int currentnumber;
        int totalnumber;
        int splitcounter;
        int picwidth;
        int picheight;
        //坐标位置的偏移量 注意绘图坐标系（原点在左下角）和图像坐标系（原点在左上角）的不同。
        int biasx;
        int biasy;
        //  public delegate void AddInfo_Delgegate(string message);
        double ft1;
        private static object lockojb = new object();
        byte[] bytearray;
        //根据配置生成的数组
        ArrayList itemarraylist = new ArrayList();
        S7MultiVar Reader;
        private static MainForm mainform1;
        private bool listFlag = false;
        private S7Client Client;
        private bool connectedFlag = false;
        string tooltipstr1;
        string tooltipstr2;
        char[] chararray = new char[255];
        string dbPath;
        string xml_FilePath = "";//用来记录当前打开文件的路径的
                                 //创建DataSet对象
        DataSet ds = new DataSet();
        //创建DataTable对象
        System.Data.DataTable dtg = new System.Data.DataTable();
        ArrayList shortarray = new ArrayList();
        ArrayList longarray = new ArrayList();
        System.Data.DataTable realtimedt;
        SQLiteConnection conn;
        //是否启用局部放大,FALSE--未启用

        public xmldate()
        {

            InitializeComponent();
            Client = new S7Client();
            if (IntPtr.Size == 4)
                this.Text = this.Text + " - Running 32 bit Code";
            else
                this.Text = this.Text + " - Running 64 bit Code";
            chararray = "abcdefg".ToCharArray();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = this.tabPage1;


            OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();//一个打开文件的对话框

            string str1 = " ";
            str1 = System.Threading.Thread.GetDomain().BaseDirectory + "config";
            if (Directory.Exists(str1) == false)//如果不存
            {
                Directory.CreateDirectory(str1);
            }
            openFileDialog1.InitialDirectory = str1;

            openFileDialog1.Filter = "xml文件(*.xml)|*.xml";//设置允许打开的扩展名
            if (openFileDialog1.ShowDialog() == DialogResult.OK)//判断是否选择了文件  
            {
                xml_FilePath = openFileDialog1.FileName;//记录用户选择的文件路径
                XmlDocument xmlDocument = new XmlDocument();//新建一个XML“编辑器”
                xmlDocument.Load(xml_FilePath);//载入路径这个xml
                try
                {
                    XmlNodeList xmlNodeList = xmlDocument.SelectSingleNode("单任务").ChildNodes;//选择class为根结点并得到旗下所有子节点
                    dataGridView1.Rows.Clear();//清空dataGridView1，防止和上次处理的数据混乱
                    foreach (XmlNode xmlNode in xmlNodeList)//遍历class的所有节点
                    {
                        XmlElement xmlElement = (XmlElement)xmlNode;//对于任何一个元素，其实就是每一个<student>
                        //旗下的子节点<name>和<number>分别放入dataGridView1
                        int index = dataGridView1.Rows.Add();//在dataGridView1新加一行，并拿到改行的行标
                        dataGridView1.Rows[index].Cells[0].Value = Convert.ToByte(xmlElement.ChildNodes.Item(0).InnerText);//各个单元格分别添加
                        dataGridView1.Rows[index].Cells[1].Value = Convert.ToInt32(xmlElement.ChildNodes.Item(1).InnerText);
                        dataGridView1.Rows[index].Cells[2].Value = xmlElement.ChildNodes.Item(2).InnerText;//各个单元格分别添加
                        dataGridView1.Rows[index].Cells[3].Value = xmlElement.ChildNodes.Item(3).InnerText;
                        dataGridView1.Rows[index].Cells[4].Value = xmlElement.ChildNodes.Item(4).InnerText;//各个单元格分别添加
                        dataGridView1.Rows[index].Cells[5].Value = xmlElement.ChildNodes.Item(5).InnerText;
                    }
                }
                catch
                {
                    MessageBox.Show("XML格式不对！");
                }
            }
            else
            {
                MessageBox.Show("请打开XML文件");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            XmlDocument xmlDocument = new XmlDocument();//新建一个XML“编辑器”
            if (false & xml_FilePath != "")//如果用户已读入xml文件，我们的任务就是修改这个xml文件了
            {
                xmlDocument.Load(xml_FilePath);
                XmlNode xmlElement_class = xmlDocument.SelectSingleNode("单任务");//找到<class>作为根节点
                xmlElement_class.RemoveAll();//删除旗下所有节点
                int row = dataGridView1.Rows.Count;//得到总行数    
                int cell = dataGridView1.Rows[1].Cells.Count;//得到总列数    
                for (int i = 0; i <= row - 1; i++)//遍历这个dataGridView
                {
                    XmlElement xmlElement_student = xmlDocument.CreateElement("变量列表");//创建一个<student>节点
                    XmlElement xmlElement_name = xmlDocument.CreateElement("区域");//创建<name>节点
                    xmlElement_name.InnerText = dataGridView1.Rows[i].Cells[0].Value.ToString();//其文本就是第0个单元格的内容
                    xmlElement_student.AppendChild(xmlElement_name);//在<student>下面添加一个新的节点<name>
                    //同理添加<number>
                    XmlElement xmlElement_number = xmlDocument.CreateElement("长度");
                    xmlElement_number.InnerText = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    xmlElement_student.AppendChild(xmlElement_number);
                    XmlElement xmlElement_db = xmlDocument.CreateElement("数据块号");
                    xmlElement_db.InnerText = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    xmlElement_student.AppendChild(xmlElement_db);

                    XmlElement xmlElement_start = xmlDocument.CreateElement("起始");
                    xmlElement_start.InnerText = dataGridView1.Rows[i].Cells[3].Value.ToString();
                    xmlElement_student.AppendChild(xmlElement_start);

                    XmlElement xmlElement_num = xmlDocument.CreateElement("数量");
                    xmlElement_num.InnerText = dataGridView1.Rows[i].Cells[4].Value.ToString();
                    xmlElement_student.AppendChild(xmlElement_num);

                    XmlElement xmlElement_name1 = xmlDocument.CreateElement("名称");
                    xmlElement_name1.InnerText = dataGridView1.Rows[i].Cells[5].Value.ToString();
                    xmlElement_student.AppendChild(xmlElement_name1);


                    xmlElement_class.AppendChild(xmlElement_student);//将这个<student>节点放到<class>下方
                }
                xmlDocument.Save(xml_FilePath);//保存这个xml
            }
            else//如果用户未读入xml文件，我们的任务就新建一个xml文件了
            {
                SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();//打开一个保存对话框
                string str1 = " ";
                str1 = System.Threading.Thread.GetDomain().BaseDirectory + "config";
                if (Directory.Exists(str1) == false)//如果不存
                {
                    Directory.CreateDirectory(str1);
                }
                saveFileDialog1.InitialDirectory = str1;
                saveFileDialog1.Filter = "xml文件(*.xml)|*.xml";//设置允许打开的扩展名
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)//判断是否选择了一个文件路径
                {
                    XmlElement xmlElement_class = xmlDocument.CreateElement("单任务");//创建一个<class>节点
                    int row = dataGridView1.Rows.Count;//得到总行数    
                    //int cell = dataGridView1.Rows[1].Cells.Count;//得到总列数    
                    for (int i = 0; i <= row - 1; i++)//得到总行数并在之内循环    
                    {
                        //同上，创建一个个<student>节点，并且附到<class>之下
                        XmlElement xmlElement_student = xmlDocument.CreateElement("变量列表");//创建一个<student>节点
                        XmlElement xmlElement_name = xmlDocument.CreateElement("区域");//创建<name>节点
                        xmlElement_name.InnerText = dataGridView1.Rows[i].Cells[0].Value.ToString();//其文本就是第0个单元格的内容
                        xmlElement_student.AppendChild(xmlElement_name);//在<student>下面添加一个新的节点<name>
                                                                        //同理添加<number>
                        XmlElement xmlElement_number = xmlDocument.CreateElement("长度");
                        xmlElement_number.InnerText = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        xmlElement_student.AppendChild(xmlElement_number);
                        XmlElement xmlElement_db = xmlDocument.CreateElement("数据块号");
                        xmlElement_db.InnerText = dataGridView1.Rows[i].Cells[2].Value.ToString();
                        xmlElement_student.AppendChild(xmlElement_db);

                        XmlElement xmlElement_start = xmlDocument.CreateElement("起始");
                        xmlElement_start.InnerText = dataGridView1.Rows[i].Cells[3].Value.ToString();
                        xmlElement_student.AppendChild(xmlElement_start);

                        XmlElement xmlElement_num = xmlDocument.CreateElement("数量");
                        xmlElement_num.InnerText = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        xmlElement_student.AppendChild(xmlElement_num);

                        XmlElement xmlElement_name1 = xmlDocument.CreateElement("名称");
                        xmlElement_name1.InnerText = dataGridView1.Rows[i].Cells[5].Value.ToString();
                        xmlElement_student.AppendChild(xmlElement_name1);
                        xmlElement_class.AppendChild(xmlElement_student);
                    }
                    xmlDocument.AppendChild(xmlDocument.CreateXmlDeclaration("1.0", "utf-8", ""));//编写文件头
                    xmlDocument.AppendChild(xmlElement_class);//将这个<class>附到总文件头，而且设置为根结点
                    xmlDocument.Save(saveFileDialog1.FileName);//保存这个xml文件
                }
                else
                {
                    MessageBox.Show("请保存为XML文件");
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void xmldate_Load(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = this.tabPage6;
            Releasehelp();

            Thread thread = new Thread(new ThreadStart(DateTimeInfo));
            thread.IsBackground = true;
            thread.Start();

            //   Thread thread1 = new Thread(new ThreadStart(tooltipInfo1));
            //   thread1.IsBackground = true;
            //    thread1.Start();
            //   Thread thread2 = new Thread(new ThreadStart(tooltipInfo2));
            //     thread2.IsBackground = true;
            //    thread2.Start();
            ///  Thread thread3 = new Thread(new ThreadStart(tooltipInfo3));
            //    thread3.IsBackground = true;
            //    thread3.Start();



            textBox1.Text = System.Threading.Thread.GetDomain().BaseDirectory + "database";
            dataGridView1.Columns.Clear();
            dataGridView1.AutoGenerateColumns = false;

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Clear();

            dt.Columns.Add("Area", Type.GetType("System.String"));
            dt.Columns.Add("Name", Type.GetType("System.Byte"));
            DataRow dr1 = dt.NewRow();
            dr1["Area"] = "S7AreaPE";
            dr1["Name"] = S7Consts.S7AreaPE;
            dt.Rows.Add(dr1);

            DataRow dr2 = dt.NewRow();
            dr2["Area"] = "S7AreaPA";
            dr2["Name"] = S7Consts.S7AreaPA;
            dt.Rows.Add(dr2);

            DataRow dr3 = dt.NewRow();
            dr3["Area"] = "S7AreaMK";
            dr3["Name"] = S7Consts.S7AreaMK;
            dt.Rows.Add(dr3);

            DataRow dr4 = dt.NewRow();
            dr4["Area"] = "S7AreaDB";
            dr4["Name"] = S7Consts.S7AreaDB;
            dt.Rows.Add(dr4);

            DataRow dr5 = dt.NewRow();
            dr5["Area"] = "S7AreaCT";
            dr5["Name"] = S7Consts.S7AreaCT;
            dt.Rows.Add(dr5);

            DataRow dr6 = dt.NewRow();
            dr6["Area"] = "S7AreaTM";
            dr6["Name"] = S7Consts.S7AreaTM;
            dt.Rows.Add(dr6);
            DataGridViewComboBoxColumn combox = new DataGridViewComboBoxColumn()
            {

                DataSource = dt,
                HeaderText = "区域",
                //DataGridView数据源中的列
                DataPropertyName = "Area",
                Name = "column0",
                ToolTipText = " Area identifier",
                Width = 120,
                //AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                // 必须和数据库字段名保持一致,
                DisplayMember = "Area",
                ValueMember = "Name",
            };
            dataGridView1.Columns.Insert(0, combox);
            System.Data.DataTable dtWL = new System.Data.DataTable();
            dtWL.Columns.Add("S7WL", Type.GetType("System.String"));
            dtWL.Columns.Add("Namevalue", Type.GetType("System.Int32"));

            DataRow drWL1 = dtWL.NewRow();
            drWL1[0] = "S7WLBit";
            drWL1[1] = S7Consts.S7WLBit;
            dtWL.Rows.Add(drWL1);

            DataRow drWL2 = dtWL.NewRow();
            drWL2[0] = "S7WLByte";
            drWL2[1] = S7Consts.S7WLByte;
            dtWL.Rows.Add(drWL2);

            DataRow drWL3 = dtWL.NewRow();
            drWL3[0] = "S7WLChar";
            drWL3[1] = S7Consts.S7WLChar;
            dtWL.Rows.Add(drWL3);

            DataRow drWL4 = dtWL.NewRow();
            drWL4[0] = "S7WLWord";
            drWL4[1] = S7Consts.S7WLWord;
            dtWL.Rows.Add(drWL4);

            DataRow drWL5 = dtWL.NewRow();
            drWL5[0] = "S7WLInt";
            drWL5[1] = S7Consts.S7WLInt;
            dtWL.Rows.Add(drWL5);

            DataRow drWL6 = dtWL.NewRow();
            drWL6[0] = "S7WLDWord";
            drWL6[1] = S7Consts.S7WLDWord;
            dtWL.Rows.Add(drWL6);

            DataRow drWL7 = dtWL.NewRow();
            drWL7[0] = "S7WLDInt";
            drWL7[1] = S7Consts.S7WLDInt;
            dtWL.Rows.Add(drWL7);

            DataRow drWL8 = dtWL.NewRow();
            drWL8[0] = "S7WLReal";
            drWL8[1] = S7Consts.S7WLReal;
            dtWL.Rows.Add(drWL8);

            DataRow drWL9 = dtWL.NewRow();
            drWL9[0] = "S7WLCounter";
            drWL9[1] = S7Consts.S7WLCounter;
            dtWL.Rows.Add(drWL9);

            DataRow drWL10 = dtWL.NewRow();
            drWL10[0] = "S7WLTimer";
            drWL10[1] = S7Consts.S7WLTimer;
            dtWL.Rows.Add(drWL10);


            DataGridViewComboBoxColumn column1 = new DataGridViewComboBoxColumn()
            {
                DataSource = dtWL,
                HeaderText = "长度",
                DataPropertyName = "S7WL",
                Name = "column1",
                ToolTipText = " Word size",
                Width = 100,
                // AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                DisplayMember = "S7WL",
                ValueMember = "Namevalue",
            };
            dataGridView1.Columns.Insert(1, column1);

            DataGridViewColumn column2 = new DataGridViewTextBoxColumn()
            {
                HeaderText = "数据块号",
                DataPropertyName = "DBNumber",
                Name = "column2",
                ToolTipText = "DB Number if Area = S7AreaDB",
                Width = 60,
                // AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,

            };
            dataGridView1.Columns.Insert(2, column2);

            DataGridViewColumn column3 = new DataGridViewTextBoxColumn()
            {
                HeaderText = "起始",
                DataPropertyName = "Start",
                Name = "column3",
                ToolTipText = " Offset to start",
                // AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                Width = 60,
            };

            dataGridView1.Columns.Insert(3, column3);
            DataGridViewColumn column4 = new DataGridViewTextBoxColumn()
            {
                HeaderText = "数量",
                DataPropertyName = "Amount",
                Name = "column4",
                ToolTipText = " Amount of elements to read ",
                Width = 60,
                // AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
            };
            dataGridView1.Columns.Insert(4, column4);

            DataGridViewColumn column5 = new DataGridViewTextBoxColumn()
            {
                HeaderText = "名称,;分隔",
                DataPropertyName = "names",
                Name = "column5",
                ToolTipText = " 所有数量的名称，之间以；分割 ",
                Width = 200,
                // AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
            };
            dataGridView1.Columns.Insert(5, column5);


        }


        private void button3_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = this.tabPage1;
            dataGridView1.Rows.Clear();
            xml_FilePath = "";
        }
        public int Createlistfromgv(ArrayList alonglist, ArrayList ashortlist)
        {
            int repeat = 0;
            string str2 = "";
            alonglist.Clear();
            ashortlist.Clear();
            int row = dataGridView1.Rows.Count;//得到总行数    

            for (int i = 0; i < row; i++)//遍历这个dataGridView
            {
                str2 = dataGridView1.Rows[i].Cells[5].Value.ToString();
                int wordlength = Convert.ToInt32(dataGridView1.Rows[i].Cells[1].Value);
                int length1 = Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value);
                if (length1 <= 0) continue;
                str2 = str2.Replace("；", ";");
                string[] sArray = Regex.Split(str2, ";", RegexOptions.IgnoreCase);
                if (sArray.Length < length1)
                {
                    MessageBox.Show("元素名称的数量不能小于实际采集的数量" + (i + 1).ToString());
                    return 0;
                }
                for (int j = 0; j < length1; j++)
                {
                    str2 = Convert.ToChar(Convert.ToByte('A') + i) + "_" + sArray[j];

                    foreach (recorditem rd in alonglist)
                    {
                        if (rd.Itemname == str2) repeat = 1;
                    }

                    if (repeat <= 0)
                    {
                        recorditem ri = new recorditem(str2, wordlength);
                        ri.REAL = 0.0;
                        ri.INTEGER = 0;
                        ri.CHAR = ' ';
                        ri.Obj = "nihao";
                        alonglist.Add(ri);
                        ashortlist.Add(sArray[j]);
                    }
                    else
                    {
                        MessageBox.Show("每行的元素名称不能重复" + (i + 1).ToString());
                        return 0;
                    }
                }
            }
            return 1;
        }



        public int createarrayfromgv()
        {
            // byte[] bytearray;
            //  ArrayList itemarraylist = new ArrayList();
            //   S7MultiVar Reader;
            Int32 L_Area;
            Int32 L_WordLen;
            Int32 L_DBNumber;
            Int32 L_Start;
            Int32 L_Amount;
            Int32 L_SizeRead;// 数组长度
                             //public bool Add<T>(Int32 Area, Int32 WordLen, Int32 DBNumber, Int32 Start, Int32 Amount, ref T[] Buffer)
            Reader = new S7MultiVar(Client);
            itemarraylist.Clear();
            Reader.Clear();
            int row = dataGridView1.Rows.Count;//得到总行数    
            string str2 = "";
            for (int i = 0; i < row; i++)//遍历这个dataGridView
            {
                str2 = dataGridView1.Rows[i].Cells[5].Value.ToString();
                int wordlength = Convert.ToInt32(dataGridView1.Rows[i].Cells[1].Value);
                int length1 = Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value);
                if (length1 <= 0) continue;
                str2 = str2.Replace("；", ";");
                string[] sArray = Regex.Split(str2, ";", RegexOptions.IgnoreCase);
                if (sArray.Length < length1)
                {
                    MessageBox.Show("元素名称的数量不能小于实际采集的数量" + (i + 1).ToString());
                    return 0;
                }
                L_Area = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value.ToString());
                L_WordLen = Convert.ToInt32(dataGridView1.Rows[i].Cells[1].Value.ToString());
                L_DBNumber = Convert.ToInt32(dataGridView1.Rows[i].Cells[2].Value.ToString());
                L_Start = Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value.ToString());
                L_Amount = Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value.ToString());
                L_SizeRead = Sharp7.S7.DataSizeByte(L_WordLen) * (L_Amount + 1) + 1;
                bytearray = new byte[L_SizeRead];
                itemarraylist.Add(bytearray);
                // bytearray = (byte[])itemarraylist[i];
                Reader.Add(L_Area, L_WordLen, L_DBNumber, L_Start, L_Amount, ref bytearray);


            }
            int Result = Reader.Read();
            //    ShowResult(Result);

            parsedata();
            //  HexDump(TxtDump, bytearray, 16);
            str2 = createinsertsqlstrformarrylist(longarray);
            //  TxtDump.Text = TxtDump.Text + "\n" + str2;

            SQLiteConnection cn = new SQLiteConnection("data source=" + dbPath);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                cmd.CommandText = (str2);
                cmd.ExecuteNonQuery();

            }
            cn.Close();




            return 1;

        }
        public void realtimechartseries(ArrayList list)
        {

            conn = new SQLiteConnection("Data Source = " + dbPath);
            conn.Open();
            //选择所有数据
            SQLiteDataAdapter mAdapter = new SQLiteDataAdapter("select * from RECORD order by ID desc limit 0,1000", conn);
            realtimedt = new System.Data.DataTable();
            mAdapter.Fill(realtimedt);
            //绑定数据到DataGridView
            chart2.DataSource = realtimedt;

            if (dbPath != "")
            {

            }
            chart2.DataBind();
            //关闭数据库
            conn.Close();
        }

        public void parsedata()
        {
            Int32 datatype1;
            Int32 datalength;
            Int32 indexnum = 0;
            Int32 wordsize = 0;
            string str2 = "";
            byte[] temparray;
            // Int32 index;
            //  indexnum = longarray.Count;

            for (int i = 0; i < itemarraylist.Count; i++)
            {
                recorditem dd1 = longarray[indexnum] as recorditem;
                datatype1 = dd1.Itemtype;
                // datatype1 = (Reader.Items[i].WordLen);//Sharp7.S7.DataSizeByte
                wordsize = Sharp7.S7.DataSizeByte(datatype1);
                datalength = (Reader.Items[i].Amount / wordsize);
                temparray = (byte[])itemarraylist[i];
                switch (datatype1)
                {
                    case S7Consts.S7WLBit:

                        for (int j = 0; j < datalength; j++)
                        {
                            recorditem dd = longarray[indexnum] as recorditem;

                            dd.INTEGER = Convert.ToInt32(S7.GetBitAt(temparray, j * wordsize, 0));
                            indexnum = indexnum + 1;
                        }
                        //  str2 = "INTEGER";
                        break;
                        ;
                    case S7Consts.S7WLByte:
                        for (int j = 0; j < datalength; j++)
                        {
                            recorditem dd = longarray[indexnum] as recorditem;

                            dd.INTEGER = Convert.ToInt32(S7.GetByteAt(temparray, j * wordsize));
                            indexnum = indexnum + 1;
                        }
                        //  str2 = "INTEGER";
                        break;

                    case S7Consts.S7WLChar:
                        for (int j = 0; j < datalength; j++)
                        {
                            recorditem dd = longarray[indexnum] as recorditem;
                            dd.CHAR = Convert.ToChar(S7.GetByteAt(temparray, j));
                            indexnum = indexnum + 1;
                        }

                        //  str2 = "TEXT";
                        break;

                    case S7Consts.S7WLWord:
                        for (int j = 0; j < datalength; j++)
                        {
                            recorditem dd = longarray[indexnum] as recorditem;

                            dd.INTEGER = Convert.ToInt32(S7.GetWordAt(temparray, j * wordsize));
                            indexnum = indexnum + 1;
                        }
                        //   str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLInt:
                        for (int j = 0; j < datalength; j++)
                        {
                            recorditem dd = longarray[indexnum] as recorditem;

                            dd.INTEGER = Convert.ToInt32(S7.GetIntAt(temparray, j * wordsize));
                            indexnum = indexnum + 1;
                        }
                        //  str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLDWord:
                        for (int j = 0; j < datalength; j++)
                        {
                            recorditem dd = longarray[indexnum] as recorditem;

                            dd.INTEGER = Convert.ToInt32(S7.GetDWordAt(temparray, j * wordsize));
                            indexnum = indexnum + 1;
                        }
                        //  str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLDInt:

                        for (int j = 0; j < datalength; j++)
                        {
                            recorditem dd = longarray[indexnum] as recorditem;

                            dd.INTEGER = Convert.ToInt32(S7.GetDIntAt(temparray, j * wordsize));
                            indexnum = indexnum + 1;
                        }
                        // str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLReal:
                        for (int j = 0; j < datalength; j++)
                        {
                            recorditem dd = longarray[indexnum] as recorditem;

                            dd.REAL = Convert.ToDouble(S7.GetRealAt(temparray, j * wordsize));
                            indexnum = indexnum + 1;
                        }
                        // str2 = "REAL";
                        break;
                    case S7Consts.S7WLCounter:
                        for (int j = 0; j < datalength; j++)
                        {
                            recorditem dd = longarray[indexnum] as recorditem;

                            dd.INTEGER = Convert.ToInt32(S7.GetLWordAt(temparray, j * wordsize));
                            indexnum = indexnum + 1;
                        }
                        //  str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLTimer:

                        for (int j = 0; j < datalength; j++)
                        {
                            recorditem dd = longarray[indexnum] as recorditem;

                            dd.INTEGER = Convert.ToInt32(S7.GetLTimeAt(temparray, j * wordsize));
                            indexnum = indexnum + 1;
                        }
                        //  str2 = "INTEGER";
                        break;
                }



            }
        }
        //从 longarray生成插入语句
        public string createinsertsqlstrformarrylist(ArrayList list)
        {


            string str2 = "";
            string str3 = "";
            string str12 = "";
            string str13 = "";
            string str = "";

            str = @" insert into    RECORD (";

            foreach (recorditem rd in list)
            {
                str2 = "";
                str3 = "";
                recorditem dd = rd as recorditem;
                int ItemType1 = dd.Itemtype;

                switch (ItemType1)
                {
                    case S7Consts.S7WLBit:

                        str2 = dd.Itemname;
                        str3 = dd.INTEGER.ToString();
                        //  str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLByte:
                        str2 = dd.Itemname;
                        str3 = dd.INTEGER.ToString();
                        //  str2 = "INTEGER";
                        break;

                    case S7Consts.S7WLChar:
                        str2 = dd.Itemname;
                        str3 = "\'" + dd.CHAR.ToString() + "\'";

                        // str2 = "TEXT";
                        break;

                    case S7Consts.S7WLWord:
                        str2 = dd.Itemname;
                        str3 = dd.INTEGER.ToString();
                        //  str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLInt:
                        str2 = dd.Itemname;
                        str3 = dd.INTEGER.ToString();
                        // str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLDWord:
                        str2 = dd.Itemname;
                        str3 = dd.INTEGER.ToString();
                        //   str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLDInt:
                        str2 = dd.Itemname;
                        str3 = dd.INTEGER.ToString();
                        // str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLReal:
                        str2 = dd.Itemname;
                        str3 = dd.REAL.ToString();
                        //  str2 = "REAL";
                        break;
                    case S7Consts.S7WLCounter:
                        str2 = dd.Itemname;
                        str3 = dd.INTEGER.ToString();
                        //   str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLTimer:
                        str2 = dd.Itemname;
                        str3 = dd.INTEGER.ToString();
                        //   str2 = "INTEGER";
                        break;
                }
                if (str12 == "")
                {
                    str12 = str12 + " " + str2;
                }
                else
                {
                    str12 = str12 + "," + str2;
                }
                if (str13 == "")
                {
                    str13 = str13 + " " + str3;
                }
                else
                {
                    str13 = str13 + "," + str3;
                }

            }

            str = str + str12 + @" )";
            str = str + "values (  ";
            str = str + str13 + @" )";
            str = str + "\n";
            return str;



            /*
                        foreach (recorditem rd in list)
                        {
                            str2 = "";
                            recorditem dd = rd as recorditem;
                            int ItemType1 = dd.Itemtype;

                            switch (ItemType1)
                            {
                                case S7Consts.S7WLBit:

                                    str2 = dd.Itemname;
                                    //  str2 = "INTEGER";
                                    break;
                                case S7Consts.S7WLByte:
                                    //  str2 = "INTEGER";
                                    break;

                                case S7Consts.S7WLChar:
                                    // str2 = "TEXT";
                                    break;

                                case S7Consts.S7WLWord:
                                    //  str2 = "INTEGER";
                                    break;
                                case S7Consts.S7WLInt:
                                    // str2 = "INTEGER";
                                    break;
                                case S7Consts.S7WLDWord:
                                    //   str2 = "INTEGER";
                                    break;
                                case S7Consts.S7WLDInt:
                                    // str2 = "INTEGER";
                                    break;
                                case S7Consts.S7WLReal:
                                    //  str2 = "REAL";
                                    break;
                                case S7Consts.S7WLCounter:
                                    //   str2 = "INTEGER";
                                    break;
                                case S7Consts.S7WLTimer:
                                    //   str2 = "INTEGER";
                                    break;
                            }
                            str2 = ",\n" + dd.Itemname + "   " + str2;
                            str = str + str2;
                        }
                        str = str + "\n)";
            */


        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="fileName">文件路径</param>
        /// <param name="myDGV">控件DataGridView</param>
        private void ExportExcels(DataGridView myDGV)
        {
            string str1 = " ";
            string saveFileName = "";
            str1 = System.Threading.Thread.GetDomain().BaseDirectory + "export";
            if (Directory.Exists(str1) == false)//如果不存
            {
                Directory.CreateDirectory(str1);
            }


            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xls";
            saveDialog.Filter = "Excel文件|*.xls";
            saveDialog.InitialDirectory = str1;
            //  saveDialog.FileName = fileName;
            saveDialog.ShowDialog();
            saveFileName = saveDialog.FileName;
            if (saveFileName.IndexOf(":") < 0) return; //被点了取消
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("无法创建Excel对象，可能您的机子未安装Excel");
                return;
            }
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1
                                                                                                                                  //写入标题
            for (int i = 0; i < myDGV.ColumnCount; i++)
            {
                worksheet.Cells[1, i + 1] = myDGV.Columns[i].HeaderText;
            }
            //写入数值
            for (int r = 0; r < myDGV.Rows.Count; r++)
            {
                for (int i = 0; i < myDGV.ColumnCount; i++)
                {
                    worksheet.Cells[r + 2, i + 1] = myDGV.Rows[r].Cells[i].Value;
                }
                System.Windows.Forms.Application.DoEvents();
            }
            worksheet.Columns.EntireColumn.AutoFit();//列宽自适应
            if (saveFileName != "")
            {
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(saveFileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                }
            }
            xlApp.Quit();
            GC.Collect();//强行销毁
            MessageBox.Show("文件： " + saveFileName + ".xls 保存成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }



        public void createrealtimechartseries(ArrayList list)
        {

            conn = new SQLiteConnection("Data Source = " + dbPath + ";Read Only = True;Cache Size=5000");
            conn.Open();
            //选择所有数据
            SQLiteDataAdapter mAdapter = new SQLiteDataAdapter("select * from RECORD order by ID desc limit 0,1000", conn);
            realtimedt = new System.Data.DataTable();
            mAdapter.Fill(realtimedt);
            //绑定数据到DataGridView
            chart2.DataSource = realtimedt;

            if (dbPath != "")
            {
                chart2.Series.Clear();
                checkedListBox2.Items.Clear();
                checkedListBox4.Items.Clear();
                chart2.ChartAreas[0].CursorX.IsUserEnabled = true;
                chart2.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
                chart2.ChartAreas[0].CursorX.LineColor = Color.Pink;
                chart2.ChartAreas[0].CursorX.IntervalType = DateTimeIntervalType.Auto;
                chart2.ChartAreas[0].CursorX.SelectionColor = System.Drawing.Color.Red;

                chart2.ChartAreas[0].CursorY.IsUserEnabled = true;
                chart2.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
                chart2.ChartAreas[0].CursorY.LineColor = Color.Pink;
                chart2.ChartAreas[0].CursorY.IntervalType = DateTimeIntervalType.Auto;

                chart2.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
                chart2.ChartAreas[0].AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.All;//启用X轴滚动条按钮

                // chart2.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Seconds;
                chart2.ChartAreas[0].AxisX.Interval = 10;   //设置X轴坐标的间隔为1

                chart2.ChartAreas[0].AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount;
                chart2.ChartAreas[0].AxisX.IsLabelAutoFit = true;
                chart2.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy-MM-dd-HH:mm:ss";
                // chart2.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss";         //毫秒格式： hh:mm:ss.fff ，后面几个f则保留几位毫秒小数，此时要注意轴的最大值和最小值不要差太大
                //  chart2.ChartAreas[0].AxisX.LabelStyle.IntervalType = DateTimeIntervalType.Seconds;
                chart2.ChartAreas[0].AxisX.LabelStyle.Angle = 45;
                //  chart2.ChartAreas[0].AxisX.LabelStyle.Interval = 10;  //坐标值间隔200 ms
                // chart2.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = false;   //防止X轴坐标跳跃

                // chart2.ChartAreas[0].AxisX.LabelStyle.IsStaggered = true;
                //   chart2.ChartAreas[0].AxisX.MajorGrid.IntervalType = DateTimeIntervalType.Seconds;
                //   chart2.ChartAreas[0].AxisX.MajorGrid.Interval = 1;


                chart2.ChartAreas[0].AxisY.ScaleView.Zoomable = true;
                chart2.ChartAreas[0].AxisY.ScrollBar.ButtonStyle = ScrollBarButtonStyles.All;//启用X轴滚动条按钮
                chart2.ChartAreas[0].AxisX.Enabled = AxisEnabled.True;
                chart2.ChartAreas[0].AxisY.Enabled = AxisEnabled.True;
                chart2.ChartAreas[0].AxisX2.Enabled = AxisEnabled.False;
                chart2.ChartAreas[0].AxisY2.Enabled = AxisEnabled.True;


                foreach (recorditem rd in list)
                {
                    recorditem dd = rd as recorditem;
                    chart2.Series.Add(dd.Itemname);//添加
                    chart2.Series[dd.Itemname].ChartType = SeriesChartType.Line;
                    chart2.Series[dd.Itemname].IsXValueIndexed = true;
                    chart2.Series[dd.Itemname].XValueMember = "PCTIME";
                    chart2.Series[dd.Itemname].XValueType = ChartValueType.DateTime;
                    chart2.Series[dd.Itemname].YValueMembers = dd.Itemname;
                    chart2.Series[dd.Itemname].LegendText = dd.Itemname;
                    chart2.Series[dd.Itemname].BorderColor = Color.FromArgb(180, 26, 59, 105);
                    chart2.Series[dd.Itemname].IsValueShownAsLabel = false;//图上不显示数据点的值
                    chart2.Series[dd.Itemname].ToolTip = "#SER\n #VALY";//鼠标停留在数据点上，显示XY值
                    checkedListBox2.Items.Add(dd.Itemname, true);
                    checkedListBox4.Items.Add(dd.Itemname, true);
                }





            }
            chart2.DataBind();
            //关闭数据库
            conn.Close();
        }

        public void createchartseries(ArrayList list)
        {

            //1、在屏幕的右下角显示窗体
            //这个区域不包括任务栏的
            //  Rectangle ScreenArea = System.Windows.Forms.Screen.GetWorkingArea(this);
            //这个区域包括任务栏，就是屏幕显示的物理范围
            System.Drawing.Rectangle ScreenArea = System.Windows.Forms.Screen.GetBounds(this);
            picwidth = ScreenArea.Width;  //屏幕宽度 
            picheight = ScreenArea.Height; //屏幕高度

            picwidth = PrimaryScreen.DESKTOP.Width;
            picheight = PrimaryScreen.DESKTOP.Height; //屏幕高度

            int counter;
            conn = new SQLiteConnection("Data Source = " + dbPath + ";Read Only = True;Cache Size=5000");
            conn.Open();

            string sqlstr2;

            /*       chart控件绑定数据超过2000，明显会变卡顿
所以建议如果数据过多，根据条件对数据进行分段或分组显示
对于UI来说，加载数据过多也没有意义
2.读取sql查询语句的内容使用SqlDataReader()方法
而SqlCommand.ExecuteScalar()方法的作用就是

执行查询，并返回查询所返回的结果集中第一行的第一列。忽略其他行或列，返回值为object类型
*/
            //       select count(*) from table   SELECT COUNT(*) FROM table_name
            //"select * from RECORD order by ID desc limit 0,1000"
            SQLiteCommand cmd = conn.CreateCommand();
            string strsql1 = " SELECT COUNT(*) FROM RECORD";
            cmd.CommandText = strsql1;
            counter = Convert.ToInt32(cmd.ExecuteScalar());
            if (counter == 0) return;
            label22.Text = "总行数：" + counter.ToString();
            splitcounter = Convert.ToInt32(textBox2.Text);
            totalnumber = counter / splitcounter + 1;
            //确定图像尺寸的宽度 

            if (totalnumber == 1) picwidth = counter;
            else picwidth = splitcounter;


            if (currentnumber >= totalnumber) currentnumber = 0;
            if (currentnumber < 0) currentnumber = totalnumber - 1;

            //选择所有数据 select * from table order by ID desc limit 0,20
            if (checkBox5.Checked)
            {
                sqlstr2 = "select * from RECORD order by ID ASC limit " + (currentnumber * splitcounter).ToString() + "," + splitcounter.ToString();
            }
            else
            {//select * from dd1    where id-(id/3)*3=0
                sqlstr2 = "select * from RECORD where " + "ID % "+ totalnumber.ToString() + "= " + currentnumber.ToString() + " order by ID ASC ";

            }

            SQLiteDataAdapter mAdapter = new SQLiteDataAdapter(sqlstr2, conn);
            System.Data.DataTable dt = new System.Data.DataTable();
            mAdapter.Fill(dt);
            //绑定数据到DataGridView
            chart1.DataSource = dt;


            if (dbPath != "")
            {
                chart1.Series.Clear();
                checkedListBox1.Items.Clear();
                checkedListBox3.Items.Clear();
                chart1.ChartAreas[0].CursorX.IsUserEnabled = true;
                chart1.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
                chart1.ChartAreas[0].CursorX.LineColor = Color.Pink;
                chart1.ChartAreas[0].CursorX.IntervalType = DateTimeIntervalType.Auto;
                chart1.ChartAreas[0].CursorY.IsUserEnabled = true;
                chart1.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
                chart1.ChartAreas[0].CursorY.LineColor = Color.Pink;
                chart1.ChartAreas[0].CursorY.IntervalType = DateTimeIntervalType.Auto;
                chart1.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
                // chart1.ChartAreas[0].AxisX.ScaleView.Zoom
                chart1.ChartAreas[0].AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.All;//启用X轴滚动条按钮
                chart1.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy-MM-dd-HH:mm:ss";
                // chart1.ChartAreas[0].AxisX.LabelStyle.IsStaggered = true;   //设置是否交错显示,比如数据多的时间分成两行来显
                chart1.ChartAreas[0].AxisX.LabelStyle.Angle = 45;
                chart1.ChartAreas[0].AxisY.ScaleView.Zoomable = true;
                //  chart1.ChartAreas[0].AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount;
                //  chart1.ChartAreas[0].AxisX.IsLabelAutoFit = false;
                chart1.ChartAreas[0].AxisY.ScrollBar.ButtonStyle = ScrollBarButtonStyles.All;//启用X轴滚动条按钮
                chart1.ChartAreas[0].AxisX.Enabled = AxisEnabled.True;
                chart1.ChartAreas[0].AxisY.Enabled = AxisEnabled.True;
                chart1.ChartAreas[0].AxisX2.Enabled = AxisEnabled.False;
                chart1.ChartAreas[0].AxisY2.Enabled = AxisEnabled.True;


                /*
                 * 
                 *     //X轴设置
            chart1.ChartAreas[0].AxisX.Title = "时间";
            chart1.ChartAreas[0].AxisX.TitleAlignment = StringAlignment.Near;
            chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;//不显示竖着的分割线
 
            /************************************************************************/
                /* 本文重点讲解时间格式的设置
                 * 如果想显示原点第一个时间坐标,需要设置最小时间,时间间隔类型，时间间隔值等三个参数
             
                chart1.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss"; //X轴显示的时间格式，HH为大写时是24小时制，hh小写时是12小时制
                chart1.ChartAreas[0].AxisX.Minimum = DateTime.Parse("09:10:02").ToOADate();
                chart1.ChartAreas[0].AxisX.Maximum = DateTime.Parse("09:10:21").ToOADate();
                chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Seconds;//如果是时间类型的数据，间隔方式可以是秒、分、时
                chart1.ChartAreas[0].AxisX.Interval = DateTime.Parse("00:00:02").Second;//间隔为2秒
                */

                foreach (recorditem rd in list)
                {
                    recorditem dd = rd as recorditem;
                    chart1.Series.Add(dd.Itemname);//添加
                    chart1.Series[dd.Itemname].ChartType = SeriesChartType.Line;
                    chart1.Series[dd.Itemname].IsXValueIndexed = true;
                    chart1.Series[dd.Itemname].XValueMember = "PCTIME";
                    chart1.Series[dd.Itemname].XValueType = ChartValueType.DateTime;
                    chart1.Series[dd.Itemname].YValueMembers = dd.Itemname;
                    chart1.Series[dd.Itemname].LegendText = dd.Itemname;
                    chart1.Series[dd.Itemname].BorderColor = Color.FromArgb(180, 26, 59, 105);
                    chart1.Series[dd.Itemname].IsValueShownAsLabel = false;//图上不显示数据点的值
                    chart1.Series[dd.Itemname].ToolTip = "#SER\n #VALY";//鼠标停留在数据点上，显示XY值
                    checkedListBox1.Items.Add(dd.Itemname, true);
                    checkedListBox3.Items.Add(dd.Itemname, true);
                }



            }


            chart1.DataBind();
            counter = chart1.Series[0].Points.Count;
            chart1.PerformLayout();
            chart1.ApplyPaletteColors();//将调色板颜色属性设置到serial clolr属性上

            if (counter > 1)
                ft1 = chart1.Series[0].Points[chart1.Series[0].Points.Count - 1].XValue;
            chart1.ChartAreas[0].AxisX.ScaleView.Size = 500;// 可视区域数据点数
            chart1.ChartAreas[0].AxisX.ScaleView.SizeType = DateTimeIntervalType.Number;
            //   chart1.ChartAreas[0].AxisX.ScaleView.Zoom( DateTime.FromOADate(ft1).AddHours(-0.5).ToOADate() , ft1);
            //  chart1.ChartAreas[0].AxisX.ScaleView.Scroll(ft1);

            ;
            //关闭数据库
            conn.Close();
            textBox3.Text = picwidth.ToString();
            textBox12.Text = picheight.ToString();

        }


        // 保存当前显示的图表到图片文件。
        public void savechartimage()
        {


            double y1totallength, y2totallength, x1totallength;
            int count1 = chart1.Series[0].Points.Count;
            //图纸的大小 及图像的定义
            picwidth = Convert.ToInt32(textBox3.Text);
            picheight = Convert.ToInt32(textBox12.Text);


            //坐标位置的偏移量 注意绘图坐标系（原点在左下角）和图像坐标系（原点在左上角）的不同。
            biasx = 100;
            biasy = 100;
            Bitmap BM1 = new Bitmap(picwidth+ biasx*2+10, picheight);
            BM1.MakeTransparent(Color.White);
            Graphics flagGraphics = Graphics.FromImage(BM1);
            SolidBrush backgroud = new SolidBrush(Color.White);//这里修改背景颜色
            flagGraphics.FillRectangle(backgroud, 0, 0, picwidth + biasx * 2 + 10, picheight);
            //将颜色调色板应用到曲线系列以便获取曲线颜色chart1.Series.Color,
            chart1.ApplyPaletteColors();

            SolidBrush drawBrush = new SolidBrush(Color.Black);// Create point for upper-left corner of drawing.







            //坐标轴y1，y2轴 x1轴上下限及其范围大小
            int y1min = Convert.ToInt32(textBox5.Text);
            int y1max = Convert.ToInt32(textBox4.Text);
            y1totallength = y1max - y1min;
            int y2min = Convert.ToInt32(textBox7.Text);
            int y2max = Convert.ToInt32(textBox6.Text);
            y2totallength = y2max - y2min;
            DateTime dt1 = DateTime.FromOADate(chart1.Series[0].Points[0].XValue);
            DateTime dt2 = DateTime.FromOADate(chart1.Series[0].Points[count1 - 1].XValue);
            x1totallength = (dt2 - dt1).TotalSeconds;
            int x1min = (int)chart1.Series[0].Points[0].XValue;
            int x1max = (int)chart1.Series[0].Points[count1 - 1].XValue;


            #region   //坐标轴及标签的绘制
            Pen penxy = new Pen(Color.Gray, 1);

            string textxy = "";
            System.Drawing.Font fontxy = new System.Drawing.Font("Arial", 12, FontStyle.Bold, GraphicsUnit.Point);
            //y方向轴标签
            penxy.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash; //虚线






            for (int i = 0; i <= 1000; i++)
            {

                if (i % 100 == 0)
                {
                    float x1 = biasx;
                    float x2 = picwidth + biasx;
                    float y1 = (float)(picheight - (biasy + i * (picheight - biasy * 2) / 1000.0));
                    float y2 = (float)(picheight - (biasy + i * (picheight - biasy * 2) / 1000.0));
                    flagGraphics.DrawLine(penxy, x1, y1, x2, y2);
                    textxy = ((float)y1min + y1totallength / 1000.0 * i).ToString("#0.00");

                    flagGraphics.DrawString(textxy, fontxy, Brushes.Blue, 30, y1);
                    textxy = ((float)y2min + y2totallength / 1000.0 * i).ToString("#0.00");
                  
                    flagGraphics.DrawString(textxy, fontxy, Brushes.Blue, x2-50, y1);


                    continue;
                }

                if (i % 10 == 0)
                {
                    float x1 = biasx;
                    float x2 = picwidth + biasx;
                    float y1 = (float)(picheight - (biasy + i * (picheight - biasy * 2) / 1000.0));
                    float y2 = (float)(picheight - (biasy + i * (picheight - biasy * 2) / 1000.0));
                    flagGraphics.DrawLine(penxy, x1, y1, x1 + 20, y2);
                    flagGraphics.DrawLine(penxy, x2, y1, x2 - 20, y2);
                    continue;
                }


            }

            //x方向轴标签


            //////////////////
            GraphicsText graphicsText = new GraphicsText();
            graphicsText.Graphics = flagGraphics;

            StringFormat stringFormatxy = new StringFormat();
            stringFormatxy.Alignment = StringAlignment.Center;
            stringFormatxy.LineAlignment = StringAlignment.Center;


            ////////////////
            penxy.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid; //实线
            for (int i = 0; i < count1; i++)
            {
                if (i % 100 == 0)
                {
                    float x1 = (biasx + i);
                    float x2 = (biasx + i);
                    float y1 = (float)(picheight - biasy);
                    float y2 = (float)(biasy);
                    flagGraphics.DrawLine(penxy, x1, y1, x2, y2);
                  //  chart1.Series[0].Points[0].XValue+i*  x1totallength / count1;
                   // DateTime dty1 = DateTime.FromOADate(chart1.Series[0].Points[0].XValue).AddSeconds(i * x1totallength / count1);
                    DateTime dty1 = DateTime.FromOADate(chart1.Series[0].Points[i].XValue);

                    if (i == 0)
                    {
                        textxy = dty1.ToString("yyyy-MM-dd");

                    }
                    else
                    {
                        textxy = dty1.ToString("HH:mm:ss:fff");

                    }
                    graphicsText.DrawString(textxy, fontxy, Brushes.Blue, new PointF(x1 + 20, y1 + 35), stringFormatxy, 30f);

                    //   flagGraphics.DrawString(textxy, fontxy, Brushes.Blue, x1, y1 + 30);

                    continue;
                }

                if (i % 10 == 0)
                {
                    float x1 = (biasx + i);
                    float x2 = (biasx + i);
                    float y1 = (float)(picheight - biasy);
                    float y2 = (float)(biasy);
                    flagGraphics.DrawLine(penxy, x1, y1, x2, y1 - 20);
                    continue;
                }



            }

            #endregion


            #region  图表及标签打印

            for (int i = 0; i < chart1.Series.Count; i++)
            {
              if(   checkedListBox1.GetItemChecked(i)==false) continue;
                
                int serielength = chart1.Series[i].Points.Count;
                //将颜色调色板应用到曲线系列以便获取曲线颜色chart1.Series.Color,
                chart1.ApplyPaletteColors();
                penxy.Color = chart1.Series[i].Color;


                drawBrush.Color = chart1.Series[i].Color;

                if (checkedListBox3.GetItemChecked(i) == true)
                    textxy ="y1:"+i.ToString()+":" +chart1.Series[i].Name;
                else
                    textxy = "y2:" + i.ToString() + ":"+chart1.Series[i].Name;
               
                flagGraphics.DrawString(textxy, fontxy, drawBrush, biasx+40, biasy+10+i*30);


                //  DateTime.FromOADate(chart1.Series[0].Points[0].XValue).AddSeconds(i * x1totallength / count1);
                DateTime dtcurrent = DateTime.FromOADate(chart1.Series[i].Points[0].XValue);
                float x1 =   biasx;
                float y1;
                if (checkedListBox3.GetItemChecked(i) == true)
                    y1 =(float)( picheight - ((chart1.Series[i].Points[0].YValues[0] - y1min) / y1totallength * (picheight - biasx * 2) + biasx));
                else
                    y1 = (float)(picheight - ((chart1.Series[i].Points[0].YValues[0] - y2min) / y2totallength * (picheight - biasx * 2) + biasx));


                float x2, y2;



                for (int j = 1; j < serielength; j++)
                {
                  

                    dtcurrent = DateTime.FromOADate(chart1.Series[i].Points[j].XValue);
                    // x2 = (float)((dtcurrent - dt1).TotalSeconds / x1totallength * count1 + biasx);
                    x2 = x1 + 1;
                    if (checkedListBox3.GetItemChecked(i) == true)
                        y2 = (float)(picheight - ((chart1.Series[i].Points[j].YValues[0] - y1min) / y1totallength * (picheight - biasx * 2) + biasx));
                    else
                        y2 = (float)(picheight - ((chart1.Series[i].Points[j].YValues[0] - y2min) / y2totallength * (picheight - biasx * 2) + biasx)); ;



                


                    flagGraphics.DrawLine(penxy, x1, y1, x2, y2);
                    x1 = x2;
                    y1 = y2;


                    if (j== serielength-1)
                    {
                        float temp;

                        if (checkedListBox3.GetItemChecked(i) == true)
                            textxy = "y1:" + i.ToString();
                        else
                            textxy = "y2:" + i.ToString();

                        temp = x2+5 ;
                        if (temp > picwidth-30+biasx)  temp = picwidth - 30 +biasx;
            flagGraphics.DrawString(textxy, fontxy, drawBrush, temp, y2);

                    }


                }



            }



            #endregion



            #region 保存Chart中选择的曲线图片
            //保存Chart中选择的曲线图片

            string strdir1 = " ";
            strdir1 = System.Threading.Thread.GetDomain().BaseDirectory + "pic";
            if (Directory.Exists(strdir1) == false)//如果不存
            {
                Directory.CreateDirectory(strdir1);
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.InitialDirectory = strdir1;
            sfd.Filter = "BMP文件|*.bmp|JPEG文件|*.jpg|PNG文件|*.png";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                if (sfd.FilterIndex == 1)
                    BM1.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Bmp);
                if (sfd.FilterIndex == 2)
                    BM1.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                if (sfd.FilterIndex == 3)
                    BM1.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Png);
            }

            #endregion

        }
        //根据arrilist 生成实时插入的语句
        public string createsqlstr(ArrayList list)
        {
            string str = "";
            str = str + @" CREATE TABLE RECORD
        (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                PCTIME   TimeStamp NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%f', 'now', 'localtime')) ";
            string str2 = "";

            foreach (recorditem rd in list)
            {
                str2 = "";
                recorditem dd = rd as recorditem;
                int ItemType1 = dd.Itemtype;

                switch (ItemType1)
                {
                    case S7Consts.S7WLBit:
                        str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLByte:
                        str2 = "INTEGER";
                        break;

                    case S7Consts.S7WLChar:
                        str2 = "TEXT";
                        break;

                    case S7Consts.S7WLWord:
                        str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLInt:
                        str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLDWord:
                        str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLDInt:
                        str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLReal:
                        str2 = "REAL";
                        break;
                    case S7Consts.S7WLCounter:
                        str2 = "INTEGER";
                        break;
                    case S7Consts.S7WLTimer:
                        str2 = "INTEGER";
                        break;
                }
                str2 = ",\n" + dd.Itemname + "   " + str2;
                str = str + str2;
            }

            str = str + "\n)";


            return str;


        }
        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

            int returnint = 0;
            returnint = Createlistfromgv(longarray, shortarray);
            //    listFlag = createarrayfromgv() == 1 ? true : false;
            if (returnint <= 0) return;



            SaveFileDialog savedbDialog1 = new System.Windows.Forms.SaveFileDialog();//打开一个保存对话框
            string str1 = " ";
            str1 = System.Threading.Thread.GetDomain().BaseDirectory + "database";
            if (Directory.Exists(str1) == false)//如果不存
            {
                Directory.CreateDirectory(str1);
            }
            savedbDialog1.InitialDirectory = str1;
            savedbDialog1.Filter = "sqlite文件(*.sqlite)|*.sqlite";//设置允许打开的扩展名
            savedbDialog1.OverwritePrompt = false;
            if (savedbDialog1.ShowDialog() == DialogResult.OK)//判断是否选择了一个文件路径
            {
                if (System.IO.File.Exists(savedbDialog1.FileName))
                {
                    MessageBox.Show("文件已存在,数据库文件不允许覆盖，\n 可先手动删除再创建");

                    return;
                }
                else
                {
                    SQLiteConnection.CreateFile(savedbDialog1.FileName);


                    string path = savedbDialog1.FileName;
                    SQLiteConnection cn = new SQLiteConnection("data source=" + path);
                    if (cn.State != System.Data.ConnectionState.Open)
                    {
                        cn.Open();
                        SQLiteCommand cmd = new SQLiteCommand();
                        cmd.Connection = cn;

                        // 
                        //datetime('now', 'localtime')
                        cmd.CommandText = createsqlstr(longarray);
                        cmd.ExecuteNonQuery();
                        str1 = @" CREATE TABLE SETTING
                               (
                                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                PCTIME   TimeStamp NOT NULL DEFAULT(strftime('%Y-%m-%d %H:%M:%f', 'now', 'localtime')),
                                  XMLDATA          TEXT
                                )";
                        cmd.CommandText = (str1);
                        cmd.ExecuteNonQuery();


                        XmlDocument xmlDocument = new XmlDocument();//新建一个XML“编辑器”
                        XmlElement xmlElement_class = xmlDocument.CreateElement("单任务");//创建一个<class>节点
                        int row = dataGridView1.Rows.Count;//得到总行数    
                                                           //int cell = dataGridView1.Rows[1].Cells.Count;//得到总列数    
                        for (int i = 0; i <= row - 1; i++)//得到总行数并在之内循环    
                        {
                            //同上，创建一个个<student>节点，并且附到<class>之下
                            XmlElement xmlElement_student = xmlDocument.CreateElement("变量列表");//创建一个<student>节点
                            XmlElement xmlElement_name = xmlDocument.CreateElement("区域");//创建<name>节点
                            xmlElement_name.InnerText = dataGridView1.Rows[i].Cells[0].Value.ToString();//其文本就是第0个单元格的内容
                            xmlElement_student.AppendChild(xmlElement_name);//在<student>下面添加一个新的节点<name>
                                                                            //同理添加<number>
                            XmlElement xmlElement_number = xmlDocument.CreateElement("长度");
                            xmlElement_number.InnerText = dataGridView1.Rows[i].Cells[1].Value.ToString();
                            xmlElement_student.AppendChild(xmlElement_number);
                            XmlElement xmlElement_db = xmlDocument.CreateElement("数据块号");
                            xmlElement_db.InnerText = dataGridView1.Rows[i].Cells[2].Value.ToString();
                            xmlElement_student.AppendChild(xmlElement_db);

                            XmlElement xmlElement_start = xmlDocument.CreateElement("起始");
                            xmlElement_start.InnerText = dataGridView1.Rows[i].Cells[3].Value.ToString();
                            xmlElement_student.AppendChild(xmlElement_start);

                            XmlElement xmlElement_num = xmlDocument.CreateElement("数量");
                            xmlElement_num.InnerText = dataGridView1.Rows[i].Cells[4].Value.ToString();
                            xmlElement_student.AppendChild(xmlElement_num);

                            XmlElement xmlElement_name1 = xmlDocument.CreateElement("名称");
                            xmlElement_name1.InnerText = dataGridView1.Rows[i].Cells[5].Value.ToString();
                            xmlElement_student.AppendChild(xmlElement_name1);
                            xmlElement_class.AppendChild(xmlElement_student);
                        }
                        xmlDocument.AppendChild(xmlDocument.CreateXmlDeclaration("1.0", "utf-8", ""));//编写文件头
                        xmlDocument.AppendChild(xmlElement_class);//将这个<class>附到总文件头，而且设置为根结点
                        str1 = xmlToString(xmlDocument);
                        str1 = Convert.ToBase64String(System.Text.Encoding.Default.GetBytes(str1));
                        str1 = @"INSERT INTO SETTING(XMLDATA) VALUES(" + "\"" + str1;
                        str1 = str1 + "\"" + ")";
                        cmd.CommandText = str1;
                        cmd.ExecuteNonQuery();


                        toolStripStatusLabel3.Text = savedbDialog1.FileName;
                        dbPath = savedbDialog1.FileName;
                    }
                    cn.Close();



                    createchartseries(longarray);

                    createrealtimechartseries(longarray);
                }

            }
        }

        //转换为字符串
        public string xmlToString(XmlDocument xmlDoc)
        {
            MemoryStream stream = new MemoryStream();
            XmlTextWriter writer = new XmlTextWriter(stream, null);
            writer.Formatting = Formatting.Indented;
            xmlDoc.Save(writer);
            StreamReader sr = new StreamReader(stream, System.Text.Encoding.UTF8);
            stream.Position = 0;
            string xmlString = sr.ReadToEnd();
            sr.Close();
            stream.Close();
            return xmlString;
        }
        private void button10_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = this.tabPage1;
            string str1 = (dataGridView1.RowCount + 1).ToString() + "_" + "1";
            for (int i = 2; i <= 10; i++)
            {
                str1 = str1 + ";" + (dataGridView1.RowCount + 1).ToString() + "_" + i.ToString();
            }
            // this.dataGridView1.Rows[index].Cells[5].Value = str1;

            dataGridView1.Rows.Add(S7Consts.S7AreaDB, S7Consts.S7WLInt, 0, 0, 0, str1);

        }

        private void button11_Click(object sender, EventArgs e)
        {
            int returnvalue;
            returnvalue = Createlistfromgv(longarray, shortarray);
            if (returnvalue == 0)
            {

                return;
            };
            string str = "";
            foreach (recorditem ss in longarray)
            {
                recorditem dd = ss as recorditem;
                str = str + dd.Itemname + "\n";
            }
            str = str + longarray.Count + "\n";
            richTextBox1.Clear();
            richTextBox1.AppendText(str);
            richTextBox1.AppendText("\n");
            richTextBox1.AppendText(createsqlstr(longarray));
        }

        private void button5_Click(object sender, EventArgs e)
        {


            tabControl1.SelectedTab = this.tabPage3;

        }

        public static String sqliteEscape(String keyWord)
        {
            keyWord = keyWord.Replace("/", "//");
            keyWord = keyWord.Replace("'", "''");
            keyWord = keyWord.Replace("[", "/[");
            keyWord = keyWord.Replace("]", "/]");
            keyWord = keyWord.Replace("%", "/%");
            keyWord = keyWord.Replace("&", "/&");
            keyWord = keyWord.Replace("_", "/_");
            keyWord = keyWord.Replace("(", "/(");
            keyWord = keyWord.Replace(")", "/)");
            return keyWord;
        }


        public class recorditem
        {
            public string Itemname { get; set; }//类型列名
            public Int32 Itemtype { get; set; }
            public object Obj { get; set; }
            public Int32 INTEGER;
            public double REAL;
            public char CHAR;

            public recorditem()
            { }

            public recorditem(string itemname, Int32 itemtype)
            {
                Itemname = itemname;
                Itemtype = itemtype;

            }

        }

        private void button12_Click_1(object sender, EventArgs e)
        {
            
            string path = textBox1.Text;
            if (path =="")
           {
                MessageBox.Show("选择目录");
                return;
            }
            DirectoryInfo currentDir = new DirectoryInfo(path);
            // DirectoryInfo[] dirs = currentDir.GetDirectories(); //获取目录
            FileInfo[] files = currentDir.GetFiles();   //获取文件
            listView1.View = View.Details;
            listView1.BeginUpdate();
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Columns.Add("短名称", 300, HorizontalAlignment.Left);
            listView1.Columns.Add("长名称", 300, HorizontalAlignment.Left);
            listView1.Columns.Add("大小", 120, HorizontalAlignment.Left);
            listView1.Columns.Add("修改时间", 200, HorizontalAlignment.Left);


            foreach (FileInfo file in files)
            {
                ListViewItem fileItem = listView1.Items.Add(file.Name);
                if (file.Extension == ".sqlite")   //程序文件或无扩展名
                {
                    fileItem.Name = file.FullName;
                    fileItem.SubItems.Add(file.FullName);
                    fileItem.SubItems.Add(file.Length / 1000 + "KB");
                    fileItem.SubItems.Add(file.LastWriteTimeUtc.ToString());
                }

            }

            listView1.EndUpdate();


        }

        private void button13_Click(object sender, EventArgs e)
        {
            textBox1.Text = SelectPath();

        }

        private string SelectPath()
        {
            string path = string.Empty;
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            fbd.SelectedPath = System.Threading.Thread.GetDomain().BaseDirectory + "database";
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = fbd.SelectedPath;
            }
            return path;
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str1 = "";
            if (listView1.SelectedItems.Count == 0) return;
            else
            {
                string site = listView1.SelectedItems[0].Name;
                toolStripStatusLabel3.Text = site;

                conn = null;

                dbPath = site;
                //创建数据库实例，指定文件位置
                // Read Only = True;Cache Size=2000;
                conn = new SQLiteConnection("Data Source = " + dbPath + ";Read Only = True;Cache Size=5000");
                conn.Open();

                comboBox1.Items.Clear();
                using (System.Data.DataTable mTables = conn.GetSchema("Tables")) // "Tables"包含系统表详细信息；
                {
                    for (int i = 0; i < mTables.Rows.Count; i++)
                    {
                        comboBox1.Items.Add(mTables.Rows[i].ItemArray[mTables.Columns.IndexOf("TABLE_NAME")].ToString());
                    }
                    if (comboBox1.Items.Count > 0)
                    {
                        comboBox1.SelectedIndex = 0; // 默认选中第一张表.
                    }
                }


                SQLiteDataAdapter mAdapter = new SQLiteDataAdapter("select   * from SETTING Limit 1", conn);
                System.Data.DataTable dt = new System.Data.DataTable();
                mAdapter.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    str1 = dt.Rows[0]["XMLDATA"].ToString();
                    str1 = System.Text.Encoding.Default.GetString(Convert.FromBase64String(str1));
                    richTextBox1.AppendText(str1);
                }


                XmlDocument xmlDocument = new XmlDocument();//新建一个XML“编辑器”
                xmlDocument.LoadXml(str1);//载入路径这个xml
                try
                {

                    XmlNodeList xmlNodeList = xmlDocument.SelectSingleNode("单任务").ChildNodes;//选择class为根结点并得到旗下所有子节点
                    dataGridView1.Rows.Clear();//清空dataGridView1，防止和上次处理的数据混乱
                    foreach (XmlNode xmlNode in xmlNodeList)//遍历class的所有节点
                    {
                        XmlElement xmlElement = (XmlElement)xmlNode;//对于任何一个元素，其实就是每一个<student>
                        //旗下的子节点<name>和<number>分别放入dataGridView1
                        int index = dataGridView1.Rows.Add();//在dataGridView1新加一行，并拿到改行的行标
                        dataGridView1.Rows[index].Cells[0].Value = Convert.ToByte(xmlElement.ChildNodes.Item(0).InnerText);//各个单元格分别添加
                        dataGridView1.Rows[index].Cells[1].Value = Convert.ToInt32(xmlElement.ChildNodes.Item(1).InnerText);
                        dataGridView1.Rows[index].Cells[2].Value = xmlElement.ChildNodes.Item(2).InnerText;//各个单元格分别添加
                        dataGridView1.Rows[index].Cells[3].Value = xmlElement.ChildNodes.Item(3).InnerText;
                        dataGridView1.Rows[index].Cells[4].Value = xmlElement.ChildNodes.Item(4).InnerText;//各个单元格分别添加
                        dataGridView1.Rows[index].Cells[5].Value = xmlElement.ChildNodes.Item(5).InnerText;
                    }
                }
                catch
                {
                    MessageBox.Show("XML格式不对！");
                }
                conn.Close();
                toolStripStatusLabel3.Text = dbPath;
                int returnint = 0;
                returnint = Createlistfromgv(longarray, shortarray);
                //listFlag = createarrayfromgv() == 1 ? true : false;
                if (returnint <= 0) return;
                createchartseries(longarray);
                createrealtimechartseries(longarray);
            }
        }

        private void refreshgrid()
        {

            conn = null;
            //创建数据库实例，指定文件位置
            conn = new SQLiteConnection("Data Source = " + dbPath);
            conn.Open();
            //选择所有数据
            SQLiteDataAdapter mAdapter = new SQLiteDataAdapter("select * from " + comboBox1.Text, conn);
            System.Data.DataTable dt = new System.Data.DataTable();
            mAdapter.Fill(dt);
            //绑定数据到DataGridView
            dataGridView2.DataSource = dt;
            dataGridView2.Columns[1].DefaultCellStyle.Format = "yyyy/MM/dd HH:mm:ss:fff";
            //关闭数据库
            conn.Close();
            this.dataGridView2.FirstDisplayedScrollingRowIndex = this.dataGridView2.Rows.Count - 1;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            refreshgrid();
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.DataVisualization.Charting.Legend lg in chart1.Legends)
            {
                if (lg.Enabled)
                { lg.Enabled = false; }
                else
                { lg.Enabled = true; }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        #region contextMenu
        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void copy_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }

        private void selectall_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectAll();
        }

        private void clearall_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void paste_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }



        #endregion


        private void label2_Click(object sender, EventArgs e)
        {

        }



        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            Boolean checked1 = false;
            CheckedListBox clb = (CheckedListBox)sender;
            checked1 = clb.GetItemChecked(e.Index);
            if (checked1)
            {

                chart1.Series[clb.Items[e.Index].ToString()].Enabled = false;
            }
            else
            {
                chart1.Series[clb.Items[e.Index].ToString()].Enabled = true;

            }

        }



        //局部放大后，恢复视图
        private void chart1_AxisScrollBarClicked(object sender, ScrollBarEventArgs e)
        {
            if (e.ButtonType == ScrollBarButtonType.ZoomReset)
            {
                chart1.ChartAreas[0].AxisX.Interval = 0;
            }

        }

        private void button4_Click_1(object sender, EventArgs e)
        {

            string str1 = " ";
            str1 = System.Threading.Thread.GetDomain().BaseDirectory + "pic";
            if (Directory.Exists(str1) == false)//如果不存
            {
                Directory.CreateDirectory(str1);
            }


            //保存Chart1的图片
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.InitialDirectory = str1;
            sfd.Filter = "BMP文件|*.bmp|JPEG文件|*.jpg|PNG文件|*.png";
            if (sfd.ShowDialog() == DialogResult.OK)
            {

                if (sfd.FilterIndex == 1)
                    chart1.SaveImage(sfd.FileName, System.Drawing.Imaging.ImageFormat.Bmp);
                if (sfd.FilterIndex == 2)
                    chart1.SaveImage(sfd.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                if (sfd.FilterIndex == 3)
                    chart1.SaveImage(sfd.FileName, System.Drawing.Imaging.ImageFormat.Png);

            }
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            createchartseries(longarray);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            createrealtimechartseries(longarray);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.DataVisualization.Charting.Legend lg in chart2.Legends)
            {
                if (lg.Enabled)
                { lg.Enabled = false; }
                else
                { lg.Enabled = true; }
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            string str1 = " ";
            str1 = System.Threading.Thread.GetDomain().BaseDirectory + "pic";
            if (Directory.Exists(str1) == false)//如果不存
            {
                Directory.CreateDirectory(str1);
            }


            //保存Chart1的图片
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.InitialDirectory = str1;
            sfd.Filter = "BMP文件|*.bmp|JPEG文件|*.jpg|PNG文件|*.png";
            if (sfd.ShowDialog() == DialogResult.OK)
            {

                if (sfd.FilterIndex == 1)
                    chart2.SaveImage(sfd.FileName, System.Drawing.Imaging.ImageFormat.Bmp);
                if (sfd.FilterIndex == 2)
                    chart2.SaveImage(sfd.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                if (sfd.FilterIndex == 3)
                    chart2.SaveImage(sfd.FileName, System.Drawing.Imaging.ImageFormat.Png);

            }
        }

        private void refreshrealchart()
        {

            conn = new SQLiteConnection("Data Source = " + dbPath);
            conn.Open();
            //选择所有数据
            SQLiteDataAdapter mAdapter = new SQLiteDataAdapter("select * from RECORD order by ID desc limit 0," + maskedTextBox3.Text, conn);
            System.Data.DataTable dt = new System.Data.DataTable();
            mAdapter.Fill(dt);
            //绑定数据到DataGridView
            chart2.DataSource = dt;
            /*
             * C#中多线程更新Chart控件与BeginInvoke
最近把之前修改的MFC平台上的监控程序移植到C#上，需要用到图形控件显示监控曲线，C#中的现成的Chart控件为首选，但是在后台线程中更新Chart数据是总是在接收数据并刷新Chart时Chart控件上的图形变成一个大红叉，如下图所示，一个下午都没查出来为什么，后来在论坛上看到有人说需要利用BeginInvoke委托，就去看了MSDN上的介绍，按照使用说明改写后真的能刷新了，写下代码，以备不时之需。
             * */
            chart2.DataBind();
            conn.Close();

        }
        private void timer1_Tick(object sender, EventArgs e)
        {



            chart2.BeginInvoke(new ChartDelegate(refreshrealchart));


            //refreshrealchart();


        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (Convert.ToInt32(maskedTextBox1.Text) < 100)
            {
                maskedTextBox1.Text = "500";

            }
            timerchart.Interval = Convert.ToInt32(maskedTextBox1.Text);
            createrealtimechartseries(longarray);
            timerchart.Enabled = true;

        }

        private void button18_Click(object sender, EventArgs e)
        {
            timerchart.Enabled = false;
        }

        public void checkAllState(bool check)
        {
            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {
                this.checkedListBox1.SetItemChecked(i, check);
            }
        }
        public void checkAllState2(bool check)
        {
            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {
                this.checkedListBox2.SetItemChecked(i, check);
            }
        }
        public void checkAllState3(bool check)
        {
            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {
                this.checkedListBox3.SetItemChecked(i, check);
            }
        }
        public void checkAllState4(bool check)
        {
            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {
                this.checkedListBox4.SetItemChecked(i, check);
            }
        }

        private void checkedListBox2_ItemCheck(object sender, ItemCheckEventArgs e)
        {

            Boolean checked1 = false;
            CheckedListBox clb = (CheckedListBox)sender;
            checked1 = clb.GetItemChecked(e.Index);
            if (checked1)
            {

                chart2.Series[clb.Items[e.Index].ToString()].Enabled = false;
            }
            else
            {
                chart2.Series[clb.Items[e.Index].ToString()].Enabled = true;

            }
        }

        private void chart1_GetToolTipText(object sender, ToolTipEventArgs e)
        {
            //鼠标移动到值上，显示数值
            if (e.HitTestResult.ChartElementType == ChartElementType.DataPoint)
            {
                this.Cursor = Cursors.Cross;
                int i = e.HitTestResult.PointIndex;
                DataPoint dp = e.HitTestResult.Series.Points[i];
                e.Text = string.Format(e.HitTestResult.Series.Name + "\n数值:{1:F3}" + " \n日期:{0}", DateTime.FromOADate(dp.XValue), dp.YValues[0]);

            }
            else
            {
                this.Cursor = Cursors.Default;
            }

        }

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            var area = chart1.ChartAreas[0];
            try
            {

                double xValue = area.AxisX.PixelPositionToValue(e.X);
                double yValue = area.AxisY.PixelPositionToValue(e.Y);
                labX.Text = string.Format("{0:F0},{1:F0}", xValue, yValue);
                labY.Text = string.Format("{0:F0},{1:F0}", e.X, e.Y);
            }
            catch (Exception ex)
            {
                // this.textBox1.Text = ex.ToString();
                MessageBox.Show(e.X.ToString() + e.X.ToString());
            }


        }

        private void chart2_GetToolTipText(object sender, ToolTipEventArgs e)
        {
            //鼠标移动到值上，显示数值
            if (e.HitTestResult.ChartElementType == ChartElementType.DataPoint)
            {
                this.Cursor = Cursors.Cross;
                int i = e.HitTestResult.PointIndex;
                DataPoint dp = e.HitTestResult.Series.Points[i];
                e.Text = string.Format(e.HitTestResult.Series.Name + "\n数值:{1:F3}" + " \n日期:{0}", DateTime.FromOADate(dp.XValue), dp.YValues[0]);

            }
            else
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void chart2_MouseMove(object sender, MouseEventArgs e)
        {

            var area = chart2.ChartAreas[0];
            double xValue = area.AxisX.PixelPositionToValue(e.X);
            double yValue = area.AxisY.PixelPositionToValue(e.Y);
            labrealx.Text = string.Format("{0:F0},{1:F0}", xValue, yValue);
            labrealy.Text = string.Format("{0:F0},{1:F0}", e.X, e.Y);
        }




        private void Releasehelp()
        {

            string str1 = " ";
            str1 = System.Threading.Thread.GetDomain().BaseDirectory + "help";
            if (Directory.Exists(str1) == false)//如果不存
            {
                Directory.CreateDirectory(str1);
            }

            string strPath = str1 + @"\help.rtf";//设置释放路径  
            if (!File.Exists(strPath))
            {

                Assembly asm = Assembly.GetExecutingAssembly();
                string ResourceName = "CSClient.help.rtf";
                Stream pStream = asm.GetManifestResourceStream(ResourceName);
                byte[] buffer = new byte[pStream.Length];

                pStream.Read(buffer, 0, buffer.Length);


                //创建文件（覆盖模式）  
                using (FileStream fs = new FileStream(strPath, FileMode.Create))
                {
                    fs.Write(buffer, 0, buffer.Length);
                }


            }

            richTextBox2.LoadFile(strPath);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!Client.Connected)
            {
                MessageBox.Show("not connected");
                return;
            }
            button1.Enabled = false;
            button3.Enabled = false;
            button5.Enabled = false;
            button10.Enabled = false;
            button11.Enabled = false;
            button12.Enabled = false;
            button24.Enabled = false;
            dbcreate.Enabled = false;
            listView1.Enabled = false;
            dataGridView1.Enabled = false;
            realtimerecord.Interval = Convert.ToInt32(maskedTextBox2.Text);
            realtimerecord.Enabled = true;

        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (dataGridView2.RowCount > 2)
            {
                ExportExcels(dataGridView2);
            }

        }
        private void tooltipInfo1()
        {
            while (true)
            {
                ToolTipInfo1 tip1 = new ToolTipInfo1(
                delegate
                {

                    if (Client == null)
                    {
                        tooltipstr1 = "    ---------";
                    }
                    else
                    {
                        connectedFlag = Client.Connected;
                        if (connectedFlag) tooltipstr1 = "    connected";
                        else tooltipstr1 = "not connected";
                    }
                    toolStripStatusLabel1.Text = tooltipstr1;
                });
                tip1();
                Thread.Sleep(1000);
            }
            // ReSharper disable once FunctionNeverReturns
        }

        private void tooltipInfo2()
        {
            while (true)
            {
                ToolTipInfo2 tip2 = new ToolTipInfo2(
                delegate
                {
                    //toolStripStatusLabel2.Text = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff");
                    toolStripStatusLabel2.Text = new string(chararray);

                });
                tip2();
                Thread.Sleep(1000);
            }
            // ReSharper disable once FunctionNeverReturns
        }

        private void tooltipInfo3()
        {
            while (true)
            {
                ToolTipInfo3 tip3 = new ToolTipInfo3(
                delegate
                {
                    toolStripStatusLabel3.Text = dbPath;
                });
                tip3();
                Thread.Sleep(1000);
            }
            // ReSharper disable once FunctionNeverReturns
        }



        /// <summary>
        /// 线程调用
        /// </summary>
        private void DateTimeInfo()
        {
            while (true)
            {
                SetDateTime setDate = new SetDateTime(
                delegate
                {
                    //toolStripStatusLabel2.Text = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff");
                    toolStripStatusLabel2.Text = new string(chararray);

                    if (Client == null)
                    {
                        tooltipstr1 = "    ---------";
                    }
                    else
                    {
                        connectedFlag = Client.Connected;
                        if (connectedFlag) tooltipstr1 = "    connected";
                        else tooltipstr1 = "not connected";
                    }
                    toolStripStatusLabel1.Text = tooltipstr1;
                    toolStripStatusLabel3.Text = dbPath;

                });
                setDate();
                Thread.Sleep(1000);
            }
            // ReSharper disable once FunctionNeverReturns
        }

        private void button20_Click(object sender, EventArgs e)
        {
            /*
C# 变相的实现可变大小的二维数组（string） (2012-08-09 20:52:54)转载▼
标签： 可变大小 二维数组 c it	分类： Asp.Net
估计大家看个例子就明白了。(用了ArrayList）

举例说明
ArrayList list = new ArrayList();
for(int i=0;i<5;i++)
{
    string[] info = new string[3];

    info[0]=i;

    info[1] = 2*i;

    info[3] = 3*i;

    list.add(info);
}
下面通过强制类型转换得到第一条info的字段:

string[] info = (string[])list[0];
这样就可以像操作数组一样操作Info了。
             * */
            int[] a;
            int j;
            //   richTextBox2.Clear();
            ArrayList list = new ArrayList();
            list.Clear();
            string str1 = "";
            List<int> item = new List<int>(new int[] { 3, 4, 5, 6, 1, 2, 3, 4, 5, 6, 7, 8, 9, 1, 2, 3, 4, 5, 6, 7, 8, 9, 1, 2, 3, 4, 5, 6, 7 });
            for (int k = 0; k < 2000; k++)
            {
                item[0] = k;
                a = item.ToArray();
                list.Add(a);
            }
            /*
             foreach (int[] tt in list)
             {
                 j = 0;
                 foreach(  int i in tt)
                 {
                     j = j++;                   
                     str1 = str1 + i.ToString() + " ";              
                 }
                 str1 = str1 + "\n";
             }
             */
            // richTextBox2.AppendText(str1+"\n");
            richTextBox2.AppendText("1");
        }

        private void panel13_Paint(object sender, PaintEventArgs e)
        {

        }

        private void ConnectBtn_Click(object sender, EventArgs e)
        {
            int Result;
            int Rack = System.Convert.ToInt32(TxtRack.Text);
            int Slot = System.Convert.ToInt32(TxtSlot.Text);
            Result = Client.ConnectTo(TxtIP.Text, Rack, Slot);
            if (Result == 0)
                connectedFlag = true;
            ShowResult(Result);

        }
        private void ShowResult(int Result)
        {
            // This function returns a textual explaination of the error code
            tooltipstr2 = Client.ErrorText(Result);
            if (Result == 0)
                tooltipstr2 = tooltipstr2 + " (" + Client.ExecutionTime.ToString() + " ms)";
            if (tooltipstr2.Length < 200) chararray = tooltipstr2.ToCharArray();
        }
        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox3_KeyDown(object sender, KeyEventArgs e)
        {
            MaskedTextBox mtb;
            if (sender is MaskedTextBox)
            {
                mtb = sender as MaskedTextBox;

                if (e.KeyCode == Keys.Decimal)
                {
                    int pos = mtb.SelectionStart;
                    int max = (mtb.MaskedTextProvider.Length - mtb.MaskedTextProvider.EditPositionCount);
                    int nextField = 0;

                    for (int i = 0; i < mtb.MaskedTextProvider.Length; i++)
                    {
                        if (!mtb.MaskedTextProvider.IsEditPosition(i) && (pos + max) >= i)
                            nextField = i;
                    }
                    nextField += 1;

                    mtb.SelectionStart = nextField;

                }

            }
            else return;


        }

        private void DisconnectBtn_Click(object sender, EventArgs e)
        {
            Client.Disconnect();
        }

        private void ReadOrderCodeBtn_Click(object sender, EventArgs e)
        {
            ReadOrderCode();
        }
        void ReadCPUInfo()
        {
            S7Client.S7CpuInfo Info = new S7Client.S7CpuInfo();
            txtModuleTypeName.Text = "";
            txtSerialNumber.Text = "";
            txtCopyright.Text = "";
            txtAsName.Text = "";
            txtModuleName.Text = "";
            int Result = Client.GetCpuInfo(ref Info);
            ShowResult(Result);
            if (Result == 0)
            {
                txtModuleTypeName.Text = Info.ModuleTypeName;
                txtSerialNumber.Text = Info.SerialNumber;
                txtCopyright.Text = Info.Copyright;
                txtAsName.Text = Info.ASName;
                txtModuleName.Text = Info.ModuleName;
            }
        }
        void ReadOrderCode()
        {
            S7Client.S7OrderCode Info = new S7Client.S7OrderCode();
            txtOrderCode.Text = "";
            txtVersion.Text = "";
            int Result = Client.GetOrderCode(ref Info);
            ShowResult(Result);
            if (Result == 0)
            {
                txtOrderCode.Text = Info.Code;
                txtVersion.Text = Info.V1.ToString() + "." + Info.V2.ToString() + "." + Info.V3.ToString();
            }
        }

        private void ReadCPUInfoBtn_Click(object sender, EventArgs e)
        {
            ReadCPUInfo();
        }

        private void SetDateTimeBtn_Click(object sender, EventArgs e)
        {
            ShowResult(Client.SetPlcSystemDateTime());
        }

        void ShowPlcStatus()
        {
            int Status = 0;
            int Result = Client.PlcGetStatus(ref Status);
            ShowResult(Result);
            if (Result == 0)
            {
                switch (Status)
                {
                    case S7Consts.S7CpuStatusRun:
                        {
                            lblStatus.Text = "RUN";
                            lblStatus.ForeColor = System.Drawing.Color.LimeGreen;
                            break;
                        }
                    case S7Consts.S7CpuStatusStop:
                        {
                            lblStatus.Text = "STOP";
                            lblStatus.ForeColor = System.Drawing.Color.Red;
                            break;
                        }
                    default:
                        {
                            lblStatus.Text = "Unknown";
                            lblStatus.ForeColor = System.Drawing.Color.Black;
                            break;
                        }
                }
            }
            else
            {
                lblStatus.Text = "Unknown";
                lblStatus.ForeColor = System.Drawing.Color.Black;
            }
        }
        private void ReadDateTimeBtn_Click(object sender, EventArgs e)
        {
            DateTime DT = new DateTime();
            if (Client.GetPlcDateTime(ref DT) == 0)
            {
                txtDateTime.Text = DT.ToLongDateString() + " - " + DT.ToLongTimeString();
            }
        }

        private void PlcStatusBtn_Click(object sender, EventArgs e)
        {
            ShowPlcStatus();
        }

        private void PlcStopBtn_Click(object sender, EventArgs e)
        {
            ShowResult(Client.PlcStop());
            ShowPlcStatus();
        }

        private void PlcHotBtn_Click(object sender, EventArgs e)
        {
            ShowResult(Client.PlcHotStart());
            ShowPlcStatus();
        }

        private void PlcColdBtn_Click(object sender, EventArgs e)
        {
            ShowResult(Client.PlcColdStart());
            ShowPlcStatus();
        }

        private void button21_Click(object sender, EventArgs e)
        {



            if (mainform1 == null || mainform1.IsDisposed)
            {
                mainform1 = new MainForm();
                mainform1.xmldate1 = this;
                mainform1.Show();
            }
            else
            {
                mainform1.Activate();
            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button22_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = this.tabPage5;

        }

        private void button23_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = this.tabPage4;

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void button24_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = this.tabPage1;
        }

        private void realtimerecord_Tick(object sender, EventArgs e)
        {
            if (!Client.Connected)
            {
                realtimerecord.Enabled = false;
                MessageBox.Show("not connected");

                return;
            }


            createarrayfromgv();


            //  realtimerecord.Enabled = false;
        }

        private void HexDump(System.Windows.Forms.TextBox DumpBox, byte[] bytes, int Size)
        {
            if (bytes == null)
                return;
            int bytesLength = Size;
            int bytesPerLine = 16;

            char[] HexChars = "0123456789ABCDEF".ToCharArray();

            int firstHexColumn =
                  8                   // 8 characters for the address
                + 3;                  // 3 spaces

            int firstCharColumn = firstHexColumn
                + bytesPerLine * 3       // - 2 digit for the hexadecimal value and 1 space
                + (bytesPerLine - 1) / 8 // - 1 extra space every 8 characters from the 9th
                + 2;                  // 2 spaces 

            int lineLength = firstCharColumn
                + bytesPerLine           // - characters to show the ascii value
                + Environment.NewLine.Length; // Carriage return and line feed (should normally be 2)

            char[] line = (new String(' ', lineLength - 2) + Environment.NewLine).ToCharArray();
            int expectedLines = (bytesLength + bytesPerLine - 1) / bytesPerLine;
            StringBuilder result = new StringBuilder(expectedLines * lineLength);

            for (int i = 0; i < bytesLength; i += bytesPerLine)
            {
                line[0] = HexChars[(i >> 28) & 0xF];
                line[1] = HexChars[(i >> 24) & 0xF];
                line[2] = HexChars[(i >> 20) & 0xF];
                line[3] = HexChars[(i >> 16) & 0xF];
                line[4] = HexChars[(i >> 12) & 0xF];
                line[5] = HexChars[(i >> 8) & 0xF];
                line[6] = HexChars[(i >> 4) & 0xF];
                line[7] = HexChars[(i >> 0) & 0xF];

                int hexColumn = firstHexColumn;
                int charColumn = firstCharColumn;

                for (int j = 0; j < bytesPerLine; j++)
                {
                    if (j > 0 && (j & 7) == 0) hexColumn++;
                    if (i + j >= bytesLength)
                    {
                        line[hexColumn] = ' ';
                        line[hexColumn + 1] = ' ';
                        line[charColumn] = ' ';
                    }
                    else
                    {
                        byte b = bytes[i + j];
                        line[hexColumn] = HexChars[(b >> 4) & 0xF];
                        line[hexColumn + 1] = HexChars[b & 0xF];
                        line[charColumn] = (b < 32 ? '·' : (char)b);
                    }
                    hexColumn += 3;
                    charColumn++;
                }
                result.Append(line);
            }
            DumpBox.Text = DumpBox.Text + result.ToString();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            // true;

            button1.Enabled = true;
            button3.Enabled = true;
            button5.Enabled = true;
            button10.Enabled = true;
            button11.Enabled = true;
            button12.Enabled = true;
            button24.Enabled = true;
            dbcreate.Enabled = true;
            listView1.Enabled = true;
            dataGridView1.Enabled = true;
            realtimerecord.Enabled = false;
        }

        private void button26_Click(object sender, EventArgs e)
        {
            refreshgrid();
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkbox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkbox1.Checked)
            {
                checkAllState(true);
            }
            else
            {
                checkAllState(false);
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                checkAllState2(true);
            }
            else
            {
                checkAllState2(false);
            }

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void checkedListBox3_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            Boolean checked1 = false;
            CheckedListBox clb = (CheckedListBox)sender;
            checked1 = clb.GetItemChecked(e.Index);
            if (checked1)
            {

                chart1.Series[clb.Items[e.Index].ToString()].YAxisType = AxisType.Secondary;
            }
            else
            {
                chart1.Series[clb.Items[e.Index].ToString()].YAxisType = AxisType.Primary;

            }


        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

            if (checkBox3.Checked)
            {
                checkAllState3(true);
            }
            else
            {
                checkAllState3(false);
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                checkAllState4(true);
            }
            else
            {
                checkAllState4(false);
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {

            if (dateTimePicker2.Value > dateTimePicker1.Value)
            {
                chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Auto;
                //  chart1.ChartAreas[0].AxisX.Interval = (dateTimePicker2.Value - dateTimePicker1.Value).Days / 2;
                chart1.ChartAreas[0].AxisX.Minimum = dateTimePicker1.Value.ToOADate();
                chart1.ChartAreas[0].AxisX.Maximum = dateTimePicker2.Value.ToOADate();
                chart1.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
                chart1.ChartAreas[0].AxisX.ScaleView.ZoomReset();
                // chart1.ChartAreas[0].AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount;
            }


        }

        private void button28_Click(object sender, EventArgs e)
        {
            chart1.ChartAreas[0].AxisY.Minimum = Convert.ToInt32(textBox5.Text);
            chart1.ChartAreas[0].AxisY.Maximum = Convert.ToInt32(textBox4.Text);
        }

        private void button29_Click(object sender, EventArgs e)
        {
            chart1.ChartAreas[0].AxisY2.Minimum = Convert.ToInt32(textBox7.Text);
            chart1.ChartAreas[0].AxisY2.Maximum = Convert.ToInt32(textBox6.Text);
        }

        private void button31_Click(object sender, EventArgs e)
        {
            chart2.ChartAreas[0].AxisY.Minimum = Convert.ToInt32(textBox11.Text);
            chart2.ChartAreas[0].AxisY.Maximum = Convert.ToInt32(textBox10.Text);
        }

        private void button30_Click(object sender, EventArgs e)
        {
            chart2.ChartAreas[0].AxisY2.Minimum = Convert.ToInt32(textBox9.Text);
            chart2.ChartAreas[0].AxisY2.Maximum = Convert.ToInt32(textBox8.Text);
        }

        private void button32_Click(object sender, EventArgs e)
        {
            // chart1.ChartAreas[0].AxisY.Minimum = Convert.ToInt32(textBox5.Text);
            //  chart1.ChartAreas[0].AxisY.Maximum = Convert.ToInt32(textBox4.Text);
        }

        private void button38_Click(object sender, EventArgs e)
        {
            if (checkedListBox2.Items.Count == 0) return;
            chart2.ChartAreas[0].AxisX.Maximum = System.Double.NaN;
            chart2.ChartAreas[0].AxisX.Minimum = System.Double.NaN;
            chart2.ChartAreas[0].AxisX.ScaleView.ZoomReset(1);

        }

        private void button37_Click(object sender, EventArgs e)
        {
            if (checkedListBox2.Items.Count == 0) return;
            chart2.ChartAreas[0].AxisY.Maximum = System.Double.NaN;
            chart2.ChartAreas[0].AxisY.Minimum = System.Double.NaN;
            chart2.ChartAreas[0].AxisY.ScaleView.ZoomReset(1);
        }

        private void button36_Click(object sender, EventArgs e)
        {

            if (checkedListBox2.Items.Count == 0) return;
            chart2.ChartAreas[0].AxisY2.Maximum = System.Double.NaN;
            chart2.ChartAreas[0].AxisY2.Minimum = System.Double.NaN;
            chart2.ChartAreas[0].AxisY2.ScaleView.ZoomReset(1);
        }

        private void button33_Click(object sender, EventArgs e)
        {
         if    (checkedListBox1.Items.Count==0)   return;
            // chart1.ChartAreas[0].AxisX.Maximum = System.Double.NaN;
            // chart1.ChartAreas[0].AxisX.Minimum = System.Double.NaN;
            //   chart1.ChartAreas[0].AxisX.ScaleView.ZoomReset(0);
            //ft1 = chart1.Series[0].Points[chart1.Series[0].Points.Count - 1].XValue;
            chart1.ChartAreas[0].AxisX.ScaleView.Size = 2000;// Convert.ToInt32( textBox2.Text);// 可视区域数据点数
            chart1.ChartAreas[0].AxisX.ScaleView.SizeType = DateTimeIntervalType.Number;
            // chart1.ChartAreas[0].AxisX.ScaleView.ZoomReset();

        }

        private void button34_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.Items.Count == 0) return;
            chart1.ChartAreas[0].AxisY.Maximum = System.Double.NaN;
            chart1.ChartAreas[0].AxisY.Minimum = System.Double.NaN;
            chart1.ChartAreas[0].AxisY.ScaleView.ZoomReset(1);
            //chart1.ChartAreas[0].AxisX.ScaleView.ZoomReset(1); —— 撤销一次放大动作

            // chart1.ChartAreas[0].AxisX.ScaleView.ZoomReset(0); —— 撤销所有放大动作
        }

        private void button35_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.Items.Count == 0) return;
            chart1.ChartAreas[0].AxisY2.Maximum = System.Double.NaN;
            chart1.ChartAreas[0].AxisY2.Minimum = System.Double.NaN;
            chart1.ChartAreas[0].AxisY2.ScaleView.ZoomReset(1);
        }

        private void checkedListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkedListBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkedListBox4_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            Boolean checked1 = false;
            CheckedListBox clb = (CheckedListBox)sender;
            checked1 = clb.GetItemChecked(e.Index);
            if (checked1)
            {

                chart2.Series[clb.Items[e.Index].ToString()].YAxisType = AxisType.Secondary;
            }
            else
            {
                chart2.Series[clb.Items[e.Index].ToString()].YAxisType = AxisType.Primary;

            }
        }

        private void maskedTextBox3_MaskInputRejected_1(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void button40_Click(object sender, EventArgs e)
        {
            currentnumber = currentnumber + 1;
            createchartseries(longarray);
        }

        private void button39_Click(object sender, EventArgs e)
        {
            currentnumber = currentnumber - 1;
            createchartseries(longarray);
        }

        private void button41_Click(object sender, EventArgs e)
        {
            savechartimage();
        }
    }


    public class SystemInfo
    {

        public virtual List<string> GetMonitorPnpDeviceId()
        {
            List<string> rt = new List<string>();

            using (ManagementClass mc = new ManagementClass("Win32_DesktopMonitor"))
            {
                using (ManagementObjectCollection moc = mc.GetInstances())
                {
                    foreach (var o in moc)
                    {
                        var each = (ManagementObject)o;
                        object obj = each.Properties["PNPDeviceID"].Value;
                        if (obj == null)
                            continue;

                        rt.Add(each.Properties["PNPDeviceID"].Value.ToString());
                    }
                }
            }

            return rt;
        }

        public virtual byte[] GetMonitorEdid(string monitorPnpDevId)
        {
            return (byte[])Registry.GetValue(@"HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Enum\" + monitorPnpDevId + @"\Device Parameters", "EDID", new byte[] { });
        }

        //获取显示器物理尺寸(cm)
        public virtual SizeF GetMonitorPhysicalSize(string monitorPnpDevId)
        {
            byte[] edid = GetMonitorEdid(monitorPnpDevId);
            if (edid.Length < 23)
                return SizeF.Empty;

            return new SizeF(edid[21], edid[22]);
        }

        //通过屏显示器理尺寸转换为显示器大小(inch)
        public static float MonitorScaler(SizeF moniPhySize)
        {
            double mDSize = Math.Sqrt(Math.Pow(moniPhySize.Width, 2) + Math.Pow(moniPhySize.Height, 2)) / 2.54d;
            return (float)Math.Round(mDSize, 1);
        }
    }


    public class PrimaryScreen
    {
        #region Win32 API
        [DllImport("user32.dll")]
        static extern IntPtr GetDC(IntPtr ptr);
        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(
        IntPtr hdc, // handle to DC
        int nIndex // index of capability
        );
        [DllImport("user32.dll", EntryPoint = "ReleaseDC")]
        static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDc);
        #endregion
        #region DeviceCaps常量
        const int HORZRES = 8;
        const int VERTRES = 10;
        const int LOGPIXELSX = 88;
        const int LOGPIXELSY = 90;
        const int DESKTOPVERTRES = 117;
        const int DESKTOPHORZRES = 118;
        #endregion

        #region 属性
        /// <summary>
        /// 获取屏幕分辨率当前物理大小
        /// </summary>
        public static Size WorkingArea
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                Size size = new Size();
                size.Width = GetDeviceCaps(hdc, HORZRES);
                size.Height = GetDeviceCaps(hdc, VERTRES);
                ReleaseDC(IntPtr.Zero, hdc);
                return size;
            }
        }
        /// <summary>
        /// 当前系统DPI_X 大小 一般为96
        /// </summary>
        public static int DpiX
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                int DpiX = GetDeviceCaps(hdc, LOGPIXELSX);
                ReleaseDC(IntPtr.Zero, hdc);
                return DpiX;
            }
        }
        /// <summary>
        /// 当前系统DPI_Y 大小 一般为96
        /// </summary>
        public static int DpiY
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                int DpiX = GetDeviceCaps(hdc, LOGPIXELSY);
                ReleaseDC(IntPtr.Zero, hdc);
                return DpiX;
            }
        }
        /// <summary>
        /// 获取真实设置的桌面分辨率大小
        /// </summary>
        public static Size DESKTOP
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                Size size = new Size();
                size.Width = GetDeviceCaps(hdc, DESKTOPHORZRES);
                size.Height = GetDeviceCaps(hdc, DESKTOPVERTRES);
                ReleaseDC(IntPtr.Zero, hdc);
                return size;
            }
        }

        /// <summary>
        /// 获取宽度缩放百分比
        /// </summary>
        public static float ScaleX
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                int t = GetDeviceCaps(hdc, DESKTOPHORZRES);
                int d = GetDeviceCaps(hdc, HORZRES);
                float ScaleX = (float)GetDeviceCaps(hdc, DESKTOPHORZRES) / (float)GetDeviceCaps(hdc, HORZRES);
                ReleaseDC(IntPtr.Zero, hdc);
                return ScaleX;
            }
        }
        /// <summary>
        /// 获取高度缩放百分比
        /// </summary>
        public static float ScaleY
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                float ScaleY = (float)(float)GetDeviceCaps(hdc, DESKTOPVERTRES) / (float)GetDeviceCaps(hdc, VERTRES);
                ReleaseDC(IntPtr.Zero, hdc);
                return ScaleY;
            }
        }
        #endregion
    }


    public class GraphicsText
    {
        private Graphics _graphics;

        public GraphicsText()
        {

        }

        public Graphics Graphics
        {
            get { return _graphics; }
            set { _graphics = value; }
        }

        /// <summary>
        /// 绘制根据矩形旋转文本
        /// </summary>
        /// <param name="s">文本</param>
        /// <param name="font">字体</param>
        /// <param name="brush">填充</param>
        /// <param name="layoutRectangle">局部矩形</param>
        /// <param name="format">布局方式</param>
        /// <param name="angle">角度</param>
        public void DrawString(string s, System.Drawing.Font font, Brush brush, RectangleF layoutRectangle, StringFormat format, float angle)
        {
            // 求取字符串大小
            SizeF size = _graphics.MeasureString(s, font);

            // 根据旋转角度，求取旋转后字符串大小
            SizeF sizeRotate = ConvertSize(size, angle);

            // 根据旋转后尺寸、布局矩形、布局方式计算文本旋转点
            PointF rotatePt = GetRotatePoint(sizeRotate, layoutRectangle, format);

            // 重设布局方式都为Center
            StringFormat newFormat = new StringFormat(format);
            newFormat.Alignment = StringAlignment.Center;
            newFormat.LineAlignment = StringAlignment.Center;

            // 绘制旋转后文本
            DrawString(s, font, brush, rotatePt, newFormat, angle);
        }

        /// <summary>
        /// 绘制根据点旋转文本，一般旋转点给定位文本包围盒中心点
        /// </summary>
        /// <param name="s">文本</param>
        /// <param name="font">字体</param>
        /// <param name="brush">填充</param>
        /// <param name="point">旋转点</param>
        /// <param name="format">布局方式</param>
        /// <param name="angle">角度</param>
        public void DrawString(string s, System.Drawing.Font font, Brush brush, PointF point, StringFormat format, float angle)
        {
            // Save the matrix
            Matrix mtxSave = _graphics.Transform;

            Matrix mtxRotate = _graphics.Transform;
            mtxRotate.RotateAt(angle, point);
            _graphics.Transform = mtxRotate;

            _graphics.DrawString(s, font, brush, point, format);

            // Reset the matrix
            _graphics.Transform = mtxSave;
        }

        private SizeF ConvertSize(SizeF size, float angle)
        {
            Matrix matrix = new Matrix();
            matrix.Rotate(angle);

            // 旋转矩形四个顶点
            PointF[] pts = new PointF[4];
            pts[0].X = -size.Width / 2f;
            pts[0].Y = -size.Height / 2f;
            pts[1].X = -size.Width / 2f;
            pts[1].Y = size.Height / 2f;
            pts[2].X = size.Width / 2f;
            pts[2].Y = size.Height / 2f;
            pts[3].X = size.Width / 2f;
            pts[3].Y = -size.Height / 2f;
            matrix.TransformPoints(pts);

            // 求取四个顶点的包围盒
            float left = float.MaxValue;
            float right = float.MinValue;
            float top = float.MaxValue;
            float bottom = float.MinValue;

            foreach (PointF pt in pts)
            {
                // 求取并集
                if (pt.X < left)
                    left = pt.X;
                if (pt.X > right)
                    right = pt.X;
                if (pt.Y < top)
                    top = pt.Y;
                if (pt.Y > bottom)
                    bottom = pt.Y;
            }

            SizeF result = new SizeF(right - left, bottom - top);
            return result;
        }

        private PointF GetRotatePoint(SizeF size, RectangleF layoutRectangle, StringFormat format)
        {
            PointF pt = new PointF();

            switch (format.Alignment)
            {
                case StringAlignment.Near:
                    pt.X = layoutRectangle.Left + size.Width / 2f;
                    break;
                case StringAlignment.Center:
                    pt.X = (layoutRectangle.Left + layoutRectangle.Right) / 2f;
                    break;
                case StringAlignment.Far:
                    pt.X = layoutRectangle.Right - size.Width / 2f;
                    break;
                default:
                    break;
            }

            switch (format.LineAlignment)
            {
                case StringAlignment.Near:
                    pt.Y = layoutRectangle.Top + size.Height / 2f;
                    break;
                case StringAlignment.Center:
                    pt.Y = (layoutRectangle.Top + layoutRectangle.Bottom) / 2f;
                    break;
                case StringAlignment.Far:
                    pt.Y = layoutRectangle.Bottom - size.Height / 2f;
                    break;
                default:
                    break;
            }

            return pt;
        }
    }


}