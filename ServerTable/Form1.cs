using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text.pdf;

namespace ServerTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string Template, drivetype="监控点位";

        private void Form1_Load(object sender, EventArgs e)
        {
            Template = Application.StartupPath + @"\Template\Template.pdf";
            //TempPDF = Application.StartupPath + @"Template\Template.pdf";
            tbpath.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)+@"\Export\";
            string path = Application.StartupPath + @"\TW.txt";

            if (System.IO.File.Exists(path))
            {
                tbrwdd.Text = File.ReadAllText(path);
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {

            listView.Items.Clear();
            listView.BeginUpdate();
            int X = Convert.ToInt32(tbnum.Text);
            for (int i = 0; i < X; i++)
            {
                ListViewItem lvItem = new ListViewItem();
                lvItem.Text = (i + 1).ToString();

                Random rdRandom = new Random(Guid.NewGuid().GetHashCode());//生成随机数

                string[] users = tbuser.Text.Split(new char[] { '|' }); //分割用户名
                int userid = rdRandom.Next(0, users.Length);
                // System.Threading.Thread.Sleep(5);
                lvItem.SubItems.Add(users[userid]); //取随机用户


                string[] enger = tbeng.Text.Split(new char[] { '|' }); //分割用户名
                int engid = rdRandom.Next(0, enger.Length);
                // System.Threading.Thread.Sleep(5);
                lvItem.SubItems.Add(enger[engid]); //取随机用户


                var userDatetime = dateTime.Value.AddDays(i);//生成随机时间
                if (radioButton1.Checked == true)
                {
                    lvItem.SubItems.Add(userDatetime.ToString("yyyy-MM-dd") + radioButton1.Text);
                }
                else
                {

                    lvItem.SubItems.Add(userDatetime.ToString("yyyy-MM-dd") + radioButton2.Text);
                }


                lvItem.SubItems.Add(tbzysx.Text); //取随机类型


                string[] servert = tbtype.Text.Split(new char[] { '|' }); //分割响应类型
                int servertid = rdRandom.Next(0, servert.Length);
                // System.Threading.Thread.Sleep(3);
                lvItem.SubItems.Add(servert[servertid]); //取随机类型


                int gznum = rdRandom.Next(1, 5);
                string[] serverAdd = tbrwdd.Text.Split(new char[] { '|' });//故障地点
                string[] gztype = tbrwms.Text.Split(new char[] { '|' });//故障描述

                string rwzs = " ";

                for (int j = 0; j < gznum; j++)
                {
                    int addid = rdRandom.Next(0, serverAdd.Length);//取地点
                    //System.Threading.Thread.Sleep(5);

                    int gzid = rdRandom.Next(0, gztype.Length);

                    string rwms = serverAdd[addid] + drivetype+ gztype[gzid] + "\n ";

                    rwzs += rwms;
                }
                lvItem.SubItems.Add(rwzs); //添加任务描述到列表


                string[] bjStrings = tbbjxx.Text.Split(new char[] { '|' });//部件种类型
                string bjsl = " ";

                for (int j = 0; j < bjStrings.Length; j++)
                {
                    int bjs = 0;
                    if (bjStrings[j] != "")
                    {
                        Regex regex = new Regex(bjStrings[j]);
                        foreach (Match match in regex.Matches(rwzs))
                        {
                            bjs += 1;
                        }

                        if (bjs > 0 && regex.Match(rwzs).Value != "")
                        {
                            bjsl += regex.Match(rwzs).Value + bjs.ToString() + "个\n ";
                        }
                    }
                }
                lvItem.SubItems.Add(bjsl);


                string[] ylwtStrings = tbylwt.Text.Split(new char[] { '|' });
                int ylwtid = rdRandom.Next(0, ylwtStrings.Length);

                lvItem.SubItems.Add(ylwtStrings[ylwtid]); //取随机类型


                listView.Items.Add(lvItem);

            }

            listView.EndUpdate();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            

            for (int i = 0; i < listView.Items.Count; i++)
            {
                PdfReader reader = new PdfReader(Template);
                BaseFont bsFont = BaseFont.CreateFont(@"C:\Windows\Fonts\simsun.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                ListViewItem lv = listView.Items[i];
                string FileNamePath = ExportPath(lv.SubItems[3].Text);
                FileStream fileStream = new FileStream(FileNamePath, FileMode.Create);
                
                //PDF字段操作
                PdfStamper stamper = new PdfStamper(reader, fileStream);
                AcroFields coderAcroFields = stamper.AcroFields;
                coderAcroFields.AddSubstitutionFont(bsFont);
                
                //PDF字段填充



               // coderAcroFields.SetFieldProperty("ProjectName", "textfont", 15, null);


                coderAcroFields.SetField("ProjectName", tbxmmc.Text);
                coderAcroFields.SetField("ServerAdd", tbfwdz.Text);
                coderAcroFields.SetField("User", lv.SubItems[1].Text);
                coderAcroFields.SetField("Tel", tblxdh.Text);
                coderAcroFields.SetField("Engineer", lv.SubItems[2].Text);
                coderAcroFields.SetField("OderTime", lv.SubItems[3].Text);
                coderAcroFields.SetField("Note", lv.SubItems[4].Text);
                coderAcroFields.SetField("ServerType", lv.SubItems[5].Text);
                coderAcroFields.SetField("Server", lv.SubItems[6].Text);
                coderAcroFields.SetField("Parts", lv.SubItems[7].Text);
                coderAcroFields.SetField("Others", lv.SubItems[8].Text);

                //stamper.FormFlattening = true;
               
                stamper.Close();
                reader.Close();
            }



            System.Diagnostics.Process.Start("Explorer.exe", tbpath.Text);



        }


        private void button5_Click(object sender, EventArgs e)
        {
            tbrwms.Text = "摄像头故障|收发器故障|电源故障|供电故障|电路故障|传输链路故障";
            string path = Application.StartupPath + @"\TW.txt";
            drivetype = "监控点位";
            if (System.IO.File.Exists(path))
            {
                tbrwdd.Text = File.ReadAllText(path);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            tbrwms.Text = "红绿灯配时方案修改|收发器故障|供电故障|空开故障";
            string path = Application.StartupPath + @"\FX.txt";
            drivetype = "非现设备";
            if (System.IO.File.Exists(path))
            {
                tbrwdd.Text = File.ReadAllText(path);
            }
        }

        private string ExportPath(string filename)
        {

            if (!Directory.Exists(tbpath.Text))
            {
                Directory.CreateDirectory(tbpath.Text);
            }


            return tbpath.Text+ filename + ".PDF";

        }

    }
}
