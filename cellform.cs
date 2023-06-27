using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using WdExportFormat = Microsoft.Office.Interop.Word.WdExportFormat;
using Word = Microsoft.Office.Interop.Word;

namespace proeval
{
    public partial class cellform : Form
    {
        public cell classcell = new cell();
        public string datapath = System.Environment.CurrentDirectory + "/";
        public cellform()
        {
            InitializeComponent();
        }
        private void subform2_Load(object sender, EventArgs e)
        {
            this.textBox2.Text = classcell.cellname;
            this.textBox3.Text = classcell.cellindex;
            this.textBox4.Text = classcell.cellCEO;
            this.textBox5.Text = classcell.cellCEOphone;
            update2lists();
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox1.Text = classcell.Getval(listBox1.SelectedItem.ToString()).ToString();
            if (listBox1.SelectedItem.ToString() == "严重质量问题" || listBox1.SelectedItem.ToString() == "重大安全事故")
                label2.Text = "提示:此要素为精细级门槛型子要素，通过请打1，不通过请打-1。";
            else if (listBox1.SelectedItem.ToString() == "装备自动化水平" || listBox1.SelectedItem.ToString() == "设备运行状态在线监控" || listBox1.SelectedItem.ToString() == "柔性化程度" || listBox1.SelectedItem.ToString() == "设备布局合理性")
                label2.Text = "提示:此要素为精益级门槛型子要素达到精益请打4，达到卓越请打5。";
            else if (listBox1.SelectedItem.ToString() == "产品智能检测水平")
                label2.Text = "提示:此要素为卓越级级门槛型子要素达到卓越请打5。";
            else if (listBox1.SelectedItem.ToString() == "试验（检测）成功率" || listBox1.SelectedItem.ToString() == "产品一次交检合格率")
                label2.Text = "产品一次交检合格率与试验（检测）成功率根据单元类型二选一，请剔除掉不符合单元类型的子要素";
            else
                this.label2.Text = "提示:";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex < 0)
                return;
            string name = this.listBox1.SelectedItem.ToString();
            double b = double.Parse(textBox1.Text);
            classcell.setval(name, b);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                classcell.setval(listBox1.SelectedItem.ToString(), int.Parse(textBox1.Text));
            }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (classcell.scorecheck() != "ok")
                MessageBox.Show(classcell.scorecheck());
            string cellscore = "单元最终分数为：" + classcell.scoring()[0].ToString();
            MessageBox.Show(cellscore);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            classcell.cellname = this.textBox2.Text;
            classcell.cellindex = this.textBox3.Text;
            classcell.cellCEO = this.textBox4.Text;
            classcell.cellCEOphone = this.textBox5.Text;
            FileStream fs = new FileStream(@".\" + classcell.cellname + ".dat", FileMode.Create);
            BinaryFormatter bf = new BinaryFormatter();
            bf.Serialize(fs, classcell);
            fs.Close();
            mainform ff;
            ff = (mainform)this.Owner;
            ff.updatecelllist();
            MessageBox.Show("单元评价信息已保存");

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            this.Text = this.textBox2.Text;
        }

        public void form2update()
        {
            this.textBox2.Text = classcell.cellname;
            this.textBox3.Text = classcell.cellindex;
            this.textBox4.Text = classcell.cellCEO;
            this.textBox5.Text = classcell.cellCEOphone;
        }


        private void button4_Click(object sender, EventArgs e)
        {
            label2.Text = "单元等级报告正在生成中，请稍后...";
            RadarDemo radarDemo1 = new RadarDemo();
            radarDemo1.mData = classcell.Ret_Dic1();
            radarDemo1.Show(datapath, "1");
            RadarDemo radarDemo2 = new RadarDemo();
            radarDemo2.mData = classcell.Ret_Dic2();
            radarDemo2.Show(datapath, "2");
            RadarDemo radarDemo3 = new RadarDemo();
            radarDemo3.mData = classcell.Ret_Dic3();
            radarDemo3.Show(datapath, "3");
            Gen_Word();
        }
        private void button5_Click_1(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex < 0)
                return;
            string name = listBox1.SelectedItem.ToString();
            if (name == "严重质量问题" || name == "重大安全事故" || name == "装备自动化水平" || name == "设备运行状态在线监控" || name == "柔性化程度" || name == "目视化管理" || name == "设备布局合理性" || name == "产品智能检测水平")
            {
                MessageBox.Show("该要素不可以被裁剪");
            }
            else
            {
                classcell.deleteditems.Add(name);
                update2lists();
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex < 0)
                return;
            string name = listBox2.SelectedItem.ToString();
            classcell.deleteditems.Remove(name);
            update2lists();
        }
        private void update2lists()
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            foreach (var item in classcell.WeightDic)
            {
                if (classcell.deleteditems.Contains(item.Key))
                {
                    listBox2.Items.Add(item.Key);
                    continue;
                }
                this.listBox1.Items.Add(item.Key);
            }
        }

        private Size beforeResizeSize = Size.Empty;
        protected override void OnResizeBegin(EventArgs e)
        {
            base.OnResizeBegin(e);
            beforeResizeSize = this.Size;
        }
        protected override void OnResizeEnd(EventArgs e)
        {
            base.OnResizeEnd(e);
            //窗口resize之后的大小
            Size endResizeSize = this.Size;
            //获得变化比例
            float percentWidth = (float)endResizeSize.Width / beforeResizeSize.Width;
            float percentHeight = (float)endResizeSize.Height / beforeResizeSize.Height;
            foreach (Control control in this.Controls)
            {
                if (control is DataGridView)
                    continue;
                //按比例改变控件大小
                control.Width = (int)(control.Width * percentWidth);
                control.Height = (int)(control.Height * percentHeight);
                //为了不使控件之间覆盖 位置也要按比例变化
                control.Left = (int)(control.Left * percentWidth);
                control.Top = (int)(control.Top * percentHeight);
            }
        }

        void Gen_Word()
        {
            object oMissing = Missing.Value;
            object oEndOfDoc = "\\endofdoc";
            /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = false;
            object oTemplate = datapath + "等级评分表.docx";
            oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);

            object oBookMark1 = "单元名称";
            oDoc.Bookmarks[ref oBookMark1].Range.Text = classcell.cellname;
            oBookMark1 = "单元编号";
            oDoc.Bookmarks[ref oBookMark1].Range.Text = classcell.cellindex;
            oBookMark1 = "单元负责人";
            oDoc.Bookmarks[ref oBookMark1].Range.Text = classcell.cellCEO;
            oBookMark1 = "负责人电话";
            oDoc.Bookmarks[ref oBookMark1].Range.Text = classcell.cellCEOphone;
            oBookMark1 = "评价日期";
            oDoc.Bookmarks[ref oBookMark1].Range.Text = DateTime.Now.ToLongDateString().ToString();

            //否决要素分数
            object oBookMark2 = "子要素裁剪";
            if (classcell.deleteditems.ToArray().Length > 0)
            {
                string tempS = "被裁剪掉的子要素有：";
                foreach (var item in classcell.deleteditems)
                {
                    tempS += item + " ";
                }
                oDoc.Bookmarks[ref oBookMark2].Range.Text = tempS + "。";
            }
            else
            {
                oDoc.Bookmarks[ref oBookMark2].Range.Text = "暂无被裁剪掉的子要素。";
            }
            oBookMark2 = "严重质量问题";
            oDoc.Bookmarks[ref oBookMark2].Range.Text = classcell.ScoreDic["严重质量问题"].ToString();
            oBookMark2 = "重大安全事故";
            oDoc.Bookmarks[ref oBookMark2].Range.Text = classcell.ScoreDic["重大安全事故"].ToString();

            //等级判定陈述 
            oBookMark2 = "初步等级判定陈述";
            if (classcell.ScoreDic["重大安全事故"] <= 0)
                oDoc.Bookmarks[ref oBookMark2].Range.Text = "根据否决型子要素的得分情况，重大安全事故得分未达到要求，因此，初步判定该生产试验过程不得参与卓越等级评价。";
            else if (classcell.ScoreDic["严重质量问题"] <= 0)
                oDoc.Bookmarks[ref oBookMark2].Range.Text = "根据否决型子要素的得分情况，严重质量问题未达到要求，因此，初步判定该生产试验过程不得参与卓越等级评价。";
            else
                oDoc.Bookmarks[ref oBookMark2].Range.Text = "根据否决型子要素的得分情况，门槛型子要素均达到要求，因此，初步判定该生产试验过程可以参与卓越等级评价。";

            oBookMark2 = "装备自动化水平";
            oDoc.Bookmarks[ref oBookMark2].Range.Text = classcell.ScoreDic["装备自动化水平"].ToString();
            oBookMark2 = "设备运行状态在线监控";
            oDoc.Bookmarks[ref oBookMark2].Range.Text = classcell.ScoreDic["设备运行状态在线监控"].ToString();
            oBookMark2 = "设备布局合理性";
            oDoc.Bookmarks[ref oBookMark2].Range.Text = classcell.ScoreDic["设备布局合理性"].ToString();
            oBookMark2 = "柔性化程度";
            oDoc.Bookmarks[ref oBookMark2].Range.Text = classcell.ScoreDic["柔性化程度"].ToString();
            oBookMark2 = "目视化管理";
            oDoc.Bookmarks[ref oBookMark2].Range.Text = classcell.ScoreDic["目视化管理"].ToString();
            oBookMark2 = "产品智能检测水平";
            oDoc.Bookmarks[ref oBookMark2].Range.Text = classcell.ScoreDic["产品智能检测水平"].ToString();
            oBookMark2 = "等级判定1";
            object oBookMark3 = "权重占比图";
            object oBookMark4 = "权重占比2";

            Word.InlineShape pictureRank;
            if (classcell.isqualified())
                oDoc.Bookmarks[ref oBookMark2].Range.Text = "有资格参评";
            else
            {
                oDoc.Bookmarks[ref oBookMark2].Range.Text = "无资格参评";
                pictureRank = oDoc.InlineShapes.AddPicture(datapath + "0.png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark3].Range);
                oDoc.Bookmarks[ref oBookMark4].Range.Text = "N/A";
            }
            oBookMark2 = "等级判定2";
            switch (classcell.set_final_weight())
            {
                case "卓越级":
                    oDoc.Bookmarks[ref oBookMark2].Range.Text = "卓越级";
                    pictureRank = oDoc.InlineShapes.AddPicture(datapath + "3.png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark3].Range);
                    oDoc.Bookmarks[ref oBookMark4].Range.Text = "卓越级";
                    pictureRank.Height = (float)(pictureRank.Height * 0.75);
                    pictureRank.Width = (float)(pictureRank.Width * 0.75);
                    break;
                case "精益级":
                    oDoc.Bookmarks[ref oBookMark2].Range.Text = "精益级";
                    pictureRank = oDoc.InlineShapes.AddPicture(datapath + "2.png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark3].Range);
                    oDoc.Bookmarks[ref oBookMark4].Range.Text = "精益级";
                    pictureRank.Height = (float)(pictureRank.Height * 0.75);
                    pictureRank.Width = (float)(pictureRank.Width * 0.75);
                    break;
                case "精细级":
                    oDoc.Bookmarks[ref oBookMark2].Range.Text = "精细级";
                    pictureRank = oDoc.InlineShapes.AddPicture(datapath + "1.png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark3].Range);
                    oDoc.Bookmarks[ref oBookMark4].Range.Text = "精细级";
                    pictureRank.Height = (float)(pictureRank.Height * 0.75);
                    pictureRank.Width = (float)(pictureRank.Width * 0.75);
                    break;
                case "无评级":
                    oDoc.Bookmarks[ref oBookMark2].Range.Text = "无评级";
                    pictureRank = oDoc.InlineShapes.AddPicture(datapath + "0.png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark3].Range);
                    oDoc.Bookmarks[ref oBookMark4].Range.Text = "无评级";
                    pictureRank.Height = (float)(pictureRank.Height * 0.75);
                    pictureRank.Width = (float)(pictureRank.Width * 0.75);
                    break;
            }

            //if (classcell.isoutstanding())
            //{
            //    oDoc.Bookmarks[ref oBookMark2].Range.Text = "卓越级";
            //    pictureRank = oDoc.InlineShapes.AddPicture(datapath + "3.png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark3].Range);
            //    oDoc.Bookmarks[ref oBookMark4].Range.Text = "卓越级";
            //}
            //else if (classcell.islean())
            //{
            //    oDoc.Bookmarks[ref oBookMark2].Range.Text = "精益级";
            //    pictureRank = oDoc.InlineShapes.AddPicture(datapath + "2.png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark3].Range);
            //    oDoc.Bookmarks[ref oBookMark4].Range.Text = "精益级";
            //}
            //else
            //{
            //    oDoc.Bookmarks[ref oBookMark2].Range.Text = "精细级";
            //    pictureRank = oDoc.InlineShapes.AddPicture(datapath + "1.png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark3].Range);
            //    oDoc.Bookmarks[ref oBookMark4].Range.Text = "精细级";
            //}


            Word.InlineShape pictureRadar;
            object oBookMark5 = "质量要素雷达图";
            pictureRadar = oDoc.InlineShapes.AddPicture(datapath + "radar/1.png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark5].Range);
            pictureRadar.Height = 225;
            pictureRadar.Width = 225;
            oBookMark5 = "效率要素雷达图";
            pictureRadar = oDoc.InlineShapes.AddPicture(datapath + "radar/2.png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark5].Range);
            pictureRadar.Height = 225;
            pictureRadar.Width = 225;
            oBookMark5 = "效益要素雷达图";
            pictureRadar = oDoc.InlineShapes.AddPicture(datapath + "radar/3.png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark5].Range);
            pictureRadar.Height = 225;
            pictureRadar.Width = 225;

            object oBookMark6 = "质量子要素最好";
            object oBookMark7 = "质量子要素最差";
            string tempstr1 = "";
            string tempstr2 = "";
            double tempver1 = 100;
            double tempver2 = 0;
            foreach (var item in classcell.Ret_Dic1())
            {
                if (item.Value < tempver1)
                {
                    tempstr1 = item.Key;
                    tempver1 = item.Value;
                }
                if (item.Value > tempver2)
                {
                    tempstr2 = item.Key;
                    tempver2 = item.Value;
                }
            }
            oDoc.Bookmarks[ref oBookMark7].Range.Text = tempstr1;
            oDoc.Bookmarks[ref oBookMark6].Range.Text = tempstr2;

            oBookMark6 = "效率子要素最好";
            oBookMark7 = "效率子要素最差";
            tempstr1 = "";
            tempstr2 = "";
            tempver1 = 100;
            tempver2 = 0;
            foreach (var item in classcell.Ret_Dic2())
            {
                if (item.Value < tempver1)
                {
                    tempstr1 = item.Key;
                    tempver1 = item.Value;
                }
                if (item.Value > tempver2)
                {
                    tempstr2 = item.Key;
                    tempver2 = item.Value;
                }
            }
            oDoc.Bookmarks[ref oBookMark7].Range.Text = tempstr1;
            oDoc.Bookmarks[ref oBookMark6].Range.Text = tempstr2;

            oBookMark6 = "效益子要素最好";
            oBookMark7 = "效益子要素最差";
            tempstr1 = "";
            tempstr2 = "";
            tempver1 = 100;
            tempver2 = 0;
            foreach (var item in classcell.Ret_Dic3())
            {
                if (item.Value < tempver1)
                {
                    tempstr1 = item.Key;
                    tempver1 = item.Value;
                }
                if (item.Value > tempver2)
                {
                    tempstr2 = item.Key;
                    tempver2 = item.Value;
                }
            }
            oDoc.Bookmarks[ref oBookMark7].Range.Text = tempstr1;
            oDoc.Bookmarks[ref oBookMark6].Range.Text = tempstr2;

            //Word.InlineShape picturePie;
            //oBookMark7 = "饼图1";
            //DrawPieChart2(classcell.scoring()[1] / classcell.Weights[0], 5 - classcell.scoring()[1] / classcell.Weights[0], "质量得分", "未得分", "质量得分情况");
            //picturePie = oDoc.InlineShapes.AddPicture(datapath + classcell.cellname + ".png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark7].Range);
            //picturePie.Height = 175;
            //picturePie.Width = 175;
            //oBookMark7 = "饼图2";
            //DrawPieChart2(classcell.scoring()[2] / classcell.Weights[1], 5 - classcell.scoring()[2] / classcell.Weights[1], "效率得分", "未得分", "效率得分情况");
            //picturePie = oDoc.InlineShapes.AddPicture(datapath + classcell.cellname + ".png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark7].Range);
            //picturePie.Height = 175;
            //picturePie.Width = 175;
            //oBookMark7 = "饼图3";
            //DrawPieChart2(classcell.scoring()[3] / classcell.Weights[2], 5 - classcell.scoring()[3] / classcell.Weights[2], "效益得分", "未得分", "效益得分情况");
            //picturePie = oDoc.InlineShapes.AddPicture(datapath + classcell.cellname + ".png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark7].Range);
            //picturePie.Height = 175;
            //picturePie.Width = 175;
            //oBookMark7 = "饼图4";
            //DrawPieChart2(classcell.scoring()[0], 5 - classcell.scoring()[0], "总得分", "未得分", "总得分情况");
            //picturePie = oDoc.InlineShapes.AddPicture(datapath + classcell.cellname + ".png", ref oMissing, ref oMissing, oDoc.Bookmarks[ref oBookMark7].Range);
            //picturePie.Height = 175;
            //picturePie.Width = 175;


            object oBookMark8 = "质量占比";
            oDoc.Bookmarks[ref oBookMark8].Range.Text = (classcell.scoring()[4]).ToString("f2");
            oBookMark8 = "效率占比";
            oDoc.Bookmarks[ref oBookMark8].Range.Text = (classcell.scoring()[5]).ToString("f2");
            oBookMark8 = "效益占比";
            oDoc.Bookmarks[ref oBookMark8].Range.Text = (classcell.scoring()[6]).ToString("f2");
            oBookMark8 = "最终分数";
            oDoc.Bookmarks[ref oBookMark8].Range.Text = classcell.scoring()[0].ToString("f2");
            oBookMark8 = "最终评级";
            oDoc.Bookmarks[ref oBookMark8].Range.Text = classcell.set_final_weight();
            Dictionary<string, double> tempdic = new Dictionary<string, double>()
            {
                { "质量",(0.2 * classcell.scoring()[1] / classcell.Weights[0])},
                { "效率",(0.2 * classcell.scoring()[2] / classcell.Weights[1])},
                { "效益",(0.2 * classcell.scoring()[3] / classcell.Weights[2])},
            };
            tempstr1 = "";
            tempver1 = 100;
            foreach (var item in tempdic)
            {
                if (item.Value < tempver1)
                {
                    tempstr1 = item.Key;
                    tempver1 = item.Value;
                }
            }
            oBookMark8 = "方面短板";
            oDoc.Bookmarks[ref oBookMark8].Range.Text = tempstr1;
            oBookMark8 = "专家评价";
            oDoc.Bookmarks[ref oBookMark8].Range.Text = classcell.expertcomment;


            object oBookMark9 = "t1";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["产品一次交检合格率"].ToString();
            oBookMark9 = "t2";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["合格率稳定性"].ToString();
            oBookMark9 = "t3";
            Console.WriteLine(classcell.ScoreDic);
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["结果有效性"].ToString();
            oBookMark9 = "t4";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["参数获得率"].ToString();
            oBookMark9 = "t5";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["文件管理"].ToString();
            oBookMark9 = "t6";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["过程风险防控"].ToString();
            oBookMark9 = "t7";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["操作标准化"].ToString();
            oBookMark9 = "t9";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["可检可测程度"].ToString();
            oBookMark9 = "t10";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["人员资质"].ToString();
            oBookMark9 = "t11";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["人为质量问题"].ToString();
            oBookMark9 = "t13";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["数控程序管理"].ToString();
            oBookMark9 = "t15";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["6s管理"].ToString();
            oBookMark9 = "t16";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["产能需求匹配度"].ToString();
            oBookMark9 = "t17";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["准时完成率"].ToString();
            oBookMark9 = "t18";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["作业计划管理"].ToString();
            oBookMark9 = "t19";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["产能提升率"].ToString();
            oBookMark9 = "t20";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["目视化管理"].ToString();
            oBookMark9 = "t22";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["设备故障率"].ToString();
            oBookMark9 = "t21";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["应急能力"].ToString();
            oBookMark9 = "t23";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["成本管控水平"].ToString();
            oBookMark9 = "t24";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["月度成本费用率偏差"].ToString();
            oBookMark9 = "t25";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["单位工时产出偏差"].ToString();
            oBookMark9 = "t26";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["设备利用率"].ToString();
            oBookMark9 = "t27";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["节能降耗"].ToString();
            oBookMark9 = "t28";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["环境减排"].ToString();
            oBookMark9 = "t29";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["运行安全性"].ToString();
            oBookMark9 = "t30";
            oDoc.Bookmarks[ref oBookMark9].Range.Text = classcell.ScoreDic["物流周转水平"].ToString();

            oDoc.SaveAs(datapath + classcell.cellname);
            //oDoc.SaveAs2(datapath + classcell.cellname, WdExportFormat.wdExportFormatPDF);
            oDoc.Close(false);
            Process myProcess = new Process();
            try
            {
                myProcess.StartInfo.FileName = datapath + classcell.cellname + ".doc";
                myProcess.StartInfo.Verb = "Open";
                myProcess.StartInfo.CreateNoWindow = true;
                myProcess.Start();
            }
            catch (Exception)
            {
                myProcess.StartInfo.FileName = datapath + classcell.cellname + ".docx";
                myProcess.StartInfo.Verb = "Open";
                myProcess.StartInfo.CreateNoWindow = true;
                myProcess.Start();
            }
        }
        private void DrawPieChart2(double value1, double value2, string str1, string str2, string title)
        {
            //Chart chart1 = new Chart();
            //reset your chart series and legends
            chart1.Series.Clear();
            chart1.Legends.Clear();

            //Add a new Legend(if needed) and do some formating
            chart1.Legends.Add("MyLegend");
            chart1.Legends[0].LegendStyle = LegendStyle.Table;
            chart1.Legends[0].Docking = Docking.Bottom;
            chart1.Legends[0].Alignment = StringAlignment.Center;
            chart1.Legends[0].Title = title;
            chart1.Legends[0].BorderColor = Color.Black;

            //Add a new chart-series
            string seriesname = "MySeriesName";
            chart1.Series.Add(seriesname);
            //set the chart-type to "Pie"
            chart1.Series[seriesname].ChartType = SeriesChartType.Pie;

            //Add some datapoints so the series. in this case you can pass the values to this method
            chart1.Series[seriesname].Points.AddXY(str1, value1);
            chart1.Series[seriesname].Points.AddXY(str2, value2);
            chart1.SaveImage(datapath + classcell.cellname + ".png", ChartImageFormat.Png);
        }
        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            comment commentform = new comment();
            commentform.Owner = this;
            commentform.text = classcell.expertcomment;
            commentform.ShowDialog();
        }


    }
}
