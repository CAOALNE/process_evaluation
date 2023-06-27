using System;
using System.Collections.Generic;
using System.Linq;


namespace proeval
{

    [Serializable]
    public class cell
    {
        public string cellname = "请输入单元名称";
        public string cellindex = "请输入单元编号";
        public string cellCEO = "请输入单元负责人   ";
        public string cellCEOphone = "请输入负责人电话";
        public string expertcomment = "暂无";
        public double[] Weights;
        public Dictionary<string, int> WeightDic = new Dictionary<string, int>()
        {
            {"产品一次交检合格率",16},
            {"合格率稳定性",8},
            {"参数获得率",8},
            {"结果有效性",4},
            {"试验（检测）成功率",16},
            {"严重质量问题",0},
            {"文件管理",8},
            {"过程风险防控",4},
            {"操作标准化",8},
            {"可检可测程度",4},
            {"产品智能检测水平",0},
            {"人员资质",2},
            {"人为质量问题",4},
            {"装备自动化水平",0},
            {"数控程序管理",8},
            {"设备运行状态在线监控",0},
            {"6s管理",4},
            {"产能需求匹配度",8},
            {"准时完成率",8},
            {"作业计划管理",8},
            {"产能提升率",2},
            {"柔性化程度",0},
            {"设备故障率",4},
            {"目视化管理",4},
            {"应急能力",4},
            {"设备布局合理性",0},
            {"物流周转水平",4},
            {"成本管控水平",4},
            {"月度成本费用率偏差",8},
            {"单位工时产出偏差",4},
            {"设备利用率",4},
            {"节能降耗",2},
            {"环境减排",1},
            {"运行安全性",4},
            {"重大安全事故",0},
        };
        public Dictionary<string, double> ScoreDic = new Dictionary<string, double>()
        {
            {"产品一次交检合格率",0},
            {"合格率稳定性",0},
            {"参数获得率",0},
            {"结果有效性",0},
            {"试验（检测）成功率",0},
            {"严重质量问题",0},
            {"文件管理",0},
            {"过程风险防控",0},
            {"操作标准化",0},
            {"可检可测程度",0},
            {"产品智能检测水平",0},
            {"人员资质",0},
            {"人为质量问题",0},
            {"装备自动化水平",0},
            {"数控程序管理",0},
            {"设备运行状态在线监控",0},
            {"6s管理",0},
            {"产能需求匹配度",0},
            {"准时完成率",0},
            {"作业计划管理",0},
            {"产能提升率",0},
            {"柔性化程度",0},
            {"设备故障率",0},
            {"目视化管理",0},
            {"应急能力",0},
            {"设备布局合理性",0},
            {"物流周转水平",0},
            {"成本管控水平",0},
            {"月度成本费用率偏差",0},
            {"单位工时产出偏差",0},
            {"设备利用率",0},
            {"节能降耗",0},
            {"环境减排",0},
            {"运行安全性",0},
            {"重大安全事故",0},
        };
        public List<string> deleteditems = new List<string>();
        public bool isqualified()
        {
            if (ScoreDic["严重质量问题"] > 0 && ScoreDic["重大安全事故"] > 0)
                return true;
            else
                return false;
        }
        public bool islean()
        {
            if (isqualified() && scoring()[0] >= 4 && ScoreDic["装备自动化水平"] >= 4 && ScoreDic["设备运行状态在线监控"] >= 4 && ScoreDic["柔性化程度"] >= 4 && ScoreDic["目视化管理"] >= 4 && ScoreDic["设备布局合理性"] >= 4)
                return true;
            else
                return false;
        }

        public bool isoutstanding()
        {
            if (islean() && ScoreDic["产品智能检测水平"] >= 5 && scoring()[0] >= 4.5)
                return true;
            else
                return false;
        }
        public void setval(string a, double b)//设置a的分值为b
        {
            if (ScoreDic.ContainsKey(a) && b >= -1 && b <= 5)
                ScoreDic[a] = b;
        }
        public double Getval(string a)
        {
            if (ScoreDic.ContainsKey(a))
                return ScoreDic[a];
            else
                return -1;
        }
        public int Getweight(string a)
        {
            if (WeightDic.ContainsKey(a))
                return WeightDic[a];
            else
                return -1;
        }


        public int ini_weight()
        {
            Weights = new double[3] { 0.4, 0.3, 0.3 };
            if (ScoreDic["严重质量问题"] > 0 && ScoreDic["重大安全事故"] > 0 && ScoreDic["装备自动化水平"] >= 5 && ScoreDic["设备运行状态在线监控"] >= 5 && ScoreDic["柔性化程度"] >= 5 && ScoreDic["目视化管理"] >= 5 && ScoreDic["设备布局合理性"] >= 5 && ScoreDic["产品智能检测水平"] >= 5)
            {
                Weights[0] = 0.3;
                Weights[1] = 0.3;
                Weights[2] = 0.4;
                return 3;
            }
            else if (ScoreDic["严重质量问题"] > 0 && ScoreDic["重大安全事故"] > 0 && ScoreDic["装备自动化水平"] >= 4 && ScoreDic["设备运行状态在线监控"] >= 4 && ScoreDic["柔性化程度"] >= 4 && ScoreDic["目视化管理"] >= 4 && ScoreDic["设备布局合理性"] >= 4)
            {
                Weights[0] = 0.3;
                Weights[1] = 0.4;
                Weights[2] = 0.3;
                return 2;
            }
            else if (ScoreDic["严重质量问题"] > 0 && ScoreDic["重大安全事故"] > 0)
            {
                Weights[0] = 0.4;
                Weights[1] = 0.3;
                Weights[2] = 0.3;
                return 1;
            }
            return 0;
        }
        public double[] scoring()
        {
            double[] scores = new double[7];
            if (Weights == null)
                ini_weight();
            double TotalScore1 = 0;
            double TotalScore2 = 0;
            double TotalScore3 = 0;
            int SumWeight = 0;
            for (int i = 0; i < 17; i++)
            {
                if (deleteditems.Contains(ScoreDic.ElementAt(i).Key))
                    continue;
                TotalScore1 += ScoreDic.ElementAt(i).Value * WeightDic.ElementAt(i).Value;
                SumWeight += WeightDic.ElementAt(i).Value;
            }
            TotalScore1 /= SumWeight;
            SumWeight = 0;
            for (int i = 17; i < 28; i++)
            {
                if (deleteditems.Contains(ScoreDic.ElementAt(i).Key))
                    continue;
                TotalScore2 += ScoreDic.ElementAt(i).Value * WeightDic.ElementAt(i).Value;
                SumWeight += WeightDic.ElementAt(i).Value;
            }
            TotalScore2 /= SumWeight;
            SumWeight = 0;
            for (int i = 28; i < ScoreDic.Count(); i++)
            {
                if (deleteditems.Contains(ScoreDic.ElementAt(i).Key))
                    continue;
                TotalScore3 += ScoreDic.ElementAt(i).Value * WeightDic.ElementAt(i).Value;
                SumWeight += WeightDic.ElementAt(i).Value;
            }
            TotalScore3 /= SumWeight;
            scores[0] = TotalScore1 * Weights[0] + TotalScore2 * Weights[1] + TotalScore3 * Weights[2];
            scores[1] = TotalScore1 * Weights[0];
            scores[2] = TotalScore2 * Weights[1];
            scores[3] = TotalScore3 * Weights[2];
            scores[4] = TotalScore1;
            scores[5] = TotalScore2;
            scores[6] = TotalScore3;

            //0为总得分（5分制）
            return scores;
        }

        public string set_final_weight()
        {
            string ret = "";
            int W = ini_weight();
            if (W == 3)
            {
                if (scoring()[0] > 4.5)
                    ret = "卓越级";
                else
                {
                    Weights[0] = 0.3;
                    Weights[1] = 0.4;
                    Weights[2] = 0.3;
                    W = 2;//转为精益级权重
                }
            }
            if (W == 2)
            {
                if (scoring()[0] > 4)
                    ret = "精益级";
                else
                {
                    Weights[0] = 0.4;
                    Weights[1] = 0.3;
                    Weights[2] = 0.3;
                    W = 1;//转为精细级权重
                }
            }
            if (W == 1)
            {
                if (scoring()[0] > 3)
                    ret = "精益级";
                else
                {
                    ret = "无评级";
                    W = 0;//转为精细级权重
                }
            }
            return ret;
        }

        public string scorecheck()
        {
            string ret = "";
            foreach (var key in ScoreDic)
            {
                if (key.Value == 0)
                {
                    ret += key + " ";
                }
            }
            if (ret != "")
                ret = "以下子要素可能还没有打分" + ret + " 若子要素已剔除请忽略。";

            if ((deleteditems.Contains("产品一次交检合格率") && deleteditems.Contains("试验（检测）成功率")))
                ret += "请注意产品一次交检合格率与试验（检测）成功率是否根据单元类型进行了裁剪";
            if (ret == "")
                return "ok";
            else
                return ret;
        }

        public bool IsLeanThreshold(string a)
        {
            foreach (var item in WeightDic)
            {
                if (a == item.Key && item.Value == 0)
                    return true;
            }
            return false;
        }

        float Cal_Subelement(string str)
        {
            string[] arr = str.Split(',');
            float sum = 0;
            float ret = 0;
            foreach (string s in arr)
            {
                if (deleteditems.Contains(s))
                    continue;
                ret +=(float) Getval(s) * Getweight(s);
                sum += 5 * Getweight(s);
            }
            if (sum == 0)
                return 0;
            ret /= sum;
            return ret;
        }
        public Dictionary<string, float> Ret_Dic1()
        {
            Dictionary<string, float> Quality = new Dictionary<string, float>()
            {
                { "产品质量", 100*Cal_Subelement("产品一次交检合格率,合格率稳定性,参数获得率,结果有效性,试验（检测）成功率,严重质量问题") },
                { "过程控制", 100*Cal_Subelement("文件管理,过程风险防控,操作标准化") },
                { "检测能力", 100*Cal_Subelement("可检可测程度,产品智能检测水平") },
                { "人员素质", 100*Cal_Subelement("人员资质,人为质量问题") },
                { "设备能力", 100*Cal_Subelement("装备自动化水平,数控程序管理,设备运行状态在线监控") },
                { "现场环境", 100*Cal_Subelement("6s管理") },
            };
            return Quality;
        }
        public Dictionary<string, float> Ret_Dic2()
        {
            Dictionary<string, float> Efficency = new Dictionary<string, float>()
            {
                { "产能", 100*Cal_Subelement("产能需求匹配度,准时完成率,作业计划管理,产能提升率") },
                { "响应速度", 100*Cal_Subelement("柔性化程度,设备故障率,目视化管理,应急能力") },
                { "物流效率",100*Cal_Subelement("设备布局合理性,物流周转水平") },
            };
            return Efficency;
        }
        public Dictionary<string, float> Ret_Dic3()
        {
            Dictionary<string, float> Benefit = new Dictionary<string, float>()
            {
                { "成本",100*Cal_Subelement("成本管控水平,月度成本费用率偏差,单位工时产出偏差,设备利用率") },
                { "绿色制造",100*Cal_Subelement("节能降耗,环境减排") },
                { "安全", 100*Cal_Subelement("运行安全性,重大安全事故") },
            };
            return Benefit;
        }
    }
}
