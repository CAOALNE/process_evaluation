using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;

//把这个类粘贴到你的项目中，执行RadarDemo.Show(); 就会在你的根目录里生成雷达图了，为了方便理解怎么画出来的，我把画每一个步骤时的图片都保存下来了。可以自行运行查看
public class RadarDemo
{
    float mW = 1200;
    float mH = 1200;
    public Dictionary<string, float> mData = new Dictionary<string, float>
    {
    };//维度数据
    float mCount = 5; //边数
    float mCenter = 1200 * 0.5f; //中心点
    float mRadius = 1200 * 0.5f - 100; //半径(减去的值用于给绘制的文本留空间)
    double mAngle = (Math.PI * 2) / 5; //角度
    Graphics graphics = null;
    int mPointRadius = 5; // 各个维度分值圆点的半径  
    int textFontSize = 32;  //顶点文字大小 px
    const string textFontFamily = "Microsoft Yahei"; //顶点字体
    Color lineColor = Color.Green;
    Color fillColor = Color.FromArgb(128, 255, 0, 0);
    Color fontColor = Color.Black;
    string path = "C:/Users/ALEX/Desktop/proeval/bin/x64/Release/";
    public void Show(string s, string name)//路径，图片名字
    {
        if (s != "")
            path = s;
        mCount = mData.Count; //边数
        mCenter = mW * 0.5f; //中心点
        mRadius = mCenter - 100; //半径(减去的值用于给绘制的文本留空间)
        mAngle = (Math.PI * 2) / mCount; //角度
        graphics = null;
        lineColor = Color.Green;
        fillColor = Color.FromArgb(128, 255, 0, 0);
        fontColor = Color.Black;
        Bitmap img = new Bitmap((int)mW, (int)mH);
        graphics = Graphics.FromImage(img);
        graphics.Clear(Color.White);
        DrawPolygon(graphics);
        DrawLines(graphics);
        DrawText(graphics);
        DrawRegion(graphics);
        DrawCircle(graphics);
        img.Save(path + "radar/" + name + ".png", ImageFormat.Png);
        img.Dispose();
        graphics.Dispose();
    }
    // 绘制多边形边
    private void DrawPolygon(Graphics ctx)
    {
        var r = mRadius / mCount; //单位半径
        Pen pen = new Pen(lineColor);
        //画6个圈
        for (var i = 0; i < mCount; i++)
        {
            var points = new List<PointF>();
            var currR = r * (i + 1); //当前半径
                                     //画6条边
            for (var j = 0; j < mCount; j++)
            {
                var x = (float)(mCenter + currR * Math.Cos(mAngle * j));
                var y = (float)(mCenter + currR * Math.Sin(mAngle * j));
                points.Add(new PointF { X = x, Y = y });
            }
            ctx.DrawPolygon(pen, points.ToArray());
            //break;
        }
        ctx.Save();
    }
    //顶点连线
    private void DrawLines(Graphics ctx)
    {
        for (var i = 0; i < mCount; i++)
        {
            var x = (float)(mCenter + mRadius * Math.Cos(mAngle * i));
            var y = (float)(mCenter + mRadius * Math.Sin(mAngle * i));
            ctx.DrawLine(new Pen(lineColor), new PointF { X = mCenter, Y = mCenter }, new PointF { X = x, Y = y });
            //break;
        }
        ctx.Save();
    }
    //绘制文本
    private void DrawText(Graphics ctx)
    {
        var fontSize = textFontSize;//mCenter / 12;
        Font font = new Font(textFontFamily, fontSize, FontStyle.Bold);
        int i = 0;
        foreach (var item in mData)
        {
            var x = (float)(mCenter + mRadius * Math.Cos(mAngle * i) * 0.8);
            var y = (float)(mCenter + mRadius * Math.Sin(mAngle * i) * 0.8 - fontSize);
            ctx.DrawString(item.Key + "\r\n分数：" + (item.Value / 20).ToString("f2"), font, new SolidBrush(fontColor), x - 100, y);
            i++;
        }
        ctx.Save();
    }
    //绘制数据区域
    private void DrawRegion(Graphics ctx)
    {
        int i = 0;
        List<PointF> points = new List<PointF>();
        foreach (var item in mData)
        {
            var x = (float)(mCenter + mRadius * Math.Cos(mAngle * i) * item.Value / 100);
            var y = (float)(mCenter + mRadius * Math.Sin(mAngle * i) * item.Value / 100);
            points.Add(new PointF { X = x, Y = y });
            //ctx.DrawArc(new Pen(lineColor), x, y, r, r, 0, (float)Math.PI * 2); 
            i++;
        }
        //GraphicsPath path = new GraphicsPath();
        //path.AddLines(points.ToArray());
        ctx.FillPolygon(new SolidBrush(fillColor), points.ToArray());
        ctx.Save();
    }
    //画点
    private void DrawCircle(Graphics ctx)
    {
        //var r = mCenter / 18;
        var r = mPointRadius;
        int i = 0;
        foreach (var item in mData)
        {
            var x = (float)(mCenter + mRadius * Math.Cos(mAngle * i) * item.Value / 100);
            var y = (float)(mCenter + mRadius * Math.Sin(mAngle * i) * item.Value / 100);
            ctx.FillPie(new SolidBrush(fillColor), x - r, y - r, r * 2, r * 2, 0, 360);
            //ctx.DrawArc(new Pen(lineColor), x, y, r, r, 0, (float)Math.PI * 2); 
            i++;
        }
        ctx.Save();
    }
}