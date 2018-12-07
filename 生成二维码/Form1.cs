using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using ZXing;
using ZXing.Common;
using ZXing.QrCode.Internal;

namespace 生成二维码
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 生成二维码图片
        /// </summary>
        /// <param name="strMessage">要生成二维码的字符串</param>
        /// <param name="width">二维码图片宽度（单位：像素）</param>
        /// <param name="height">二维码图片高度（单位：像素）</param>
        /// <returns></returns>
        private Bitmap GetQRCodeByZXingNet(String strMessage, Int32 width, Int32 height)
        {
            Bitmap result = null;
            try
            {
                BarcodeWriter barCodeWriter = new BarcodeWriter();
                barCodeWriter.Format = BarcodeFormat.QR_CODE;
                barCodeWriter.Options.Hints.Add(EncodeHintType.CHARACTER_SET, "UTF-8");
                barCodeWriter.Options.Hints.Add(EncodeHintType.ERROR_CORRECTION, ZXing.QrCode.Internal.ErrorCorrectionLevel.H);
                barCodeWriter.Options.Height = height;
                barCodeWriter.Options.Width = width;
                barCodeWriter.Options.Margin = 0;
                BitMatrix bm = barCodeWriter.Encode(strMessage);
                result = barCodeWriter.Write(bm);
            }
            catch { }
            return result;
        }

        /// <summary>
        /// 生成中间带有图片的二维码图片
        /// </summary>
        /// <param name="contents">要生成二维码包含的信息</param>
        /// <param name="middleImg">要生成到二维码中间的图片</param>
        /// <param name="width">生成的二维码宽度（单位：像素）</param>
        /// <param name="height">生成的二维码高度（单位：像素）</param>
        /// <returns>中间带有图片的二维码</returns>
        public Bitmap GetQRCodeByZXingNet(string contents, Image middleImg, int width, int height)
        {
            if (string.IsNullOrEmpty(contents))
            {
                return null;
            }
            if (middleImg == null)
            {
                return GetQRCodeByZXingNet(contents,width,height);
            }
            //构造二维码写码器
            MultiFormatWriter mutiWriter = new MultiFormatWriter();
            Dictionary<EncodeHintType, object> hint = new Dictionary<EncodeHintType, object>();
            hint.Add(EncodeHintType.CHARACTER_SET, "UTF-8");
            hint.Add(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H);
            //生成二维码
            BitMatrix bm = mutiWriter.encode(contents, BarcodeFormat.QR_CODE, width, height, hint);
            BarcodeWriter barcodeWriter = new BarcodeWriter();
            Bitmap bitmap = barcodeWriter.Write(bm);
            //获取二维码实际尺寸（去掉二维码两边空白后的实际尺寸）
            int[] rectangle = bm.getEnclosingRectangle();
            //计算插入图片的大小和位置
            int middleImgW = Math.Min((int)(rectangle[2] / 3.5), middleImg.Width);
            int middleImgH = Math.Min((int)(rectangle[3] / 3.5), middleImg.Height);
            int middleImgL = (bitmap.Width - middleImgW) / 2;
            int middleImgT = (bitmap.Height - middleImgH) / 2;
            //将img转换成bmp格式，否则后面无法创建 Graphics对象
            Bitmap bmpimg = new Bitmap(bitmap.Width, bitmap.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmpimg))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                g.DrawImage(bitmap, 0, 0);
            }
            //在二维码中插入图片
            Graphics myGraphic = Graphics.FromImage(bmpimg);
            //白底
            myGraphic.FillRectangle(Brushes.White, middleImgL, middleImgT, middleImgW, middleImgH);
            myGraphic.DrawImage(middleImg, middleImgL, middleImgT, middleImgW, middleImgH);
            return bmpimg;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Bitmap qrCode;
            for (int i = Convert.ToInt32(textBox4.Text.Trim()); i <= Convert.ToInt32(textBox5.Text.Trim()); i++)
            {
                if (checkBox1.Checked)
                    qrCode = GetQRCodeByZXingNet(textBox1.Text.TrimEnd(new char[] { '/' }) + "/" + i + comboBox1.Text, pictureBox1.Image, Convert.ToInt32(textBox2.Text), Convert.ToInt32(textBox2.Text));
                else
                    qrCode = GetQRCodeByZXingNet(textBox1.Text.TrimEnd(new char[] { '/' }) + "/" + i + comboBox1.Text, Convert.ToInt32(textBox2.Text), Convert.ToInt32(textBox2.Text));
                qrCode.Save("Image\\"+i + ".eps", ImageFormat.Tiff);
            }
            MessageBox.Show("生成成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                groupBox2.Visible = true;
            else
                groupBox2.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp";
            if(openFile.ShowDialog()==DialogResult.OK)
            {
                pictureBox1.Image = Image.FromFile(openFile.FileName);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", Application.StartupPath + "\\Image\\");
        }
    }
}
