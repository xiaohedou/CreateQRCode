using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.OleDb;
using ZXing;
using ZXing.Common;
using ZXing.QrCode.Internal;
using System.Text;

namespace 生成二维码
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        //设置二维码颜色
        Color setGolor()
        {
            Color color;
            switch (comboBox3.SelectedIndex)
            {
                case 0:
                default:
                    color = Color.Red;
                    break;
                case 1:
                    color = Color.FromArgb(51, 128, 47);
                    break;
                case 2:
                    color = Color.Blue;
                    break;
                case 3:
                    color = Color.Purple;
                    break;
                case 4:
                    color = Color.Black;
                    break;
            }
            return color;
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
                //设置生成彩色二维码
                barCodeWriter.Renderer = new ZXing.Rendering.BitmapRenderer { Background = Color.White, Foreground = setGolor() };
                barCodeWriter.Format = BarcodeFormat.QR_CODE;
                barCodeWriter.Options.Hints.Add(EncodeHintType.CHARACTER_SET, "UTF-8");
                barCodeWriter.Options.Hints.Add(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H);
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
                return GetQRCodeByZXingNet(contents, width, height);
            }
            //构造二维码写码器
            MultiFormatWriter mutiWriter = new MultiFormatWriter();
            Dictionary<EncodeHintType, object> hint = new Dictionary<EncodeHintType, object>();
            hint.Add(EncodeHintType.CHARACTER_SET, "UTF-8");
            hint.Add(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H);
            //生成二维码
            BitMatrix bm = mutiWriter.encode(contents, BarcodeFormat.QR_CODE, width, height, hint);
            BarcodeWriter barcodeWriter = new BarcodeWriter();
            //设置生成彩色二维码
            barcodeWriter.Renderer = new ZXing.Rendering.BitmapRenderer { Background = Color.White, Foreground = setGolor() };
            Bitmap bitmap = barcodeWriter.Write(bm);
            //获取二维码实际尺寸（去掉二维码两边空白后的实际尺寸）
            int[] rectangle = bm.getEnclosingRectangle();
            //计算插入图片的大小和位置
            int middleImgW = Math.Min((int)(rectangle[2] / 3.5), middleImg.Width);
            int middleImgH = Math.Min((int)(rectangle[3] / 3.5), middleImg.Height);
            int middleImgL = (bitmap.Width - middleImgW) / 2;
            int middleImgT = (bitmap.Height - middleImgH) / 2;
            //将img转换成bmp格式，否则后面无法创建 Graphics对象
            Bitmap bmpimg = new Bitmap(bitmap.Width, bitmap.Height, PixelFormat.Format32bppArgb);
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

        string path = "";
        //生成二维码
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog floder = new FolderBrowserDialog();
            if (floder.ShowDialog() == DialogResult.OK)
            {
                path = floder.SelectedPath.TrimEnd(new char[] { '\\' }) + "\\";
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                Bitmap qrCode;
                if (tabControl1.SelectedIndex == 0)
                {
                    if (dataGridView1.Rows.Count > 0)
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            string strName = dataGridView1.Rows[i].Cells[Convert.ToInt32(textBox7.Text.Trim()) - 1].Value.ToString();
                            if (checkBox1.Checked)
                                qrCode = GetQRCodeByZXingNet(strName, pictureBox1.Image, Convert.ToInt32(textBox2.Text), Convert.ToInt32(textBox2.Text));
                            else
                                qrCode = GetQRCodeByZXingNet(strName, Convert.ToInt32(textBox2.Text), Convert.ToInt32(textBox2.Text));
                            qrCode.Save(path + strName.Substring(strName.LastIndexOf('/') + 1, strName.LastIndexOf('.') - strName.LastIndexOf('/') - 1) + ".png", ImageFormat.Png);
                        }
                    }
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    for (int i = Convert.ToInt32(textBox4.Text.Trim()); i <= Convert.ToInt32(textBox5.Text.Trim()); i++)
                    {
                        if (checkBox1.Checked)
                            qrCode = GetQRCodeByZXingNet(textBox1.Text.TrimEnd(new char[] { '/' }) + "/" + i + comboBox1.Text, pictureBox1.Image, Convert.ToInt32(textBox2.Text), Convert.ToInt32(textBox2.Text));
                        else
                            qrCode = GetQRCodeByZXingNet(textBox1.Text.TrimEnd(new char[] { '/' }) + "/" + i + comboBox1.Text, Convert.ToInt32(textBox2.Text), Convert.ToInt32(textBox2.Text));
                        qrCode.Save(path + i + ".png", ImageFormat.Png);
                    }
                }
                if (checkBox2.Checked)
                {
                    if (textBox8.Text != string.Empty)
                    {
                        string[] files = Directory.GetFiles(path);
                        for (int i = 0; i < files.Length; i++)
                        {
                            CombinImage(textBox8.Text, files[i], path + Path.GetFileNameWithoutExtension(textBox8.Text) + "_" + Path.GetFileNameWithoutExtension(files[i]) + ".png");
                            File.Delete(files[i]);
                        }
                    }
                }
                MessageBox.Show("生成成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //选择是否包含图片
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                button2.Enabled = true;
            else
                button2.Enabled = false;
        }

        //选择要包含的图片文件
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                pictureBox1.Image = Image.FromFile(openFile.FileName);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = comboBox2.SelectedIndex = comboBox3.SelectedIndex = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", path);
        }

        OleDbConnection olecon;
        public static FileInfo fInfo;
        public static string strPath = "";
        public static string strExtension = "";

        public OleDbConnection getOledb(string strPath, string strExtension)
        {
            if (strExtension == ".xls")
                olecon = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strPath + ";Extended Properties=Excel 8.0");//连接Excel 2003及以下版本
            else
                olecon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strPath + ";Extended Properties=Excel 12.0");//连接Excel 2007及以上版本
            return olecon;
        }

        /// <summary>
        /// 将CSV文件的数据读取到DataTable中
        /// </summary>
        /// <param name="fileName">CSV文件路径</param>
        /// <returns>返回读取了CSV数据的DataTable</returns>
        public static DataTable OpenCSV(string filePath)
        {
            DataTable dt = new DataTable();
            string strpath = filePath; //csv文件的路径
            try
            {
                int intColCount = 0;
                bool blnFlag = true;
                DataColumn mydc;
                DataRow mydr;
                string strline;
                string[] aryline;
                StreamReader mysr = new StreamReader(strpath, Encoding.Default);

                while ((strline = mysr.ReadLine()) != null)
                {
                    aryline = strline.Split(new char[] { ',' });
                    //给datatable加上列名
                    if (blnFlag)
                    {
                        blnFlag = false;
                        intColCount = aryline.Length;
                        for (int i = 0; i < aryline.Length; i++)
                        {
                            mydc = new DataColumn(aryline[i]);
                            dt.Columns.Add(mydc);
                        }
                    }

                    //填充数据并加入到datatable中
                    mydr = dt.NewRow();
                    for (int i = 0; i < intColCount; i++)
                    {
                        mydr[i] = aryline[i];
                    }
                    dt.Rows.Add(mydr);
                }
                dt.Rows.RemoveAt(0);
                return dt;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            OpenFileDialog selectFile = new OpenFileDialog();
            selectFile.Filter = "表格文件(*.csv;*.xls;*.xlsx)|*.csv;*.xls;*.xlsx";
            if (selectFile.ShowDialog() == DialogResult.OK)
            {
                strPath = selectFile.FileName;
                textBox6.Text = strPath;//在文本框中显示Excel文件名
                fInfo = new FileInfo(strPath);
                strExtension = fInfo.Extension;
                if (strExtension == ".xls" || strExtension == ".xlsx")
                {
                    try
                    {
                        olecon = getOledb(strPath, strExtension);
                        //从工作表中查询数据
                        OleDbDataAdapter oledbda = new OleDbDataAdapter("select * from [" + fInfo.Name.Remove(fInfo.Name.LastIndexOf(".")) + "$]", olecon);
                        DataSet myds = new DataSet();//创建数据集对象
                        oledbda.Fill(myds);//填充数据集
                        dataGridView1.DataSource = myds.Tables[0];
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    dataGridView1.DataSource = OpenCSV(strPath);
                }
            }
        }

        //切换二维码生成方式
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                dataGridView1.Visible = true;
                this.Height = 720;
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                dataGridView1.Visible = false;
                this.Height = 390;
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar != 8 && !char.IsDigit(e.KeyChar)) && e.KeyChar != 13)
            {
                MessageBox.Show("此处只能输入数字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);   //弹出信息提示
                e.Handled = true;
            }
        }

        private void 识别二维码ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new FrmValidate().Show();
        }

        private void 使用说明ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("使用说明.txt");
        }

        //是否设置背景图
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
                label12.Visible = textBox8.Visible = button5.Visible = groupBox3.Visible = true;
            else
                label12.Visible = textBox8.Visible = button5.Visible = groupBox3.Visible = false;
        }

        //选择背景图
        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog selectFile = new OpenFileDialog();
            selectFile.Filter = "图片文件(*.png;*.jpg;*.jpeg;*.bmp;*.svg)|*.png;*.jpg;*.jpeg;*.bmp;*.svg";
            if (selectFile.ShowDialog() == DialogResult.OK)
            {
                textBox8.Text = selectFile.FileName;//在文本框中显示图片文件名
            }
        }

        /// <summary>
        /// 合成图片
        /// </summary>
        /// <param name="sourceImg">粘贴的源图片</param>
        /// <param name="destImg">粘贴的目标图片</param>
        public void CombinImage(string sourceImg, string destImg, string newImg)
        {
            Image imgBack = Image.FromFile(sourceImg);//相框图片 
            Image img = Image.FromFile(destImg);//照片图片
            //从指定的Image创建新的Graphics绘图对象
            Graphics g = Graphics.FromImage(imgBack);
            //g.DrawImage(img, 照片与相框的左边距, 照片与相框的上边距, 层叠图片宽, 层叠图片高);
            switch (comboBox2.SelectedIndex)
            {
                case 0:
                default:
                    g.DrawImage(img, imgBack.Width - img.Width, imgBack.Height - img.Height, img.Width, img.Height);
                    break;
                case 1:
                    g.DrawImage(img, 0, imgBack.Height - img.Height, img.Width, img.Height);
                    break;
                case 2:
                    g.DrawImage(img, imgBack.Width - img.Width, 0, img.Width, img.Height);
                    break;
                case 3:
                    g.DrawImage(img, 0, 0, img.Width, img.Height);
                    break;
            }
            GC.Collect();
            img.Dispose();
            imgBack.Save(newImg, ImageFormat.Png);
        }
    }
}
