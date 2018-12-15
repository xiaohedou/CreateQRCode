using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ZXing;

namespace 生成二维码
{
    public partial class FrmValidate : Form
    {
        public FrmValidate()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 读取二维码
        /// 读取失败，返回空字符串
        /// </summary>
        /// <param name="filename">指定二维码图片位置</param>
        static string Read(string filename)
        {
            BarcodeReader reader = new BarcodeReader();
            reader.Options.CharacterSet = "UTF-8";
            Bitmap map = new Bitmap(filename);
            Result result = reader.Decode(map);
            map.Dispose();//释放图片文件占用资源
            return result == null ? "" : result.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string[] strFiles = openFileDialog1.FileNames;
                for (int i = 0; i < strFiles.Count(); i++)
                {
                    listView1.Items.Add(strFiles[i]);//显示文件列表
                }
            }
        }

        //识别二维码
        private void button1_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count > 0)//判断是否存在要对应的图片列表项
            {
                listView2.Items.Clear();//清空对比列表
                string filename, result, flag = "";//要比较的图片文件名\图片对应二维码地址\是否重复
                for (int i = 0; i < listView1.Items.Count; i++)//遍历所有要比较的图片列表
                {
                    filename = listView1.Items[i].Text;//记录当前要比较的图片文件名
                    result = Read(filename);//获取图片对应二维码
                    if (result != "")//如果存在对应二维码地址，则添加到对比列表中
                    {
                        //使用二维码编号生成列表子项
                        ListViewItem lvItem = new ListViewItem(new string[] { filename.Substring(filename.LastIndexOf('\\') + 1), result.Substring(result.LastIndexOf('\\') + 1), flag });
                        listView2.Items.Add(lvItem);//将列表子项添加到对比列表中
                    }
                }
            }
            else
            {
                MessageBox.Show("请确认存在要识别的二维码图片列表！", "温馨提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
