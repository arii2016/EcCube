using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EcCube
{
    public partial class Form1 : Form
    {
        //--------------------------------------------------------------
        public Form1()
        {
            InitializeComponent();
        }
        //--------------------------------------------------------------
        private void Pl_FileLoad_DragEnter(object sender, DragEventArgs e)
        {
            //コントロール内にドラッグされたとき実行される
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                //ドラッグされたデータ形式を調べ、ファイルのときはコピーとする
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                //ファイル以外は受け付けない
                e.Effect = DragDropEffects.None;
            }
        }
        //--------------------------------------------------------------
        private void Pl_FileLoad_DragDrop(object sender, DragEventArgs e)
        {
            //ドロップされたすべてのファイル名を取得する
            string[] FileNameList = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            for (int i = 0; i < FileNameList.Length; i++)
            {
                Form2 clsForm = new Form2();
                clsForm.ShowDlg(FileNameList[i]);
                clsForm.Dispose();
            }
        }
        //--------------------------------------------------------------
    }
}
