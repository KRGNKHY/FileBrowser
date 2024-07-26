using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;


namespace FileBrowser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // 初期フォルダは、Cドライブ
            string[] drives = Environment.GetLogicalDrives();

            foreach (string drive in drives)
            {
                // +ボタンを表示させるために仮のノードを追加しておく
                TreeNode node = new TreeNode(drive);
                node.Nodes.Add(new TreeNode());
                this.treeView1.Nodes.Add(node);
            }
            // Cドライブのノードのみ開く
            this.treeView1.Nodes[0].Expand();

            // ListViewにドライブの情報を表示
            OpenFolder(drives.First());
        }

        private void OpenFolder(string folder)
        {
            try
            {
                if (!Directory.Exists(folder))
                {
                    return;
                }

                // 現在のフォルダパスを、テキストボックスに表示
                this.addressText.Text = folder;

                // 次のディレクトリに移動する際に、前のディレクトリの情報を破棄する
                this.folderList.Items.Clear();
                this.folderList.Items.Add("上のフォルダへ");

                // 初期フォルダのサブフォルダを一覧表示
                string[] dirs = Directory.GetDirectories(folder);
                foreach (string dir in dirs)
                {
                    // Path.GetFileName()：フルパスからファイル名のみを取得
                    string[] subItems = new string[] { Path.GetFileName(dir), "フォルダ" };

                    // this.folderListはListView(デザインの左側のペイン)の名前
                    ListViewItem item = new ListViewItem(subItems);
                    this.folderList.Items.Add(item);
                }

                // 初期フォルダのファイルの一覧表示
                string[] files = Directory.GetFiles(folder);
                foreach (string file in files)
                {
                    // ファイル名を取得
                    string fileName = Path.GetFileName(file);

                    //拡張子を取得
                    string extension = Path.GetExtension(file);

                    string[] subItems = new string[] { fileName, "ファイル", extension };

                    // this.folderListはListView(デザインの左側のペイン)の名前
                    ListViewItem item = new ListViewItem(subItems);
                    this.folderList.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                string parentDir = Path.GetDirectoryName(this.addressText.Text);
                OpenFolder(parentDir);
                MessageBox.Show(ex.Message + "\nこのフォルダまたはファイルは開くことができません。", "選択エラー");
            }
        }

        private void folderList_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // マウスがどのアイテムをクリックしたか判定
            // e.Locationはマウスがダブルクリックした位置
            ListViewHitTestInfo hti = this.folderList.HitTest(e.Location);
            ListViewItem item = hti.Item;

            // 該当するアイテムを開く
            OpenItem(item);
        }

        private void OpenItem(ListViewItem item)
        {
            // 該当するアイテムを開く
            // 相対パスでの記述だとDirectoryNotFoundExceptionの例外でエラーになるので絶対パス指定にする
            // 現在のフォルダパス+目的のフォルダパス名＝目的のフォルダの絶対パスになる
            string path = Path.Combine(this.addressText.Text, item.Text);
            if (item.Text == "上のフォルダへ")
            {
                string parentDir = Path.GetDirectoryName(this.addressText.Text);
                OpenFolder(parentDir);
                return;
            }
            switch (item.SubItems[1].Text)
            {
                case "フォルダ":
                    {
                        OpenFolder(path);
                        break;
                    }
                case "ファイル":
                    {
                        // 開いたファイルが画像である場合はOpenImage、そうでない場合はOpenFile
                        if (IsValidImage(path))
                        {
                            OpenImage(path);
                        }
                        else
                        {
                            OpenFile(path);
                        }
                        break;
                    }
            }
        }

        private void OpenFile(string path)
        {
            try
            {
                if (!File.Exists(path))
                {
                    return;
                }

                // 拡張子
                string extension = Path.GetExtension(path);

                if (extension == ".txt")
                {
                    // ファイルの中身をテキストボックスに表示する
                    this.contentsText.Text = File.ReadAllText(path);
                }
                else if (extension.Contains(".xls"))
                {
                    // Excelファイルを開く
                    OpenBook(path);
                }
                else if (extension.Contains(".doc"))
                {
                    Microsoft.Office.Interop.Word.Application docApp = new Microsoft.Office.Interop.Word.Application();
                    Documents docs = docApp.Documents;
                    Document doc = docs.Open(path);
                    docApp.Visible = true;
                    while (docApp.Visible == true)
                    {
                    }
                }
                else if (extension.Contains(".ppt"))
                {
                    Microsoft.Office.Interop.PowerPoint.Application ppApp = new Microsoft.Office.Interop.PowerPoint.Application();
                    Presentations presentations = ppApp.Presentations;
                    Presentation presentation = presentations.Open(path);
                    presentation.Save();
                    presentation.Close();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "選択エラー");
            }
        }


        private static void OpenBook(string path)
        {
            var xlApp = new Microsoft.Office.Interop.Excel.Application();

            //Workbooksオブジェクトの作成
            Workbooks xlBooks = xlApp.Workbooks;

            //エクセルファイルの指定
            Workbook xlBook = xlBooks.Open(path);

            //操作中のエクセルを表示したいときにTrueにする
            xlApp.Visible = true;

            while (xlApp.Visible == true)
            {
            }

            if (xlBook != null)
            {
                Marshal.ReleaseComObject(xlBook);
            }
            if (xlBooks != null)
            {
                Marshal.ReleaseComObject(xlBooks);
            }

            if (xlApp != null)
            {
                //エクセルを閉じる
                xlApp.Quit();
            }
            Marshal.ReleaseComObject(xlApp);

        }

        private static void OpenImage(string path)
        {
            // Form2のコンストラクタでパスに相当する画像イメージを取得
            Form2 form2 = new Form2(path);

            // 取得した画像イメージをフォーム2に表示する
            form2.ShowDialog();
        }

        // 画像形式ファイルかどうかを判断するメソッド
        private static bool IsValidImage(string path)
        {
            List<ImageFormat> imageFormats = new List<ImageFormat>()
            {
                ImageFormat.Bmp,
                ImageFormat.Gif,
                ImageFormat.Jpeg,
                ImageFormat.Png,
            };
            try
            {
                if (!File.Exists(path))
                {
                    return false;
                }
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                using (Image image = Image.FromStream(fs))
                {
                    return true;
                }
            }
            catch (Exception)
            {

                return false;
            }
        }

        private void folderList_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 何も選択されていないと例外が発生するので、ガードをかけておく
            if (this.folderList.SelectedItems.Count == 0)
            {
                return;
            }

            // e.KeyChar：押されたキーが格納される。今回はEnterを押した場合にOpenItemメソッドが実行される
            if (e.KeyChar == (char)Keys.Enter)
            {
                // 選択中のアイテムを開く
                OpenItem(this.folderList.SelectedItems[0]);
            }
        }

        private void treeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            // ノードの取得
            TreeNode node = e.Node;

            // パスの取得
            string path = node.FullPath;

            // 展開するノードは以下のノードをクリア
            node.Nodes.Clear();
            try
            {
                string[] dirPaths = Directory.GetDirectories(path);

                foreach (string dirPath in dirPaths)
                {
                    // フォルダ情報を取得
                    DirectoryInfo dirInfo = new DirectoryInfo(dirPath);

                    // 展開するノードは以下にサブフォルダのノードを追加
                    // +ボタンを表示させるために仮のノードを追加しておく
                    TreeNode subNode = new TreeNode(dirInfo.Name);
                    subNode.Nodes.Add(new TreeNode());
                    node.Nodes.Add(subNode);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\nこのフォルダまたはファイルは開くことができません。", "選択エラー");
            }

        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            OpenFolder(e.Node.FullPath);
        }

        private void addressText_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                string path = this.addressText.Text;
                {
                    if (!string.IsNullOrEmpty(path))
                    {
                        OpenFolder(path);
                    }
                }
            }
        }
    }
}
