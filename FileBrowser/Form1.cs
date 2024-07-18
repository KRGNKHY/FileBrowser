using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            // 初期フォルダは、MyDocuments
            string initialFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            OpenFolder(initialFolder);
            DirTreeNode driveNode = new DirTreeNode(Path.GetFileName(initialFolder), initialFolder);
            treeView1.Nodes.Add(driveNode);

            // +ボタンを表示させるために仮のノードを追加しておく
            driveNode.Nodes.Add(new DirTreeNode());
        }

        private void OpenFolder(string folder)
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
                string[] subItems = new string[] { Path.GetFileName(file), "ファイル" };

                // this.folderListはListView(デザインの左側のペイン)の名前
                ListViewItem item = new ListViewItem(subItems);
                this.folderList.Items.Add(item);
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
                        OpenFile(path);
                        break;
                    }
            }
        }

        private void OpenFile(string path)
        {
            // ファイルの中身をテキストボックスに表示する
            this.contentsText.Text = File.ReadAllText(path);
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
            DirTreeNode dirNode = (DirTreeNode)e.Node;

            // 展開するノードは以下のノードをクリア
            dirNode.Nodes.Clear();

            string[] dirPaths = Directory.GetDirectories(dirNode.DirectoryPath);

            foreach (string dirPath in dirPaths)
            {
                // フォルダ情報を取得
                DirectoryInfo dirInfo = new DirectoryInfo(dirPath);

                // OSのシステムフォルダはここでは除外する
                // System.UnauthorizedAccessExceptionの例外となるため
                if ((dirInfo.Attributes & FileAttributes.System) != FileAttributes.System)
                {
                    // 展開するノードは以下にサブフォルダのノードを追加
                    DirTreeNode subDirNode = new DirTreeNode(dirInfo.Name, dirPath);
                    dirNode.Nodes.Add(subDirNode);

                    if (Directory.GetDirectories(dirPath).Length > 0)
                    {
                        subDirNode.Nodes.Add(new DirTreeNode());
                    }
                }
            }
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            DirTreeNode dirTreeNode = (DirTreeNode)e.Node;
            DirectoryInfo dirList = new DirectoryInfo(dirTreeNode.FullPath);
            foreach (DirectoryInfo dirInfo in dirList.GetDirectories(dirTreeNode.DirectoryPath)) { 
                ListViewItem item = new ListViewItem(dirInfo.Name);
                OpenItem(item);
            }
        }
    }

    class DirTreeNode : TreeNode
    {
        // 仮のノード用のコンストラクタ
        public DirTreeNode() { }

        // コンストラクタにてフォルダ名、フォルダパスを設定
        public DirTreeNode(string directoryName, string directoryPath) : base(directoryName)
        {
            this.DirectoryPath = directoryPath;
        }

        //フォルダパス
        public string DirectoryPath { get; set; }
    }
}
