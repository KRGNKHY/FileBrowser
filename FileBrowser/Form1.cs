﻿using System;
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
                    string[] subItems = new string[] { Path.GetFileName(file), "ファイル" };

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

                    // OSのシステムフォルダはここでは除外する
                    // System.UnauthorizedAccessExceptionの例外となるため
                    if ((dirInfo.Attributes & FileAttributes.System) != FileAttributes.System)
                    {
                        // 展開するノードは以下にサブフォルダのノードを追加
                        // +ボタンを表示させるために仮のノードを追加しておく
                        TreeNode subNode = new TreeNode(dirInfo.Name);
                        subNode.Nodes.Add(new TreeNode());
                        node.Nodes.Add(subNode);
                    }
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

        /*    class DirTreeNode : TreeNode
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
            }*/
    }
}
