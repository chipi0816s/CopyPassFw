/*
    開発に当たり、参考にしたサイト
    https://qiita.com/tinymouse/items/eb8aebb39ddd5c103347 
    https://qiita.com/EvilSpirit39/items/c959fc990d1660b77180 
*/
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CopyPassFw
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.AllFilesAndFolders)]
    public class Class1 : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            // 常に表示する
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            // メニューを生成して項目を追加する
            var menu = new ContextMenuStrip();

            //１つ目を作成する
            var item1 = new ToolStripMenuItem
            {
                Text = "パスのコピー(\" \"無し)"
            };
            item1.Click += (sender, args) => MakePath(false);

            //2つ目を作成する
            var item2 = new ToolStripMenuItem
            {
                Text = "パスのコピー(\" \"有り)"
            };
            item2.Click += (sender, args) => MakePath(true);

            menu.Items.Add(item1);
            menu.Items.Add(item2);


            // メニューを返す
            return menu;
        }
        private void  MakePath(bool kanma)
        {
            try
            {
                //選択したファイルのパスを取得
                List<string> pass = new List<string>();
                foreach (var path in SelectedItemPaths)
                {
                    pass.Add(path);
                }

                //”有り設定の場合、”を入れる処理を行う
                if (kanma)
                {
                    List<string> passAri = new List<string>();
                    foreach (var are in pass)
                    {
                        passAri.Add("\"" + are + "\"");
                    }
                    Clipboard.SetText(PathMoji(passAri));
                    //MessageBox.Show(PathMoji(passAri), "こぴーせいこう(テストウインドウ)");
                    return;
                }

                Clipboard.SetText(PathMoji(pass));
                //MessageBox.Show(PathMoji(pass), "こぴーせいこう(テストウインドウ)");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,"コピー失敗");
            }

            return;
        }

        private string PathMoji(List<string> vs)
        {
            //複数のパスを一つの文字列にまとめる処理
            string result = null;
            foreach(string pass in vs)
            {
                result += pass + "\r\n";
            }
            return result.TrimEnd();
        }
    }
}
