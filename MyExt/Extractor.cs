//depends https://sharpshell.codeplex.com/
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.IO;

namespace MyExt
{
   [ComVisible(true)]
   [COMServerAssociation(AssociationType.Directory)]
    public class Extractor : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            return true;
        }

        protected override System.Windows.Forms.ContextMenuStrip CreateMenu()
        {
            var menu = new ContextMenuStrip();
            var mymenuitem = new ToolStripMenuItem {Text = "Extract"};
            mymenuitem.Click +=mymenuitem_Click;
            menu.Items.Add(mymenuitem);
            return menu;
        }

        public static bool IsDirectoryEmpty(DirectoryInfo directory)
        {
            FileInfo[] files = directory.GetFiles();
            DirectoryInfo[] subdirs = directory.GetDirectories();

            return (files.Length == 0 && subdirs.Length == 0);
        }
        public static bool IsDirectoryEmpty(string path)
        {
            return IsDirectoryEmpty(new DirectoryInfo(path));
        }
        private void mymenuitem_Click(object sender, System.EventArgs e)
        {
            foreach (var VARIABLE in SelectedItemPaths)
            {
                FileAttributes attr = File.GetAttributes(VARIABLE);
                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    var tmppath = Path.GetDirectoryName(VARIABLE);
                    var files = Directory.GetFiles(VARIABLE);
                    foreach (var file in files)
                    {
                        File.Move(file, tmppath + "\\" + Path.GetFileName(file));
                        if (IsDirectoryEmpty(VARIABLE))
                        {
                            Directory.Delete(VARIABLE);
                        }
                    }
                }
            }
        }
    }
}
