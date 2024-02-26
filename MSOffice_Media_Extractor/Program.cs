using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;

namespace MSOffice_Media_Extractor
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = String.Empty;
            string extract = String.Empty;
            try
            {
                filePath = args[0];
                extract = Path.GetDirectoryName(filePath) + @"\" + Path.GetFileNameWithoutExtension(filePath);
                if (Directory.Exists(Path.GetDirectoryName(filePath) + $@"\{Path.GetFileNameWithoutExtension(filePath)}"))
                {
                    Directory.Delete(Path.GetDirectoryName(filePath) + $@"\{Path.GetFileNameWithoutExtension(filePath)}",true);

                }
            }
            catch (System.IndexOutOfRangeException ex)
            {
                MessageBox.Show("このアプリはGUIがありません\n直接アプリにファイルをドラッグアンドドロップしてください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                ZipFile.ExtractToDirectory(filePath, extract);
                File.Delete(extract + @"\[Content_Types].xml");
                string[] Folders = Directory.GetDirectories(extract);
                foreach (string folder in Folders)
                {
                    Debug.WriteLine(folder);
                    Folders = Directory.GetDirectories(extract);
                    if (folder.Contains("_rels") || folder.Contains("docProps"))
                    {
                        Debug.WriteLine("!!!!!!!!!!!!!!!!!!" + folder);
                    }
                    else
                    {
                        Directory.Move(folder + @"\media", extract + @"\media");
                    }
                    Directory.Delete(folder, true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラーが発生しました\nmediaフォルダがない可能性があります", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Directory.Delete(extract, true);
                return;
            }
            MessageBox.Show($"展開が完了しました！\nFILEPATH:{extract + @"\media"}","完了",MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
