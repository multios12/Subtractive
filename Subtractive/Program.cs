namespace ArchiveXlsx
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Text.RegularExpressions;
    using Subtractive;

    /// <summary>
    /// プログラム
    /// </summary>
    public class Program
    {
        /// <summary>
        /// プログラムのスタートアップポイント
        /// </summary>
        /// <param name="args">コマンドライン引数</param>
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine(Subtractive.Properties.Resources.CopyRight + Environment.NewLine
                    + "ファイルパスを指定してください");

                return;
            }

            string filePath = args[0];
            if (File.Exists(filePath) == false)
            {
                Console.WriteLine(Subtractive.Properties.Resources.CopyRight + Environment.NewLine
                    + "ファイルが見つかりません");
                return;
            }

            if (Path.GetExtension(filePath.ToLower()) == ".xlsx" ||
                Path.GetExtension(filePath.ToLower()) == ".xlsm")
            {
                QuantXlsx(filePath, "xl");
            }
            else if (Path.GetExtension(filePath.ToLower()) == ".pptx")
            {
                QuantXlsx(filePath, "ppt");
            }
            else if (Path.GetExtension(filePath.ToLower()) == ".mdb")
            {
                MdbCompactation.compact(filePath);
            }
            else
            {
                Console.WriteLine(Subtractive.Properties.Resources.CopyRight + Environment.NewLine
                    + "このファイルには対応していません。");
            }
        }

        /// <summary>指定されたエクセルブックファイルに含まれる画像を圧縮します。</summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="ext">拡張子</param>
        private static void QuantXlsx(string filePath, string ext)
        {
            Regex regex = new Regex(string.Format(@"{0}/media/.*\.png", ext));

            Console.WriteLine("減色を開始します：{0}", filePath);

            string uncompressPath = Path.GetTempPath();

            ZipArchive archive = null;
            PngQuant pngquant = new PngQuant();
            string distinationFilePath = Path.GetDirectoryName(filePath);
            distinationFilePath = Path.Combine(distinationFilePath, "減色済" + Path.GetFileName(filePath));
            ZipArchive distinationArchive =
                new ZipArchive(new FileStream(distinationFilePath, FileMode.Create, FileAccess.Write), ZipArchiveMode.Create);
            try
            {
                archive = ZipFile.Open(filePath, ZipArchiveMode.Read);
                IEnumerable<ZipArchiveEntry> entries = archive.Entries;

                foreach (ZipArchiveEntry entry in entries)
                {
                    if (regex.IsMatch(entry.FullName.ToLower()) == false)
                    {
                        using (Stream stream = entry.Open())
                        {
                            ZipArchiveEntry e = distinationArchive.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                            using (Stream s = e.Open())
                            {
                                byte[] bs = stream.ToByteArray((int)entry.Length);
                                s.Write(bs, 0, bs.Length);
                            }
                        }

                        continue;
                    }

                    // 解凍して一時ファイルを作成し、減色する
                    string entryFilePath = Path.GetTempFileName();
                    entryFilePath = pngquant.SubtractiveToTemporaryFile(entry);

                    // 減色したファイルをブックに再設定
                    distinationArchive.CreateEntryFromFile(entryFilePath, entry.FullName, CompressionLevel.Optimal);
                    Console.WriteLine("・{0}", entry.FullName);
                }

                // ブックファイルの保存
                // distinationArchive.(distinationFilePath);
            }
            finally
            {
                distinationArchive.Dispose();

                if (archive != null)
                {
                    archive.Dispose();
                }

                pngquant.Dispose();
            }
        }
    }
}
