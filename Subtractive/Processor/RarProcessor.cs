﻿namespace Subtractive.Processor
{
    using System;
    using System.Drawing;
    using System.IO;
    using System.IO.Compression;
    using System.Text.RegularExpressions;
    using SharpCompress.Archives.Rar;

    /// <summary>
    /// Rarファイル内のPNGファイルを減色して圧縮します
    /// </summary>
    public class RarProcessor : IProcessor
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public RarProcessor()
        {
            if (Properties.Settings.Default.ResizeMinWidth > 0 &&
                Properties.Settings.Default.ResizeMinHeight > 0)
            {
                this.ResizeSize = new Size(Properties.Settings.Default.ResizeMinWidth, Properties.Settings.Default.ResizeMinHeight);
            }
        }

        /// <summary>ファイル名変更完了イベント</summary>
        public event EventHandler<FileNameChangedEventArgs> FileNameChanged;

        /// <summary>拡張子</summary>
        public virtual string[] Extensions => new string[] { ".zip" };

        /// <summary>BMP/JPEGをPNGにコンバートする場合、True</summary>
        public virtual bool IsConvertToPng { get; set; } = true;

        /// <summary>出力したファイルのパス</summary>
        public virtual string OutputFilePath { get; set; }

        /// <summary>リサイズ時の最大サイズ</summary>
        public Size? ResizeSize { get; set; } = null;

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        public virtual void Execute(string filePath)
        {
            this.QuantFromZipFile(filePath);
        }

        /// <summary>指定されたエクセルブックファイルに含まれる画像を圧縮します。</summary>
        /// <param name="sourceFilePath">ファイルパス</param>
        /// <param name="targetInnerPath">処理対象内部パス</param>
        internal void QuantFromZipFile(string sourceFilePath, string targetInnerPath = "")
        {
            string reg = this.IsConvertToPng ? @"{0}.*\.png|{0}.*\.jpeg|{0}.*\.jpg|{0}.*\.bmp" : @"{0}.*\.png";
            reg = string.Format(reg, targetInnerPath);
            Regex regex = new Regex(reg);

            Console.WriteLine("減色を開始します：{0}", sourceFilePath);

            string distinationFileName = Path.GetFileNameWithoutExtension(sourceFilePath) + ".zip";
            string.Format(Properties.Settings.Default.FileNameTemplate, distinationFileName);
            string distinationFilePath = Path.GetDirectoryName(sourceFilePath);
            distinationFilePath = Path.Combine(distinationFilePath, distinationFileName);
            this.OutputFilePath = distinationFilePath;

            using (RarArchive archive = RarArchive.Open(sourceFilePath))
            using (ZipArchive distArchive = new ZipArchive(new FileStream(distinationFilePath, FileMode.Create, FileAccess.Write), ZipArchiveMode.Create))
            using (PngQuant pngquant = new PngQuant())
            {
                foreach (RarArchiveEntry entry in archive.Entries)
                {
                    if (regex.IsMatch(entry.Key.ToLower()) == false)
                    {
                        ZipArchiveEntry distEntry = distArchive.CreateEntry(entry.Key);
                        using (Stream srcStream = entry.OpenEntryStream())
                        using (Stream distStream = distEntry.Open())
                        {
                            srcStream.CopyTo(distStream);
                        }

                        continue;
                    }

                    // 解凍して一時ファイルを作成し、減色する
                    Stream tranStream;
                    using (Stream s = entry.OpenEntryStream())
                    {
                        tranStream = pngquant.Subtractive(s);
                    }

                    // 減色したファイルをブックに再設定
                    string distEntryName = PathUtils.ChangeExtension(entry.Key, ".png");
                    var e = distArchive.CreateEntry(distEntryName, CompressionLevel.NoCompression);
                    using (var stream = e.Open())
                    {
                        tranStream.CopyTo(stream);
                    }

                    Console.WriteLine("・{0}", entry.Key);

                    // 変換イベントを発生させる
                    this.FileNameChanged?.Invoke(this, new FileNameChangedEventArgs(entry.Key, distEntryName));
                }
            }
        }
    }
}
