namespace Subtractive
{
    using System;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.IO.Compression;
    using ImageProcessor;
    using ImageProcessor.Imaging;

    /// <summary>PNG減色処理クラス</summary>
    public class PngQuant : IDisposable
    {
        /// <summary>一時フォルダパス</summary>
        private string temporaryFolderPath;

        /// <summary>イメージファクトリ</summary>
        private ImageFactory imageFactory = new ImageFactory();

        /// <summary>リサイズ時の最大サイズ</summary>
        private Size? size = null;

        /// <summary>コンストラクタ</summary>
        /// <param name="size">サイズ</param>
        public PngQuant(Size? size = null)
        {
            this.size = size;
            this.temporaryFolderPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(this.temporaryFolderPath);
        }

        /// <summary>一時フォルダパス</summary>
        public string TemporaryFolderPath
        {
            get
            {
                return this.temporaryFolderPath;
            }
        }

        /// <summary>disposeが完了した場合、true</summary>
        private bool _disposed { get; set; }

        /// <summary>PngQuant.exeファイルのパス</summary>
        private string _exePath { get; set; }

        /// <summary>アンマネージ リソースの解放およびリセットに関連付けられているアプリケーション定義のタスクを実行します。</summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>Pngquant.exeを検索し、見つからない場合Falseを返します。</summary>
        /// <returns>Pngquant.exeが見つかった場合、true</returns>
        public bool SearchPngquantExe()
        {
            string appPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            appPath = Path.GetDirectoryName(appPath);

            // exeの検索
            string exePath = Path.Combine(appPath, "pngquant.exe");
            if (File.Exists(exePath))
            {
                this._exePath = exePath;
                return true;
            }

            // ZIPファイルの検索、見つからなければ、インターネットからダウンロードを試みる
            string zipPath = Path.Combine(appPath, "pngquant-windows.zip");
            if (!File.Exists(zipPath))
            {
                using (var client = new System.Net.WebClient())
                {
                    Console.WriteLine("pngquant.exeをダウンロードします");
                    client.DownloadFile("https://pngquant.org/pngquant-windows.zip", zipPath);
                }
            }

            // ZIPファイルの展開
            using (var zipFile = ZipFile.OpenRead(zipPath))
            {
                foreach (var entry in zipFile.Entries)
                {
                    if (entry.Name.ToLower() != "pngquant.exe")
                    {
                        continue;
                    }

                    entry.ExtractToFile(exePath);
                    this._exePath = exePath;
                    return true;
                }
            }

            return false;
        }

        /// <summary>指定した画像ファイルを減色します。</summary>
        /// <param name="entry">ZIPエントリ</param>
        /// <returns>ファイルパス</returns>
        public string Subtractive(ZipArchiveEntry entry)
        {
            string entryFilePath = Path.Combine(this.TemporaryFolderPath, Path.GetRandomFileName() + ".png");

            if (this.size != null)
            {
                ImageProcessor.Imaging.Formats.PngFormat pngFormat = new ImageProcessor.Imaging.Formats.PngFormat();
                ResizeLayer resizeLayer = new ResizeLayer(new Size(1400, 800), ResizeMode.Min);
                this.imageFactory.Load(ImageExtensions.FromEntry(entry))
                    .Resize(resizeLayer)
                    .Format(pngFormat)
                    .Save(entryFilePath);
            }
            else
            {
                entry.ExtractToPngFile(entryFilePath);
            }

            this.Subtractive(entryFilePath);
            return entryFilePath;
        }

        /// <summary>指定した画像ファイルを減色します。</summary>
        /// <param name="stream">ストリーム</param>
        /// <returns>ファイルパス</returns>
        public string Subtractive(Stream stream)
        {
            string entryFilePath = Path.Combine(this.TemporaryFolderPath, Path.GetRandomFileName() + ".png");

            if (this.size != null)
            {
                ImageProcessor.Imaging.Formats.PngFormat pngFormat = new ImageProcessor.Imaging.Formats.PngFormat();
                ResizeLayer resizeLayer = new ResizeLayer(new Size(1400, 800), ResizeMode.Min);
                this.imageFactory.Load(Image.FromStream(stream))
                    .Resize(resizeLayer)
                    .Format(pngFormat)
                    .Save(entryFilePath);
            }
            else
            {
                stream.ExtractToPngFile(entryFilePath);
            }

            this.Subtractive(entryFilePath);
            return entryFilePath;
        }

        /// <summary>指定した画像ファイルを減色します。</summary>
        /// <param name="filePath">pngファイルのファイルパス</param>
        /// <returns>ファイルパス</returns>
            public string Subtractive(string filePath)
        {
            // pngquant.exeが展開されていない場合、展開します。
            if (this._exePath == null)
            {
                bool existsExe = this.SearchPngquantExe();
                if (!existsExe)
                {
                    return filePath;
                }
            }

            string argument = " --force --verbose --ext=.png --speed=1 --ordered 256 {0}";
            argument = string.Format(argument, filePath);

            Process process = new Process();
            process.StartInfo.FileName = this._exePath;
            process.StartInfo.Arguments = argument;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
            process.Start();
            process.WaitForExit();
            return filePath;
        }

        /// <summary>アンマネージ リソースの解放およびリセットに関連付けられているアプリケーション定義のタスクを実行します。</summary>
        /// <param name="disposing">ディスポーズ</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!this._disposed)
            {
                if (disposing)
                {
                    this.imageFactory.Dispose();

                    // 一時フォルダを削除します。
                    Directory.Delete(this.temporaryFolderPath, true);
                }
            }

            this._disposed = true;
        }
    }
}
