namespace Subtractive
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.IO.Compression;

    /// <summary>PNG減色処理クラス</summary>
    public class PngQuant : IDisposable
    {
        /// <summary>コンストラクタ</summary>
        public PngQuant()
        {
            // 一時フォルダの作成
            this._temporaryFolderPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(this._temporaryFolderPath);
        }

        /// <summary>disposeが完了した場合、true</summary>
        private bool _disposed { get; set; }

        /// <summary>PngQuant.exeファイルのパス</summary>
        private string _exePath { get; set; }

        /// <summary>一時フォルダパス</summary>
        private string _temporaryFolderPath { get; set; }

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

        /// <summary>指定したpngファイルを減色します。</summary>
        /// <param name="filePath">pngファイルのファイルパス</param>
        public void Subtractive(string filePath)
        {
            // pngquant.exeが展開されていない場合、展開します。
            if (this._exePath == null)
            {
                bool existsExe = this.SearchPngquantExe();
                if (!existsExe)
                {
                    return;
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
        }

        /// <summary>指定したストリームをファイルに保存し、減色します。</summary>
        /// <param name="entry">エントリ</param>
        /// <returns>ファイルパス</returns>
        public string SubtractiveToTemporaryFile(ZipArchiveEntry entry)
        {
            string entryFilePath = Path.Combine(this._temporaryFolderPath, Path.GetRandomFileName() + ".png");
            entry.ExtractToFile(entryFilePath);
            this.Subtractive(entryFilePath);
            return entryFilePath;
        }

        /// <summary>アンマネージ リソースの解放およびリセットに関連付けられているアプリケーション定義のタスクを実行します。</summary>
        /// <param name="disposing">ディスポーズ</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!this._disposed)
            {
                if (disposing)
                {
                    // 一時フォルダを削除します。
                    Directory.Delete(this._temporaryFolderPath, true);
                }
            }

            this._disposed = true;
        }
    }
}
