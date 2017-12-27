namespace Subtractive
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.IO.Compression;
    using Subtractive.Properties;

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

        /// <summary>指定したpngファイルを減色します。</summary>
        /// <param name="filePath">pngファイルのファイルパス</param>
        public void Subtractive(string filePath)
        {
            // pngquant.exeが展開されていない場合、展開します。
            if (this._exePath == null)
            {
                this._exePath = Path.Combine(this._temporaryFolderPath, "pngquant.exe");
                using (MemoryStream s = new MemoryStream(Resources.pngquant))
                {
                    ZipArchive archive = new ZipArchive(s);
                    archive.GetEntry("pngquant.exe").ExtractToFile(this._exePath);
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
                    // pngquant.exeが展開されている場合、削除します。
                    if (File.Exists(this._exePath) == true)
                    {
                        File.Delete(this._exePath);
                    }

                    // 一時フォルダを削除します。
                    Directory.Delete(this._temporaryFolderPath, true);
                }
            }

            this._disposed = true;
        }
    }
}
