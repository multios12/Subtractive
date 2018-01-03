namespace Subtractive
{
    using System.Drawing;
    using System.IO;
    using System.IO.Compression;

    /// <summary>
    /// エントリ拡張クラス
    /// </summary>
    public static class ZipArchiveEntryExtensions
    {
        /// <summary>
        /// 指定されたエントリをpngファイルとして保存します
        /// </summary>
        /// <param name="entry">エントリ</param>
        /// <param name="filePath">ファイルパス</param>
        public static void ExtractToPngFile(this ZipArchiveEntry entry, string filePath)
        {
            if (Path.GetExtension(entry.Name).ToLower() != ".png")
            {
                using (Stream stream = entry.Open())
                using (var image = Image.FromStream(stream))
                {
                    image.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                }
            }
            else
            {
                entry.ExtractToFile(filePath);
            }
        }
    }
}
