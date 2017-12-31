namespace Subtractive
{
    using System.Drawing;
    using System.IO;
    using System.IO.Compression;

    /// <summary>
    /// イメージ拡張クラス
    /// </summary>
    public static class ImageExtensions
    {
        /// <summary>
        /// 指定したエントリからイメージを作成します
        /// </summary>
        /// <param name="entry">エントリ</param>
        /// <returns>イメージ</returns>
        public static Image FromEntry(ZipArchiveEntry entry)
        {
            using (Stream stream = entry.Open())
            {
                return Image.FromStream(stream);
            }
        }

        /// <summary>
        /// 指定されたエントリをイメージファイルとして読み取り、ファイルに保存します。
        /// </summary>
        /// <param name="entry">エントリ</param>
        /// <param name="filePath">ファイルパス</param>
        public static void FromEntryToSave(ZipArchiveEntry entry, string filePath)
        {
            using (var image = ImageExtensions.FromEntry(entry))
            {
                image.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
            }
        }
    }
}