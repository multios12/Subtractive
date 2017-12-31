namespace Subtractive
{
    using System.IO;

    /// <summary>
    /// ファイルパスユーティリティクラス
    /// </summary>
    public static class PathUtils
    {
        /// <summary>
        /// 指定されたファイルパスの拡張子を変更します。
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="extension">拡張子</param>
        /// <returns>変更されたファイルパス</returns>
        public static string ChangeExtension(string filePath, string extension)
        {
            return Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + extension);
        }
    }
}
