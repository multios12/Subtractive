namespace Subtractive.Processor
{
    using System;
    using System.Drawing;

    /// <summary>
    /// 対応しないファイルのための処理クラス
    /// </summary>
    public class NullProcessor : IProcessor
    {
        /// <summary>拡張子</summary>
        public string[] Extensions => new string[] { string.Empty };

        /// <summary>リサイズ時の最大サイズ</summary>
        public Size? ResizeSize { get; set; } = null;

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        public void Execute(string filePath)
        {
            Console.WriteLine(Properties.Resources.CopyRight + Environment.NewLine + "このファイルには対応していません。");
        }
    }
}
