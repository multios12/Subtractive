namespace Subtractive.Processor
{
    using System.Drawing;

    /// <summary>
    /// 処理インターフェイス
    /// </summary>
    public interface IProcessor
    {
        /// <summary>拡張子</summary>
        string[] Extensions { get; }

        /// <summary>リサイズ時の最大サイズ</summary>
        Size? ResizeSize { get; set; }

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        void Execute(string filePath);
    }
}
