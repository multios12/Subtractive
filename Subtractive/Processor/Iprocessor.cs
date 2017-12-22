namespace Subtractive.Processor
{
    /// <summary>
    /// 処理インターフェイス
    /// </summary>
    public interface IProcessor
    {
        /// <summary>拡張子</summary>
        string[] Extensions { get; }

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        void Execute(string filePath);
    }
}
