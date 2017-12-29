namespace Subtractive.Processor
{
    using System;

    /// <summary>
    /// ファイルネーム変更イベントデータ
    /// </summary>
    public class QuantedEventArgs : EventArgs
    {
        public QuantedEventArgs(string oldFileName, string newFileName)
        {
            this.OldFileName = oldFileName;
            this.NewFileName = newFileName;
        }

        /// <summary>変更前ファイル名</summary>
        public string OldFileName { get; set; }

        /// <summary>変更後ファイル名</summary>
        public string NewFileName { get; set; }
    }
}
