namespace Subtractive.Processor
{
    using System;

    /// <summary>
    /// ファイルネーム変更イベントデータ
    /// </summary>
    public class FileNameChangedEventArgs : EventArgs
    {
        /// <summary>変更前ファイル名</summary>
        public string OldFileName { get; set; }

        /// <summary>変更後ファイル名</summary>
        public string NewFileName { get; set; }
    }
}
