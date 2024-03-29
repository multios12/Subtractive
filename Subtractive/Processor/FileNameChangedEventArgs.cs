﻿namespace Subtractive.Processor
{
    using System;

    /// <summary>
    /// ファイルネーム変更イベントデータ
    /// </summary>
    public class FileNameChangedEventArgs : EventArgs
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="oldFileName">変更前ファイル名</param>
        /// <param name="newFileName">変更後ファイル名</param>
        public FileNameChangedEventArgs(string oldFileName, string newFileName)
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
