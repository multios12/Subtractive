namespace Subtractive
{
    using System.Diagnostics;
    using System.IO;
    using Microsoft.Win32;

    /// <summary>
    /// MDBファイル圧縮処理クラス
    /// </summary>
    public class MdbCompactation
    {
        /// <summary>
        /// 指定されたMDBファイルを圧縮します。
        /// </summary>
        /// <param name="filePaths">MDBファイル</param>
        public static void compact(params string[] filePaths)
        {
            // mdbファイルに関連付けられている実行コマンドを取得する
            string exePath = GetShellCommandFromClassesRoot("Access.MDBFile");

            if (string.IsNullOrEmpty(exePath) == true)
            {
                return;
            }

            foreach (string filePath in filePaths)
            {
                // 関連付け情報を利用してプロセスを起動し、圧縮を行う
                Process process = new Process();
                process.StartInfo.FileName = filePath;
                process.StartInfo.Arguments = @"/compact";
                process.Start();
                process.WaitForExit(10000);
            }
        }

        #region 関連付け情報操作
        /// <summary>
        /// 指定されたファイルに関連付けられたコマンドを取得する
        /// </summary>
        /// <param name="filePath">関連付けを調べるファイル</param>
        /// <param name="extra">アクション(open,print,editなど)</param>
        /// <returns>取得できた時は、コマンド(実行ファイルのパス+コマンドライン引数)。
        /// 取得できなかった時は、空の文字列。</returns>
        public static string FindAssociatedCommand(string filePath, string extra)
        {
            // 拡張子を取得
            string extName = Path.GetExtension(filePath);

            if (extName.Length == 0 || extName[0] != '.')
            {
                return string.Empty;
            }

            // HKEY_CLASSES_ROOT\(extName)\shell があれば、
            // HKEY_CLASSES_ROOT\(extName)\shell\(extra)\command の標準値を返す
            if (ExistClassesRootKey(extName + @"\shell"))
            {
                return GetShellCommandFromClassesRoot(extName, extra);
            }

            // HKEY_CLASSES_ROOT\(extName) の標準値を取得する
            string fileType = GetDefaultValueFromClassesRoot(extName);
            if (fileType.Length == 0)
            {
                return string.Empty;
            }

            // HKEY_CLASSES_ROOT\(fileType)\shell\(extra)\command の標準値を返す
            return GetShellCommandFromClassesRoot(fileType, extra);
        }

        /// <summary>
        /// 指定されたファイルに関連付けられたコマンドを取得します。
        /// </summary>
        /// <param name="filePath">関連付けを調べるファイル</param>
        /// <returns>
        /// 取得できた時は、コマンド(実行ファイルのパス+コマンドライン引数)。<br />
        /// 取得できなかった時は、空の文字列。
        /// </returns>
        public static string FindAssociatedCommand(string filePath)
        {
            return FindAssociatedCommand(filePath, "open");
        }

        /// <summary>指定されたキーがレジストリに存在するか確認します。</summary>
        /// <param name="keyName">キー名</param>
        /// <returns>取得した値</returns>
        private static bool ExistClassesRootKey(string keyName)
        {
            using (var regKey = Registry.ClassesRoot.OpenSubKey(keyName))
            {
                return regKey != null;
            }
        }

        /// <summary>レジストリから指定されたキーの値を取得します。</summary>
        /// <param name="keyName">キー名</param>
        /// <returns>取得した値</returns>
        private static string GetDefaultValueFromClassesRoot(string keyName)
        {
            using (var regKey = Registry.ClassesRoot.OpenSubKey(keyName))
            {
                if (regKey == null)
                {
                    return string.Empty;
                }

                return (string)regKey.GetValue(string.Empty, string.Empty);
            }
        }

        /// <summary>指定された拡張子に関連付けられた実行コマンドを返します。</summary>
        /// <param name="fileType">拡張子（ファイル種別）</param>
        /// <param name="extra">ファイルに実行するアクション</param>
        /// <returns>実行ファイルパス</returns>
        private static string GetShellCommandFromClassesRoot(string fileType, string extra = "open")
        {
            if (extra.Length == 0)
            {
                // アクションが指定されていない時は、既定のアクションを取得する
                extra = GetDefaultValueFromClassesRoot(fileType + @"shell")
                    .Split(',')[0];

                if (extra.Length == 0)
                {
                    extra = "open";
                }
            }

            return GetDefaultValueFromClassesRoot(
                string.Format(@"{0}\shell\{1}\command", fileType, extra));
        }
    }
    #endregion
}
