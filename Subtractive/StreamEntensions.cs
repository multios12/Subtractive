namespace ArchiveXlsx
{
    using System.IO;

    /// <summary>
    /// ストリーム拡張クラス
    /// </summary>
    public static class StreamExtensions
    {
        /// <summary>Streamをバイト配列に変換して返します。</summary>
        /// <param name="stream">ストリーム</param>
        /// <returns>バイト配列</returns>
        public static byte[] ToByteArray(this Stream stream)
        {
            byte[] bs = new byte[stream.Length];
            stream.Read(bs, 0, (int)stream.Length);
            return bs;
        }

        /// <summary>Streamをバイト配列に変換して返します。</summary>
        /// <param name="stream">ストリーム</param>
        /// <param name="length">サイズ</param>
        /// <returns>バイト配列</returns>
        public static byte[] ToByteArray(this Stream stream, int length)
        {
            byte[] bs = new byte[length];
            stream.Read(bs, 0, length);
            return bs;
        }

        /// <summary>StreamをMemoryStreamに変換して返します。</summary>
        /// <param name="stream">ストリーム</param>
        /// <returns>メモリストリーム</returns>
        public static MemoryStream ToMemoryStream(this Stream stream)
        {
            return new MemoryStream(ToByteArray(stream));
        }

        /// <summary>StreamをMemoryStreamに変換して返します。</summary>
        /// <param name="stream">ストリーム</param>
        /// <param name="length">サイズ</param>
        /// <returns>メモリストリーム</returns>
        public static MemoryStream ToMemoryStream(this Stream stream, int length)
        {
            return new MemoryStream(ToByteArray(stream, length));
        }
    }
}
