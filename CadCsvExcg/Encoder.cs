using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CadCsvExcg
{
    public enum Encoder
    {
        UTF7 = 0,
        UTF8 = 1,
        UTF16LE = 2,
        UTF16BE = 3,
        UTF32 = 4
    }
    static class EncoderMethods
    {
        public static Encoding GetEncoding(this Encoder encoder)
        {
            switch (encoder)
            {
                case Encoder.UTF32:
                    return Encoding.UTF32;
                case Encoder.UTF16BE:
                    return Encoding.BigEndianUnicode;
                case Encoder.UTF16LE:
                    return Encoding.Unicode;
                case Encoder.UTF8:
                    return Encoding.UTF8;
                case Encoder.UTF7:
                    return Encoding.UTF7;
            }
            return Encoding.UTF8;
        }
    }
}
