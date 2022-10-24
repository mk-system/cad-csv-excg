namespace CadCsvExcg
{
    public enum Delimiter
    {
        COMMA = 1,
        TAB = 2,
    }

    static class DelimiterMethods
    {
        public static string GetString(this Delimiter d)
        {
            switch (d)
            {
                case Delimiter.TAB:
                    return "	";
                case Delimiter.COMMA:
                default:
                    return ",";
            }
        }

        public static int GetIndex(this Delimiter d)
        {
            return (int)d;
        }
    }
}