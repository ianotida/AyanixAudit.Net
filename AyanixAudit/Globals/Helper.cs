using System;

namespace AyanixAudit
{
    public static class Helper
    {
        public static string Title(string sTitle)
        {
            string sResult = "";

            sResult = " ==============================================================================================================" + Environment.NewLine;
            sResult += "  " + sTitle + Environment.NewLine;
            sResult += " ==============================================================================================================" + Environment.NewLine;

            return sResult;
        }

        public static string PadString(string sMsg, int iLevel)
        {
            return sMsg.PadRight(iLevel);
        }

        public static string FormatSize(long byteCount)
        {
            string[] suf = { " B", " KB", " MB", " GB", " TB", " PB", " EB" }; //Longs run out around EB
            if (byteCount == 0) return "0" + suf[0];

            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);

            return (Math.Sign(byteCount) * num).ToString() + suf[place];
        }      

    }
}
