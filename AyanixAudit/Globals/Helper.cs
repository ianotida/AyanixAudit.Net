using System;

namespace AyanixAudit
{
    public static class Helper
    {
        public static string Title(string sTitle)
        {
            string sResult = "";

            sResult = " ===================================================================================================================" + Environment.NewLine;
            sResult += "  " + sTitle + Environment.NewLine;
            sResult += " ===================================================================================================================" + Environment.NewLine;

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
            double size = (Math.Sign(byteCount) * num);

            string sSize = size.ToString() + suf[place];

            if(suf[place] == " GB")
            {
                if (size > 400 && size < 600) sSize = "500 GB";
                if (size > 900 ) sSize = "1 TB";
            }

            if(suf[place] == " TB")
            {
                if (size > 1.5 ) sSize = "2 TB";
                if (size > 3.5 ) sSize = "4 TB";
            }

            return sSize;
        }

        public static string ToSize(ulong bytes)
        {
            ulong unit = 1024;
            if (bytes < unit) { return $"{bytes} B"; }

            var exp = (int)(Math.Log(bytes) / Math.Log(unit));
            return $"{bytes / Math.Pow(unit, exp):F2} {("KMGTPE")[exp - 1]}B";
        }




    }

}
