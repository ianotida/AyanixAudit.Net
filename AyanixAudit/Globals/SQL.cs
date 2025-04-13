﻿using System.Data.SqlClient;
using System.Data;

namespace AyanixAudit
{
    public class SQL
    {
		//192.168.121.210

        public static string _NPCIT
        //{ get { return "Data Source=NPC-DB;Initial Catalog=NPC_IT;User ID=sa;Password=Adm1n@npcdb;Connect Timeout=60"; } }
		{ get { return "Data Source=192.168.121.210;Initial Catalog=NPC_IT;User ID=sa;Password=Adm1n@npcdb;Connect Timeout=10"; } }


		public static bool Check_DB()
		{
			bool bCheck = false;

			DataTable Dtmp = new DataTable();
			using (SqlConnection SC = new SqlConnection(SQL._NPCIT))
			{
				SC.Open();
				SqlCommand CMD = new SqlCommand("SELECT getdate()", SC);
				Dtmp.Load(CMD.ExecuteReader());

				if(Dtmp.Rows.Count > 0) bCheck = true;
			}

			return bCheck;
		}


        public static DataSet Get_DS(string strQry)
		{
			DataSet DS = new DataSet();
			using (SqlConnection SC = new SqlConnection(SQL._NPCIT))
			{
				SC.Open();
				SqlDataAdapter SDA = new SqlDataAdapter(strQry, SC);
				SDA.Fill(DS);
			}
			return DS;
		}

		public static DataTable Get_Table(string strQry)
		{
			DataTable Dtmp = new DataTable();
            using (SqlConnection SC = new SqlConnection(SQL._NPCIT))
            {
                SC.Open();
                SqlCommand CMD = new SqlCommand(strQry, SC);
                Dtmp.Load(CMD.ExecuteReader());
            }
            return Dtmp;
		}

    }

}
