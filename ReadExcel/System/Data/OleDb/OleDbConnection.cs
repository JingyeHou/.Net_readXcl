using Microsoft.Office.Interop.Excel;

namespace System.Data.OleDb
{
    internal class OleDbConnection : OLEDBConnection
    {
        private string strConn;

        public OleDbConnection(string strConn)
        {
            this.strConn = strConn;
        }
    }
}