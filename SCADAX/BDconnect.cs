using System;
using System.Data;
using  FirebirdSql.Data.FirebirdClient;


namespace SCADAX
{
    class BDconnect
    {
        FbConnection _fb;
        string txtPath; 
        public BDconnect(string txtPath) 
        {
            this.txtPath = txtPath;
            Connect();
        }

        public void InsBd(string str)
        {
            var insertSql = new FbCommand(str, _fb);
            if (_fb.State == ConnectionState.Closed)
                _fb.Open();
            var fbt = _fb.BeginTransaction();
            insertSql.Transaction = fbt;
            try
            {
                int res = insertSql.ExecuteNonQuery();
                //MessageBox.Show("SUCCESS: " + res.ToString());
                fbt.Commit();
            }
            catch (Exception )
            {
                //MessageBox.Show(ex.Message);
                //InsertSQL.Dispose();
                //fb.Close();
                //InsBD(str);

            }
            insertSql.Dispose();
            _fb.Close();

        }
        public DataTable GetBD(string str)
        {
            var sqLcomm = new FbCommand(str, _fb);
            if (_fb.State == ConnectionState.Closed) _fb.Open();

            var fbt = _fb.BeginTransaction();
            sqLcomm.Transaction = fbt;
            var fbda = new FbDataAdapter(sqLcomm);
            var ds = new DataSet();
            try
            {
                fbda.Fill(ds);
                _fb.Close();
                return ds.Tables[0];
            }
            catch (Exception )
            {
                sqLcomm.Dispose();
                _fb.Close();
                return null;
            }
        }
        private void Connect()
        {
            var fb_con = new FbConnectionStringBuilder();
            fb_con.Charset = "WIN1251";
            fb_con.UserID = "sysdba";
            fb_con.Password = "masterkey";
            fb_con.Database = txtPath;
            fb_con.ServerType = 0;
            //fb_con.ClientLibrary = @"C:\Program Files\FirebirdClient.dll";
            _fb = new FbConnection(fb_con.ToString());
            _fb.Open();
            var fbInf = new FbDatabaseInfo(_fb);            
        }
    }
}
