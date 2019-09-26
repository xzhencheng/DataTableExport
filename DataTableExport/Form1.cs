using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace DataTableExport
{
    public partial class Form1 : Form
    {
        private string sqlStr = @"SELECT
                             表名       = Case When A.colorder=1 Then D.name Else '' End,
                             表说明     = Case When A.colorder=1 Then isnull(F.value,'') Else '' End,
                             字段序号   = A.colorder,
                             字段名     = A.name,
                             字段说明   = isnull(G.[value],'') ,
                             标识       = Case When COLUMNPROPERTY( A.id,A.name,'IsIdentity')=1 Then '√'Else '' End,
                             主键       = Case When exists(SELECT 1 FROM sysobjects Where xtype='PK' and parent_obj=A.id and name in (
                                              SELECT name FROM sysindexes WHERE indid in( SELECT indid FROM sysindexkeys WHERE id = A.id AND colid=A.colid))) then '√' else '' end,
                             类型       = B.name,
                             占用字节数 = A.Length,
                             长度       = COLUMNPROPERTY(A.id,A.name,'PRECISION'),
                             小数位数   = isnull(COLUMNPROPERTY(A.id,A.name,'Scale'),0),
                             允许空     = Case When A.isnullable=1 Then '√'Else '' End,
                             默认值     = isnull(E.Text,'')
		                         FROM
			                         syscolumns A
		                         Left Join
			                         systypes B
		                         On
			                         A.xusertype=B.xusertype
		                         Inner Join
			                         sysobjects D
		                         On
			                         A.id=D.id  and D.xtype='U' and  D.name<>'dtproperties'
		                         Left Join
			                         syscomments E
		                         on
			                         A.cdefault=E.id
		                         Left Join
		                         sys.extended_properties  G
		                         on
			                         A.id=G.major_id and A.colid=G.minor_id
		                         Left Join
		 
		                         sys.extended_properties F
		                         On
			                         D.id=F.major_id and F.minor_id=0
			                         where d.name='{0}'
			                         --如果只查询指定表,加上此条件
		                         Order By
			                         A.id,A.colorder 
                                ";

        public Form1()
        {

            InitializeComponent();


        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox5.Text != "")
            {
                button1.Name = "导出中...";
                var link = textBox1.Text;
                var user = textBox2.Text;
                var psw = textBox3.Text;
                var dbName = textBox4.Text;
                List<string> tableList = new List<string>();
                string connectionString = string.Format("Data Source={0};Initial Catalog={1};Persist Security Info=True;User ID={2};Password={3}", link, dbName, user, psw);

                var TableNames = GetDataSet(connectionString, "select  * from sysobjects D where D.xtype='U' and  D.name<>'dtproperties'", "tables");

                foreach (DataRow row in TableNames.Tables[0].Rows)
                {
                    tableList.Add(row[0].ToString());
                }
                DataSet ds = new DataSet();
                foreach (var item in tableList)
                {

                    var Table = GetDataSet(connectionString, string.Format(sqlStr, item), item).Tables[0];

                    ds.Tables.Add(Table.Copy());
                }

                NPOIHelper.ExportEasy(ds, textBox5.Text);
                button1.Name = "导出数据";

                MessageBox.Show("数据字典导出成功", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("请选择导出路径", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }

        public static SqlDataReader ExecuteReader(string connectionString, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = new SqlCommand();
            SqlConnection conn = new SqlConnection(connectionString);

            // we use a try/catch here because if the method throws an exception we want to 
            // close the connection throw code, because no datareader will exist, hence the 
            // commandBehaviour.CloseConnection will not work
            try
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                SqlDataReader rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                cmd.Parameters.Clear();
                return rdr;
            }
            catch
            {
                conn.Close();
                throw;
            }
        }


        private static void PrepareCommand(SqlCommand cmd, SqlConnection conn, SqlTransaction trans, CommandType cmdType, string cmdText, SqlParameter[] cmdParms)
        {
            if (conn.State != ConnectionState.Open)
                conn.Open();

            cmd.Connection = conn;
            cmd.CommandText = cmdText;

            if (trans != null)
                cmd.Transaction = trans;

            cmd.CommandType = cmdType;

            if (cmdParms != null)
            {
                foreach (SqlParameter parm in cmdParms)
                    cmd.Parameters.Add(parm);
            }
        }

        //查询数据库返回DataSet
        public static DataSet GetDataSet(string connString, string sql, string tablename)
        {
            using (SqlConnection conn = new SqlConnection(connString))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter(sql, conn))
                {
                    DataSet ds = new DataSet();
                    conn.Open();
                    sda.Fill(ds, tablename);
                    return ds;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "Excle文件(*.xls)|*.xls";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox5.Text = dialog.FileName;
            }
        }
    }
}
