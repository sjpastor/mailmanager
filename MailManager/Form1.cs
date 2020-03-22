using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
//using System.Data.Datable;
//using System.Data.OleDb.OleDbDataReader;
using System.IO;
using OpenPop.Pop3;
using MailManager.Opciones;

namespace GetMail
{
    public partial class frmPrincipal : Form
    {

        private String db = "GetMail.accdb";
        private OleDbConnection con;
        private const String CADENA_CONEXION = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=";


        public frmPrincipal()
        {
            InitializeComponent();
        }

        private long GetLastIndice()
        {
            String sql;
            OleDbCommand comm;
            OleDbDataReader reader;
            long idx = -1;

            try
            {
                sql = "";
                sql += "SELECT MAX(Carpeta) as idx";
                sql += "  FROM Correos";

                SimpleLogger.WriteLog(sql);
                comm = new OleDbCommand(sql, con);
                reader = comm.ExecuteReader();
                if(reader.Read())
                {
                    idx = Convert.ToInt64(reader["idx"]);
                }
                reader.Close();
                return idx;
            }
            catch (Exception ex)
            {
                SimpleLogger.WriteLog(ex.Message);
                return -1;
            }

        }

        private String formateaCarpeta(long indice)
        {
            return String.Format("{0:D10}", indice);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LeerConOpenPop();
            //LeerConOpenPop2();
        }

        private void LeerConOpenPop() { 
            String sql;
            OleDbCommand comm;
            long indice;
            String carpeta;

            indice = GetLastIndice();

            //creamos el objeto
            ConnectPop3 oCP3 = new ConnectPop3();

            //invocamos el metodo para obtener mensajes
            List<OpenPop.Mime.Message> lstMensajes = oCP3.getMensajes();

            //recorremos y mostramos el asunto
            foreach (OpenPop.Mime.Message oMensaje in lstMensajes)
            {
                indice++;
                carpeta = formateaCarpeta(indice);

                try
                {
                    sql = "";
                    sql += "INSERT INTO Correos (MensajeId, Fecha, EnviadoPor, EnviadoA, Asunto, Cuerpo, Carpeta)";
                    sql += " VALUES ('" + oMensaje.Headers.MessageId + "',";
                    sql += "  #" + Convert.ToDateTime(oMensaje.Headers.Date).ToString("MM/dd/yyyy") + "#,";
                    sql += "  '" + oMensaje.Headers.From.ToString().Replace("\"","") + "',";
                    sql += "  '" + GetTo(oMensaje) + "',";
                    sql += "  '" + oMensaje.Headers.Subject + "',";
                    if (oMensaje != null)
                    {
                        OpenPop.Mime.MessagePart messagePart = oMensaje.MessagePart.MessageParts[1];

                        string test = messagePart.BodyEncoding.GetString(messagePart.Body);
                        // Raw mensaje
                        var mailbody = ASCIIEncoding.ASCII.GetString(oMensaje.RawMessage);
                        
                        //// Si es UTF-8
                        if (mailbody.Contains("utf-8")){
                            //var encodedStringAsBytes = System.Convert.FromBase64String(mailbody);
                            //var rawMessage = System.Text.Encoding.UTF8.GetString(mailbody);
                        }
                        sql += "  '" + mailbody + "',";
                    }
                    else {
                        sql += "  '',";
                    }
                    sql += "  '" + carpeta + "')";
                    SimpleLogger.WriteLog(sql);
                    comm = new OleDbCommand(sql, con);
                    comm.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    SimpleLogger.WriteLog(ex.Message);
                }
            }

            MessageBox.Show("OK");
        }

        private String GetTo(OpenPop.Mime.Message oMensaje)
        {
            String retorno = "";

            for (int i=0; i<=oMensaje.Headers.To.Count-1; i++){
                retorno += oMensaje.Headers.To[i] + "; ";
            }

            return retorno.Substring(1,retorno.Length-2);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            con = new OleDbConnection(CADENA_CONEXION + db);
            try {
                con.Open();
            }
            catch(Exception ex){
                MessageBox.Show(ex.Message, "ERROR");
            }
        }

        private void mniSalir_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void mniConfiguracion_Click(object sender, EventArgs e)
        {
            frmConfiguracion ventana = new frmConfiguracion();

            ventana.ShowDialog();
        }
    }
}
