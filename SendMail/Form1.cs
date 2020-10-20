using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace SendMail
{
    public partial class Form1 : Form
    {
        bool blBandera = false;
        public string _strbody = "";
        public string _strfilename = "";
        string[] Dias = new string[2] { "SÁBADO", "DOMINGO" };
        int HoraMercadoInicial = 7;
        int HoraMercadoFinal = 14;
        int errorCode = 0;
        //public string TipoOperacion;
        public string soporte = "mcastillo@bvrd.com.do";
        private DataTable dtenvios = new DataTable();

        public static string ConectionString
        {
            get
            {

                //return ConfigurationManager.AppSettings["INTERFACE_ConectionString"].ToString();
                //return @"Data Source=ARP01\SQL_REPORT_01;Initial Catalog=SIOPEL_INTERFACE_DB;User ID=developer;Password=admin@123";
                return @"Data Source=ADB03;Initial Catalog=SIOPEL_INTERFACE_DB;User ID=developer;Password=admin@123";
            }

        }
        public Form1()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            EnviosEstadisticas();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //_Enviar("trabajando cliente de correo", "hola mundo", "", soporte);
            EnvioRecursivo();
        }

        private void _Enviar(string Subject, string Body, string Attachments, string To)
        {
            string server = Environment.MachineName.ToString();
            Subject = server + " : " + Subject;
            SmtpClient MailClient = new SmtpClient("smtp.office365.com", 587);            
            //SmtpClient MailClient = new SmtpClient("smtp.office365.com", 25);
            string From = "notificaciones@bvrd.com.do";
            //string To = "mcastillo@bvrd.com.do"; //tecnologia
            string AuthUsername = "notificaciones@bvrd.com.do";
            string AuthPassword = "Juko6315*f2";
            //Subject = AppName.ToString() + " - " + Subject;
            // MailClient.Credentials = new System.Net.NetworkCredential(From, "");
            // MailClient.Credentials = new System.Net.NetworkCredential(AuthUsername, AuthPassword);
            var MailMessage = new MailMessage(From, To, Subject, Body);
            MailMessage.CC.Add("melvi18@gmail.com");
            MailMessage.IsBodyHtml = true;

            if ((AuthUsername != null) && (AuthPassword != null))
            {
                {
                    //var withBlock = MailClient;
                    MailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                    //MailClient.UseDefaultCredentials = true;
                    MailClient.EnableSsl = true;
                    MailClient.Credentials = new System.Net.NetworkCredential(AuthUsername, AuthPassword);
                }
            }

            if ((Attachments.ToString() != ""))
            {
                //foreach (var FileName in Attachments)
                MailMessage.Attachments.Add(new Attachment(Attachments));
            }
            MailClient.Send(MailMessage);
            MailClient.Dispose();
            MailMessage.Dispose();
        }

        public DataSet DsOperaciones(/*string TipoOperacion*/)
        {
            /* Reviso si existen operacinoes para enviar*/
            DataSet ds = new DataSet();
            try
            {
                string strString = ConectionString;
                using (SqlConnection con = new SqlConnection(strString))
                {
                    using (SqlCommand cmd = new SqlCommand("P_ENVIOBOLETAData2", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        //cmd.Parameters.Add("@TipoOperacion", SqlDbType.VarChar).Value = TipoOperacion;
                        con.Open();

                        SqlDataAdapter adp = new SqlDataAdapter(cmd);
                        adp.Fill(ds);
                    }
                }
                return ds;
            }
            catch (Exception ex)
            {

                _strbody = "SendMail - Error: " + ex.Message.ToString() + " " + "; NOTA:  Enviando Correo de las boletas";//+ TipoOperacion;
                _strfilename = "";
                _Enviar("Error EJECUCION SendMail ", _strbody, _strfilename.ToString(), soporte);
                return ds;
            }
        }

        public void EnvioRecursivo()
        {

            int horaMercado = -1;
            DateTime Hoy = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            string NombredelDia = Hoy.ToString("dddd", new System.Globalization.CultureInfo("es-ES")).ToUpper();
            //if (Dias.)
            bool ValidaFindeSemana = Dias.Contains(NombredelDia);
            horaMercado = Convert.ToInt16(DateTime.Now.ToString("HH"));
            if (ValidaFindeSemana == false)
            {
                if (horaMercado > HoraMercadoInicial && horaMercado < HoraMercadoFinal)
                {
                    /* Envio la Boleta a los compradores */
                    //TipoOperacion = "COMPRA";
                    DataSet dsoperaciones = DsOperaciones(/*TipoOperacion*/);
                    int i = 0;
                    #region 
                    #endregion
                    if (dsoperaciones.Tables[0].Rows.Count > 0)
                    {
                        while (i <= dsoperaciones.Tables[0].Rows.Count - 1)
                        {
                            try
                            {
                                //_Enviar("Boleta de Operación", dsoperaciones.Tables[0].Rows[i]["Cuerpo"].ToString(), "", dsoperaciones.Tables[0].Rows[i]["EmailComprador"].ToString());
                                _Enviar("Boleta_Pruebas de Operación - Prueba", dsoperaciones.Tables[0].Rows[i]["Cuerpo"].ToString().Replace("@TipoOperacion", "COMPRA"), "", soporte);
                                _Enviar("Boleta_Pruebas de Operación - Prueba", dsoperaciones.Tables[0].Rows[i]["Cuerpo"].ToString().Replace("@TipoOperacion", "VENTA"), "", soporte);
                                MarcaOperacionEnviada(Convert.ToInt64(dsoperaciones.Tables[0].Rows[i]["NUMEROOPERACION"].ToString()), Convert.ToDateTime(dsoperaciones.Tables[0].Rows[i]["FECHAOPERACION"].ToString()));
                                //errorCode = 1;
                            }
                            catch (Exception ex)
                            {
                            }
                            i = i + 1;
                        } // while 
                    }//if (ds.Tables[0].Rows.Count > 0) 
                     /* Refresco data del cuadro de envios */
                    EnviosEstadisticas();
                } // if (moment > 7 && moment < 14
            } //if (ValidaFindeSemana == false) 
        }

        public void MarcaOperacionEnviada(Int64 numoperacion, DateTime FechaOperacion)
        {
            DataSet ds = new DataSet();
            try
            {

                string strString = ConectionString;
                using (SqlConnection con = new SqlConnection(strString))
                {
                    using (SqlCommand cmd = new SqlCommand("P_ENVIOBOLETAMarcoEnviada", con))
                     {
                         cmd.CommandType = CommandType.StoredProcedure;
                         con.Open();
                         cmd.Parameters.Add("@NUMEROOPERACION", SqlDbType.Decimal).Value = numoperacion;
                         cmd.Parameters.Add("@FECHAOPERACION", SqlDbType.Date).Value = FechaOperacion;
                         int result = cmd.ExecuteNonQuery();
                     } 
                }
            }
            catch (Exception ex)
            {
                _strbody = "SendMail - Error: " + ex.Message.ToString() + " " + "; MARCANDO LA OPERACION #: " + numoperacion.ToString();
                _strfilename = "";
                _Enviar("Error EJECUCION SendMail ", _strbody, _strfilename.ToString(), soporte);
            }
        }

        public void EnviosEstadisticas()
        {
            DataSet ds = new DataSet();
            try
            {

                string strString = ConectionString;
                using (SqlConnection con = new SqlConnection(strString))
                {
                     
                    using (SqlCommand cmd = new SqlCommand("P_ENVIOBOLETAEstadistica", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        con.Open();
                        //cmd.Parameters.Add("@NUMEROOPERACION", SqlDbType.Decimal).Value = numoperacion;
                        //cmd.Parameters.Add("@FECHAOPERACION", SqlDbType.Date).Value = FechaOperacion;
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        dtenvios.Clear();
                        da.Fill(dtenvios);
                        con.Close();
                        da.Dispose();
                        /*dgEnvios.DataSource = null;
                        dgEnvios.Rows.Clear(); 
                        dgEnvios.Columns.Clear();
                        dgEnvios.Refresh();*/
                        dgEnvios.DataSource = dtenvios;
                     
                        //int result = cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                _strbody = "SendMail - Error: " + ex.Message.ToString() + " " + "; Generando Estadisticas de los Envios ";
                _strfilename = "";
                _Enviar("Error EJECUCION SendMail ", _strbody, _strfilename.ToString(), soporte);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            stLapso.Enabled = true;
            stLapso.Start();
            label1.Text = "PROCESO INICIADO";
            label1.ForeColor = Color.Blue;
        }

        private void stLapso_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (blBandera) return;

            try
            {
                EnvioRecursivo();

            }// OnDebug();}
            catch (Exception ex)
            {
               // WindowsService1.MoneyMarketCevaldom MM = new WindowsService1.MoneyMarketCevaldom();
               // stLapso.Stop();
                //string _strbody = " SE DETUVO EL SERVICIO DEL ENVIO DE LAS OPERACIONES HACIA CEVALDOM; ERROR: " + ex.Message;
                //string _strfilename = "";
                //MM.SendMail("Error ENVIO WS Operaciones - Cevaldom", _strbody, _strfilename.ToString());
                //EventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
            }
            blBandera = false;
        }

        private void stLapso_Tick(object sender, EventArgs e)
        {
            EnvioRecursivo();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            stLapso.Enabled = false;
            stLapso.Stop();
            label1.Text = "PROCESO DETENIDO";
            label1.ForeColor = Color.Red;
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        { 
        }
    }
}


