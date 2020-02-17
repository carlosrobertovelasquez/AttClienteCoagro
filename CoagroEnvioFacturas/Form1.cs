using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using CrystalDecisions.CrystalReports;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Xml;
using CoagroEnvioFacturas.APITIGO;
using System.IO;
using CrystalDecisions;
using System.Globalization;

namespace CoagroEnvioFacturas
{
    public partial class FrmPrincipal : Form
    {
       

        static readonly string rootFolder = @"C:\CoagroAttCliente\Envio";
        bool blBandera = false;
        public FrmPrincipal()
        {
            InitializeComponent();
            if (!System.Diagnostics.EventLog.SourceExists("CoagroEnvioFacturas"))
            {
                System.Diagnostics.EventLog.CreateEventSource("CoagroEnvioFacturas", "MyNewLog");
            }
            eventLog1.Source = "CoagroEnvioFacturas";
            eventLog1.Log = "MyNewLog";
        }

        private void FrmPrincipal_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState==FormWindowState.Minimized)
            {
                this.Hide();
             
                notifyIcon1.BalloonTipText = "Tu Formulario ha sido en segundo Plano";
                notifyIcon1.ShowBalloonTip(1000);
            }
            
            
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
               notifyIcon1.BalloonTipText = "Tu formulario ha sido Normalizado";
               notifyIcon1.ShowBalloonTip(1000);

        }

        private void minimixarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            
                this.Hide();
                //  notifyIcon1.Icon = SystemIcons.Application;
                notifyIcon1.BalloonTipText = "Tu Formulario ha sido en segundo Plano";
                notifyIcon1.ShowBalloonTip(1000);
            
        }

        private void cerrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmPrincipal_Load(object sender, EventArgs e)
        {
           
        }

        private void restaurarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Show();
            notifyIcon1.BalloonTipText = "Tu formulario ha sido Normalizado";
            notifyIcon1.ShowBalloonTip(1000);
        }

        private void stLapso_Tick(object sender, EventArgs e)
        {


            // AQUI TENDRIAMOS QUE BOORAR
            if (blBandera) return;

           ProcesarTarea();
            borrarArchivo();
            EnviarDocVencidosTDias();
            EnviarDocVencidosJueves();
            EnviarPagos();
            blBandera = false;


        }
        private void ProcesarTarea()
        {
            eventLog1.WriteEntry("Ingreso a Time de Cinco Minutos");
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["sConexion"].ToString());


            //IList<CoagroAttCliente> coagroAttClientes = new List<CoagroAttCliente>();
            string connectionString = ConfigurationManager.ConnectionStrings["sConexion"].ToString();
            SqlConnection conn = new SqlConnection(connectionString);
            string sql = @"select * from CINCOH.CoagroAttCliente where EnvioCorreo='N' and EnvioMensajetigo='N' and EnvioEstadocuenta='N' and Procesado='N' and modulo='NEWFACT' ";
            SqlCommand command = new SqlCommand(sql, conn);
            conn.Open();
            try
            {
                blBandera = true;
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    var doc = reader[0].ToString();

                    using (var ada = new SqlDataAdapter(@"SELECT FA.CLIENTE,  FA.FACTURA,FA.EMBARCAR_A,FA.NOMBRE_CLIENTE,FA.TOTAL_FACTURA,FA.ANULADA,
                                                                TIP.TIPO,TIP.SUBTIPO, 
                                                                CASE WHEN (TIP.DESCRIPCION like '%Cons Int Propia%') THEN 'FAC' 
                                                                WHEN (TIP.DESCRIPCION like '%Cont Contribuyente%') THEN 'CCF' 
                                                                WHEN (TIP.DESCRIPCION like 'Cons Int Ticket%') THEN 'TIK' end as TIPODOC,
                                                                FECHA,CL.E_MAIL,CL.TELEFONO1,CL.TELEFONO2,
                                                                CL.RUBRO2_CLIENTE AS ENVIARMENSAJE,
                                                                CL.RUBRO3_CLIENTE AS ENVIARFACTURA,
                                                                CL.RUBRO4_CLIENTE AS ENVIARESTADOCUENTA ,
                                                                CL.RUBRO5_CLIENTE AS ENVIARCORREO,
                                                                ve.telefono as TELEFONOVENDEDOR,
                                                                ve.E_mail AS CORREOVENDEDOR 
                                                                FROM 
                                                                CINCOH.CLIENTE CL, 
                                                                CINCOH.FACTURA FA,
                                                                CINCOH.SUBTIPO_DOC_CC TIP ,
                                                                CINCOH.VENDEDOR VE
                                                                WHERE 
                                                                FA.CLIENTE=CL.CLIENTE AND 
                                                                FA.TIPO_DOC_CXC=TIP.TIPO AND 
                                                                FA.SUBTIPO_DOC_CXC=TIP.SUBTIPO and
                                                                TIP.SUBTIPO in(4,5) and
                                                                FA.VENDEDOR=VE.VENDEDOR and 
                                                                FA.TIPO_DOCUMENTO='F' AND FA.FACTURA=@doc", cn))
                    {
                        DataTable tabla2 = new DataTable();
                        ada.SelectCommand.Parameters.AddWithValue("@doc", doc);
                        ada.SelectCommand.CommandType = CommandType.Text;
                        ada.Fill(tabla2);
                        for (int y = 0; y < tabla2.Rows.Count; y++)
                        {
                            if (tabla2.Rows[y]["ENVIARFACTURA"].ToString() == "S")
                            {
                                var _factura = tabla2.Rows[y]["FACTURA"].ToString();
                                var _tipodoc = tabla2.Rows[y]["TIPODOC"].ToString();
                                EnviarFactura(_factura, _tipodoc);
                            }
                            if (tabla2.Rows[y]["ENVIARESTADOCUENTA"].ToString() == "S")
                            {
                                var _cliente = tabla2.Rows[y]["CLIENTE"].ToString();
                                var _factura = tabla2.Rows[y]["FACTURA"].ToString();
                                EnviarEstadoCuenta(_cliente, _factura);
                            }
                            if (tabla2.Rows[y]["E_MAIL"].ToString() != null && tabla2.Rows[y]["ENVIARCORREO"].ToString() == "S")
                            {
                                //Enviamos Correo Eletronico
                                var _correo = tabla2.Rows[y]["E_MAIL"].ToString();
                                var _nombre = tabla2.Rows[y]["EMBARCAR_A"].ToString();
                                var _factura = tabla2.Rows[y]["FACTURA"].ToString();
                                var _enviarfactura = tabla2.Rows[y]["ENVIARFACTURA"].ToString();
                                var _enviarestadocuenta = tabla2.Rows[y]["ENVIARESTADOCUENTA"].ToString();
                                var _totalmercaderia = string.Format("{0:#.##}", tabla2.Rows[y]["TOTAL_FACTURA"]);
                                var _correovendedor = tabla2.Rows[y]["CORREOVENDEDOR"].ToString();
                                //var _factura = tabla2.Rows[y]["FACTURA"].ToString();


                                EnviarCorreo(_correo, _nombre, _factura, _totalmercaderia, _enviarfactura, _enviarestadocuenta, _correovendedor);
                            }
                            else
                            {
                                //solo enviarmos datos al vendedor factura y estado de cuentas
                                var _correo = "N";
                                var _nombre = tabla2.Rows[y]["EMBARCAR_A"].ToString();
                                var _factura = tabla2.Rows[y]["FACTURA"].ToString();
                                var _enviarfactura = tabla2.Rows[y]["ENVIARFACTURA"].ToString();
                                var _enviarestadocuenta = tabla2.Rows[y]["ENVIARESTADOCUENTA"].ToString();
                                var _totalmercaderia = string.Format("{0:#.##}", tabla2.Rows[y]["TOTAL_FACTURA"]);
                                var _correovendedor = tabla2.Rows[y]["CORREOVENDEDOR"].ToString();
                                var _tipodoc = tabla2.Rows[y]["TIPODOC"].ToString();
                                if (tabla2.Rows[y]["ENVIARFACTURA"].ToString() == "N" || tabla2.Rows[y]["ENVIARFACTURA"].ToString() == null || tabla2.Rows[y]["ENVIARFACTURA"].ToString() == "")
                                {
                                    EnviarFactura(_factura, _tipodoc);
                                }
                                if (tabla2.Rows[y]["ENVIARESTADOCUENTA"].ToString() == "N" || tabla2.Rows[y]["ENVIARFACTURA"].ToString() == null || tabla2.Rows[y]["ENVIARFACTURA"].ToString() == "")
                                {
                                    var _cliente = tabla2.Rows[y]["CLIENTE"].ToString();
                                    //var _factura = tabla2.Rows[y]["FACTURA"].ToString();

                                    EnviarEstadoCuenta(_cliente, _factura);
                                }
                                EnviarCorreo(_correo, _nombre, _factura, _totalmercaderia, _enviarfactura, _enviarestadocuenta, _correovendedor);
                            }
                            if (tabla2.Rows[y]["ENVIARMENSAJE"].ToString() == "S")
                            {
                                var _factura = tabla2.Rows[y]["FACTURA"].ToString();
                                var _enviarmensaje = tabla2.Rows[y]["ENVIARMENSAJE"].ToString();
                                //////   EnviarMensajeTigo(_factura, _enviarmensaje);
                            }
                            else
                            {
                                var _factura = tabla2.Rows[y]["FACTURA"].ToString();
                                var _enviarmensaje = tabla2.Rows[y]["ENVIARMENSAJE"].ToString();
                                ///// EnviarMensajeTigo(_factura, _enviarmensaje);
                            }
                            var _factura2 = tabla2.Rows[y]["FACTURA"].ToString();


                        }

                    }


                }

            }
            finally
            {
                conn.Close();
            }
        }
        private void EnviarFactura(string _factura,string _tipodoc)
        {
            eventLog1.WriteEntry("Proceso de Envio de Factura");
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["sConexion"].ToString());
            ReportDocument crRpt = new ReportDocument();
            ParameterField paramfied = new ParameterField();
            ParameterFields paramFields = new ParameterFields();
            ParameterDiscreteValue parameterDiscreteValue = new ParameterDiscreteValue();
            try
            {
                if (_tipodoc == "CCF")
                {

                    crRpt.Load(@"C:\CoagroAttCliente\Reportes\CCF.rpt");

                    crRpt.SetParameterValue("@factura", _factura);

                    crRpt.ExportToDisk(ExportFormatType.PortableDocFormat, @"C:\CoagroAttCliente\Envio\" + "DOC_" + _factura + ".pdf");
                    // System.IO.File.Delete(@"C:\Destino\"+factura+".pdf");
                    eventLog1.WriteEntry("Enviamos CCF");
                }
                if (_tipodoc == "FAC")
                {

                    crRpt.Load(@"C:\CoagroAttCliente\Reportes\FAC.rpt");
                    crRpt.SetParameterValue("@factura", _factura);
                    crRpt.ExportToDisk(ExportFormatType.PortableDocFormat, @"C:\CoagroAttCliente\Envio\" + "DOC_" + _factura + ".pdf");
                    // System.IO.File.Delete(@"C:\Destino\"+factura+".pdf");
                    eventLog1.WriteEntry("Enviamos Factura");
                }

            }
            catch (Exception ex)
            {

                eventLog1.WriteEntry(ex.Message);
            }
            
            

            //Actualizamos que se proceso y envio de factura en tabla
            using (var ada = new SqlDataAdapter("UPDATE CINCOH.CoagroAttCliente set EnvioFactura='S',FechaEnvioFactura=@fecha where Factura=@factura ", cn))
            {

                ada.SelectCommand.Parameters.AddWithValue("@factura", _factura);
                ada.SelectCommand.Parameters.AddWithValue("@fecha", DateTime.Now);
                ada.SelectCommand.CommandType = CommandType.Text;
                cn.Open();
                ada.SelectCommand.ExecuteNonQuery();
                cn.Close();
            }
            eventLog1.WriteEntry("Fin de proceso de Envio de Factura");

        }
        private void EnviarEstadoCuenta(string _cliente,string _factura)
        {
            eventLog1.WriteEntry("Proceso de Estado de Cuentas");
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["sConexion"].ToString());
            //   System.IO.File.Delete(@"C:\Destino\ESTADOCUENTA.pdf");
            ReportDocument crRpt = new ReportDocument();
            ParameterField paramfied = new ParameterField();
            ParameterFields paramFields = new ParameterFields();
            ParameterDiscreteValue parameterDiscreteValue = new ParameterDiscreteValue();
            crRpt.Load(@"C:\CoagroAttCliente\Reportes\ESTADOCUENTA.rpt");
            crRpt.SetParameterValue("@cliente", _cliente);
            crRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, @"C:\CoagroAttCliente\Envio\" + "EC_" + _factura + ".pdf");

            using (var ada = new SqlDataAdapter("UPDATE CINCOH.CoagroAttCliente set EnvioEstadocuenta='S',FechaEnvEstadocuenta=@fecha where Factura=@factura ", cn))
            {

                ada.SelectCommand.Parameters.AddWithValue("@factura", _factura);
                ada.SelectCommand.Parameters.AddWithValue("@fecha", DateTime.Now);
                ada.SelectCommand.CommandType = CommandType.Text;
                cn.Open();
                ada.SelectCommand.ExecuteNonQuery();
                cn.Close();
            }
            eventLog1.WriteEntry("Fin de Proceso de Estado de Cuentas");
        }
        private void EnviarCorreo(string _correo,string _nombre,string _factura,string _totalmercaderia,string _enviarfactura,string _enviarestadocuenta,string _correovendedor)
        {
            eventLog1.WriteEntry("Proceso de Envio de Correo");
            string correoexiste = "EXISTE";
            if (blBandera) return;

            //borramos registros 7 dias


            //   System.IO.File.Delete(@"C:\Destino\" + "DOC_" + _factura + ".pdf");

            string Nfactura = _factura;
            System.Net.Mail.MailMessage correo = new System.Net.Mail.MailMessage();
            if (_correo == "N")
            {
                _correovendedor = "uber.carlosrobertovelasquez@gmail.com";
                correo.To.Add(_correovendedor);
                string EstadoCuentas = @"C:\CoagroAttCliente\Envio\" + "EC_" + _factura + ".pdf";
                correo.Attachments.Add(new System.Net.Mail.Attachment(EstadoCuentas));
                string factura = @"C:\CoagroAttCliente\Envio\" + "DOC_" + _factura + ".pdf";
                correo.Attachments.Add(new System.Net.Mail.Attachment(factura));
                correoexiste = "NOEXISTE";

            }
            else
            {
                if (ComprobarFormatoEmail(_correo) == false)
                {
                   // _correovendedor = "uber.carlosrobertovelasquez@gmail.com";
                    correo.To.Add(_correovendedor);
                }
                else
                {
                    //_correovendedor = "uber.carlosrobertovelasquez@gmail.com";
                    //_correo = "carlosrobertovelasquez@gmail.com";
                    correo.To.Add(_correo);
                    correo.CC.Add(_correovendedor);
                }

                //  string EstadoCuentas = @"c:\Destino\" + "EC_" + _factura + ".pdf";
                //  correo.Attachments.Add(new System.Net.Mail.Attachment(EstadoCuentas));
                //  string factura = @"c:\Destino\" + "DOC_" + _factura + ".pdf";
                //  correo.Attachments.Add(new System.Net.Mail.Attachment(factura));

                correoexiste = "EXISTE";
            }
            correo.Subject = "Envio de Proforma " + " " + _factura;
            correo.SubjectEncoding = System.Text.Encoding.UTF8;


            //Enviamos archivos adjuntos
            if (_enviarestadocuenta == "S")
            {
                string EstadoCuentas = @"C:\CoagroAttCliente\Envio\" + "EC_" + _factura + ".pdf";
                correo.Attachments.Add(new System.Net.Mail.Attachment(EstadoCuentas));
            }

            if (_enviarfactura == "S")
            {
                string factura = @"C:\CoagroAttCliente\Envio\" + "DOC_" + _factura + ".pdf";
                correo.Attachments.Add(new System.Net.Mail.Attachment(factura));

            }





            correo.BodyEncoding = System.Text.Encoding.UTF8;
            correo.IsBodyHtml = true;
            string htmlBody;
            // htmlBody = _factura;
            htmlBody = @"<html><body><p>Estimados Señores " + _nombre + " ,</p>" +
                "<p> Reciba un Cordial saludo de parte de Comercial Agropecuaria S.A. de C.V " +
                "<br> Aprovechamos para agradecerle por la compra de nuestros productos agropecuarios.</br> " +
                "<br>El numero de Documento es " + _factura + " , con un monto de $" + _totalmercaderia + ". </br></p>" +
                "<p>Atentamente Coagro S.A. de C.V ,</P><P> Nota :Adjunto al correo podrá encontar Documento y Estado de Cuentas " +
                "</p> <footer><h4>Mensaje automático, por favor no responder. </h4>" +
                "<h4> Nuestras Redes Sociales:</4> " +
                "<a href='https://www.facebook.com/ComercialAgropecuaria'> Nuestro Facebook</a></footer>" +
                "</body></html>";
            correo.Body = htmlBody;
            correo.From = new System.Net.Mail.MailAddress("AtecionCliente@coagro.com");
            System.Net.Mail.SmtpClient clienteCorreo = new System.Net.Mail.SmtpClient();
            clienteCorreo.Credentials = new System.Net.NetworkCredential("informacion.coagro@gmail.com", "Houdelot777$");
            clienteCorreo.Port = 587;
            clienteCorreo.EnableSsl = true;
            clienteCorreo.Host = "smtp.gmail.com";


            try
            {
                blBandera = true;
                eventLog1.WriteEntry("Se inicio proceso de envio de Informacion", EventLogEntryType.Information);
                
                clienteCorreo.Send(correo);
                SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["sConexion"].ToString());
                using (var ada = new SqlDataAdapter("UPDATE CINCOH.CoagroAttCliente set EnvioCorreo='S',FechaEnvCorreo=@fecha,Procesado='S' ,Correo=@correo where Factura=@factura ", cn))
                {

                    ada.SelectCommand.Parameters.AddWithValue("@factura", _factura);
                    ada.SelectCommand.Parameters.AddWithValue("@fecha", DateTime.Now);
                    ada.SelectCommand.Parameters.AddWithValue("@correo", correoexiste);
                    ada.SelectCommand.CommandType = CommandType.Text;
                    cn.Open();
                    ada.SelectCommand.ExecuteNonQuery();
                    cn.Close();
                }




            }
            catch (Exception ex)
            {

                eventLog1.WriteEntry(ex.Message, EventLogEntryType.Error);
            }

            eventLog1.WriteEntry("Se Envio sin problemas el correo", EventLogEntryType.Information);

          



            blBandera = false;
        }
        private void EnviarMensajeTigo(string _factura,string _enviarmensaje)
        {
            //     wsAPISmsCorpSoapClient wsAPISmsCorpSoapClient = new wsAPISmsCorpSoapClient("wsAPISmsCorpSoap");
            APITIGO.wsAPISmsCorp wsAPISmsCorp = new wsAPISmsCorp();
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["sConexion"].ToString());

            if (_enviarmensaje == "S")
            {
                string _factura2 = @"select Cast(Replace(Replace(Replace(left(ltrim(cli.TELEFONO2),9),'-',''),'(',''),')','') As Int) TelCli,left(cli.alias,28) as alias,SUBSTRING (fac.factura,6,12) as factura,CONVERT( VARCHAR ,fac.fecha,103) as fecha,cast(fac.TOTAL_FACTURA as decimal(16,2))  as totalmercaderia,fac.VENDEDOR, ve.NOMBRE,VE.TELEFONO as TelVend from cincoh.factura as fac,cincoh.cliente as cli ,cincoh.vendedor ve where cli.CLIENTE=fac.CLIENTE_ORIGEN and fac.VENDEDOR=ve.VENDEDOR and fac.TIPO_DOCUMENTO='F' and (cli.TELEFONO2 like'7%' or cli.TELEFONO2 like'6%')  and fac.factura='" + _factura + "'";
                cn.Open();
                new SqlDataAdapter(_factura2, cn);
                SqlDataReader sqlDataReader = new SqlCommand(_factura2, cn).ExecuteReader();


                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        string mensaje = string.Concat(new object[]
                        {
                            "Estimado Cliente ",
                            sqlDataReader["alias"],
                            " Coagro le informa que se genero una Fac.No: ",
                            sqlDataReader["factura"],
                            " Monto: ",
                            sqlDataReader["totalmercaderia"],
                            " Emitida El : ",
                            sqlDataReader["fecha"]
                        });
                        string mensaje2 = string.Concat(new object[]
                        {
                            "Vendedor El Cliente ",
                            sqlDataReader["alias"],
                            " Coagro le informa que se genero una Fac.No: ",
                            sqlDataReader["factura"],
                            " Monto: ",
                            sqlDataReader["totalmercaderia"],
                            " Emitida El : ",
                            sqlDataReader["fecha"]
                        });
                        string numero = "503" + sqlDataReader["TelCli"];
                        string numero2 = "503" + sqlDataReader["TelVend"];
                        //string numero = "50373355123";
                        //string numero2 = "50373355123";

                        wsAPISmsCorp.enviarSMS("COAGRO", "c@gr1hp45", numero, mensaje, "Tigo");
                        wsAPISmsCorp.enviarSMS("COAGRO", "c@gr1hp45", numero2, mensaje2, "Tigo");

                     
                        eventLog1.WriteEntry("Se Envio sin problemas el Mensaje de Tigo al Cliente", EventLogEntryType.Information);

                    }

                }
                else
                {
                   
                    string _factura3 = @"select Cast(Replace(Replace(Replace(left(ltrim(cli.TELEFONO2),9),'-',''),'(',''),')','') As Int) TelCli,left(cli.alias,28) as alias,SUBSTRING (fac.factura,6,12) as factura,CONVERT( VARCHAR ,fac.fecha,103) as fecha,cast(fac.TOTAL_FACTURA as decimal(16,2))  as totalmercaderia,fac.VENDEDOR, ve.NOMBRE,VE.TELEFONO as TelVend from cincoh.factura as fac,cincoh.cliente as cli ,cincoh.vendedor ve where cli.CLIENTE=fac.CLIENTE_ORIGEN and fac.VENDEDOR=ve.VENDEDOR and fac.TIPO_DOCUMENTO='F' and  fac.factura='" + _factura + "'";
                    cn.Open();
                    new SqlDataAdapter(_factura3, cn);
                    SqlDataReader sqlDataReader2 = new SqlCommand(_factura3, cn).ExecuteReader();

                    while (sqlDataReader2.Read())
                    {

                        string mensaje2 = string.Concat(new object[]
                        {
                            "Vendedor El Cliente ",
                            sqlDataReader2["alias"],
                            " Coagro le informa que se genero una Fac.No: ",
                            sqlDataReader2["factura"],
                            " Monto: ",
                            sqlDataReader2["totalmercaderia"],
                            " Emitida El : ",
                            sqlDataReader2["fecha"]
                        });
                        //string numero = "503" + sqlDataReader["TelCli"];
                        string numero2 = "503" + sqlDataReader2["TelVend"];
                         //string numero2 = "50373355123";
                        wsAPISmsCorp.enviarSMS("COAGRO", "c@gr1hp45", numero2, mensaje2, "Tigo");
                    }

                    
                    eventLog1.WriteEntry("Se Envio sin problemas el Mensaje de Tigo al Cliente", EventLogEntryType.Information);

                }
            }
            else
            {

                string _factura4 = @"select Cast(Replace(Replace(Replace(left(ltrim(cli.TELEFONO2),9),'-',''),'(',''),')','') As Int) TelCli,left(cli.alias,28) as alias,SUBSTRING (fac.factura,6,12) as factura,CONVERT( VARCHAR ,fac.fecha,103) as fecha,cast(fac.TOTAL_FACTURA as decimal(16,2))  as totalmercaderia,fac.VENDEDOR, ve.NOMBRE,VE.TELEFONO as TelVend from cincoh.factura as fac,cincoh.cliente as cli ,cincoh.vendedor ve where cli.CLIENTE=fac.CLIENTE_ORIGEN and fac.VENDEDOR=ve.VENDEDOR and fac.TIPO_DOCUMENTO='F'  and fac.factura='" + _factura + "'";
                cn.Open();
                new SqlDataAdapter(_factura4, cn);
                SqlDataReader sqlDataReader3 = new SqlCommand(_factura4, cn).ExecuteReader();

                while (sqlDataReader3.Read())
                {

                    string mensaje2 = string.Concat(new object[]
                    {
                    "Vendedor El Cliente ",
                    sqlDataReader3["alias"],
                    " Coagro le informa que se genero una Fac.No: ",
                    sqlDataReader3["factura"],
                    " Monto: ",
                    sqlDataReader3["totalmercaderia"],
                    " Emitida El : ",
                    sqlDataReader3["fecha"]
                    });
                    //string numero = "503" + sqlDataReader["TelCli"];
                    string numero2 = "503" + sqlDataReader3["TelVend"];

                    //string numero2 = "50373355123";

                    // wsAPISmsCorp.enviarSMS("COAGRO", "c@gr1hp45", numero, mensaje, "Tigo");
                    wsAPISmsCorp.enviarSMS("COAGRO", "c@gr1hp45", numero2, mensaje2, "Tigo");

                }
                
                eventLog1.WriteEntry("Se Envio sin problemas el Mensaje de Tigo al Cliente", EventLogEntryType.Information);
                using (var ada = new SqlDataAdapter("UPDATE CINCOH.CoagroAttCliente set EnvioMensajetigo='S',FechaEnvMensajetigo=@fecha where Factura=@factura ", cn))
                {

                    ada.SelectCommand.Parameters.AddWithValue("@factura", _factura);
                    ada.SelectCommand.Parameters.AddWithValue("@fecha", DateTime.Now);
                    ada.SelectCommand.CommandType = CommandType.Text;
                    cn.Open();
                    ada.SelectCommand.ExecuteNonQuery();

                }

                cn.Close();
            }


            
           




          
        }
        private static bool ComprobarFormatoEmail(string sEmailAComprobar)
        {
            String sFormato;
            sFormato = "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*";
            if (Regex.IsMatch(sEmailAComprobar, sFormato))
            {
                if (Regex.Replace(sEmailAComprobar, sFormato, String.Empty).Length == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        private void borrarArchivo()
        {
            string Directorio = @"C:\CoagroAttCliente\Envio";
            int HorasLimite = 1;
            foreach (var item in Directory.GetDirectories(Directorio))
            {
                if (new DirectoryInfo(item).CreationTime.Add(TimeSpan.FromHours(HorasLimite))<DateTime.Now)
                {
                    Directory.Delete(item, true);
                }
            }
        }
        private void EnviarDocVencidosTDias()
        {
            //Revisamos si son las 9:00 am
            //

            DateTime hoy = DateTime.Now;
            string dia = hoy.ToString("dddd");
            int hora = Convert.ToInt32(hoy.ToString("HH"));
            //Console.WriteLine((int)dateValue.DayOfWeek);


            if (hoy.Hour == 10 && hoy.Minute >=0)
            {

                
                string connectionString = ConfigurationManager.ConnectionStrings["sConexion"].ToString();
                SqlConnection conn = new SqlConnection(connectionString);
                string sql = @"SELECT DOC.DOCUMENTO,DOC.CLIENTE,CLI.NOMBRE,CLI.TELEFONO2,CLI.E_MAIL, DOC.FECHA,DOC.FECHA_VENCE, DOC.MONTO,DOC.SALDO ,DOC.VENDEDOR,VEN.E_MAIL,doc.SUBTIPO,
										CASE WHEN (sub.DESCRIPCION like '%Cons Int Propia%') THEN 'FAC' 
                                        WHEN (sub.DESCRIPCION like '%Cont Contribuyente%') THEN 'CCF' 
										 WHEN (sub.DESCRIPCION like 'Cons Int Ticket%') THEN 'TIK' end as TIPODOC
							FROM 
                            CINCOH.DOCUMENTOS_CC DOC,
                            CINCOH.CLIENTE CLI,
                            CINCOH.VENDEDOR VEN,
							CINCOH.SUBTIPO_DOC_CC sub
                            WHERE  
							doc.SUBTIPO=sub.SUBTIPO and
							sub.TIPO='FAC' and
                            DOC.CLIENTE_ORIGEN=CLI.CLIENTE AND
                            DOC.VENDEDOR=VEN.VENDEDOR AND
                            DAY(DOC.FECHA_VENCE)=DAY(GETDATE()) AND 
                            MONTH(DOC.FECHA_VENCE)=MONTH(GETDATE()) AND 
                            YEAR(DOC.FECHA_VENCE)=YEAR(GETDATE()) AND 
                            DOC.TIPO='FAC' AND 
                            DOC.SALDO>0 AND
                            DOC.DOCUMENTO  NOT IN (SELECT Factura FROM CINCOH.CoagroAttCliente WHERE Modulo='FA_VENCIDA' )
                            ORDER BY DOC.CLIENTE";
                SqlCommand command = new SqlCommand(sql, conn);
                conn.Open();
                try
                {
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        //Revisamos datos por cada cliente y enviamos correo y mensaje al cliente y vendedor
                        string _doc = reader[0].ToString();
                        string _cliente = reader[1].ToString();
                        string _nombre = reader[2].ToString();
                        string _telefono2 = reader[3].ToString();
                        string _email = reader[4].ToString();
                        DateTime _fecha = Convert.ToDateTime( reader[5]);
                        DateTime _fechaVence = Convert.ToDateTime(reader[6]);
                        var  _monto = string.Format("{0:#.##}", reader[7]);
                        var _saldo = string.Format("{0:#.##}", reader[8]);
                        
                        string _vendedor = reader[9].ToString();
                        string _emailvendedor = reader[10].ToString();
                        string _tipodoc = reader[12].ToString();
                        //Enviamos Correo Por documentos Vencidos
                        EnviarFactura(_doc, _tipodoc);
                        EnviarEstadoCuenta(_cliente, _doc);
                         GenerarCorreoDocvencidos(_cliente, _doc, _nombre, _email, _fecha, _fechaVence, _monto, _saldo, _emailvendedor, _tipodoc);
                     
                        
                        

                    }

                }
                catch (Exception)
                {

                    throw;
                }
                finally
                {
                    conn.Close();
                }
            }
        }
        private void GenerarCorreoDocvencidos(string _cliente, string _doc, string _nombre, string _email, DateTime _fecha,DateTime _fechaVence, string _monto, string _saldo,string _emailvendedor,string _tipodoc)
        {

            eventLog1.WriteEntry("Proceso de Envio de Correo");
            string correoexiste = "EXISTE";
          

            //borramos registros 7 dias


            //   System.IO.File.Delete(@"C:\Destino\" + "DOC_" + _factura + ".pdf");

            string Nfactura = _doc;
            System.Net.Mail.MailMessage correo = new System.Net.Mail.MailMessage();
            if (_email ==null)
            {
                _emailvendedor = "uber.carlosrobertovelasquez@gmail.com";
                correo.To.Add(_emailvendedor);
                string EstadoCuentas = @"C:\CoagroAttCliente\Envio\" + "EC_" + _doc + ".pdf";
                correo.Attachments.Add(new System.Net.Mail.Attachment(EstadoCuentas));
                string factura = @"C:\CoagroAttCliente\Envio\" + "DOC_" + _doc + ".pdf";
                correo.Attachments.Add(new System.Net.Mail.Attachment(factura));
                correoexiste = "NOEXISTE";

            }
            else
            {
                if (ComprobarFormatoEmail(_email) == false)
                {
                    _emailvendedor = "uber.carlosrobertovelasquez@gmail.com";
                    correo.To.Add(_emailvendedor);
                 //   correo.CC.Add("lh@coagro.com");
                }
                else
                {
                    _emailvendedor = "uber.carlosrobertovelasquez@gmail.com";
                    _email = "carlosrobertovelasquez@gmail.com";
                    correo.To.Add(_email);
                    correo.CC.Add(_emailvendedor);
                }

                  string EstadoCuentas = @"C:\CoagroAttCliente\Envio\" + "EC_" + _doc + ".pdf";
                  correo.Attachments.Add(new System.Net.Mail.Attachment(EstadoCuentas));
                  string factura = @"C:\CoagroAttCliente\Envio\" + "DOC_" + _doc + ".pdf";
                  correo.Attachments.Add(new System.Net.Mail.Attachment(factura));

                correoexiste = "EXISTE";
            }
            correo.Subject = "Envio de Factura " + " " + _doc;
            correo.SubjectEncoding = System.Text.Encoding.UTF8;


            //Enviamos archivos adjuntos
          //      string EstadoCuentas = @"C:\CoagroAttCliente\Envio\" + "EC_" + _doc + ".pdf";
            //    correo.Attachments.Add(new System.Net.Mail.Attachment(EstadoCuentas));
            
                                                         

            correo.BodyEncoding = System.Text.Encoding.UTF8;
            correo.IsBodyHtml = true;
            string htmlBody;
            // htmlBody = _factura;
            htmlBody = @"<html><body><p>Estimados Señores " + _nombre + " ,</p>" +
                "<p> Reciba un Cordial saludo de parte de Comercial Agropecuaria S.A. de C.V " +
                "<br> Le informamos que este dia se vence</br> " +
                "<br>El Documento Con Numero : " + _doc + " , con un Saldo de $" + _saldo + ". </br></p>" +
                "<p>Atentamente Coagro S.A. de C.V ,</P><P> Nota :Adjunto al correo podrá encontar Documento y Estado de Cuentas " +
                "</p> <footer><h4>Este es un Documento generado de forma automatica.Agradecemos no responder a esta direccion de correo </h4>" +
                "<h4>Si ya realizó el pago respectivo, por favor hacer caso omiso.</4> " +
                "<h4> Nuestras Redes Sociales:</4> " +
                "<a href='https://www.facebook.com/ComercialAgropecuaria'>Facebook </a>" +
                "<a href='https://www.coagro.com'> , www.cogaro.com</a>" +
                "</footer>" +
                "</body></html>";
            correo.Body = htmlBody;
            correo.From = new System.Net.Mail.MailAddress("AtecionCliente@coagro.com");
            System.Net.Mail.SmtpClient clienteCorreo = new System.Net.Mail.SmtpClient();
            clienteCorreo.Credentials = new System.Net.NetworkCredential("informacion.coagro@gmail.com", "Houdelot777$");
            clienteCorreo.Port = 587;
            clienteCorreo.EnableSsl = true;
            clienteCorreo.Host = "smtp.gmail.com";


            try
            {
                blBandera = true;
                eventLog1.WriteEntry("Se inicio proceso de envio de Informacion", EventLogEntryType.Information);

                clienteCorreo.Send(correo);
                SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["sConexion"].ToString());
                using (var ada = new SqlDataAdapter("insert into CINCOH.CoagroAttCliente (factura,FechaRegistro,procesado,Modulo) VALUES(@factura,@FechaRegistro,@procesado,@Modulo)", cn))
                {

                    ada.SelectCommand.Parameters.AddWithValue("@factura", _doc);
                    ada.SelectCommand.Parameters.AddWithValue("@FechaRegistro", DateTime.Now);
                    ada.SelectCommand.Parameters.AddWithValue("@procesado", "S");
                    ada.SelectCommand.Parameters.AddWithValue("@Modulo", "FA_VENCIDA");
                    ada.SelectCommand.CommandType = CommandType.Text;
                    cn.Open();
                    ada.SelectCommand.ExecuteNonQuery();
                    cn.Close();
                }




            }
            catch (Exception ex)
            {

                eventLog1.WriteEntry(ex.Message, EventLogEntryType.Error);
            }

            eventLog1.WriteEntry("Se Envio sin problemas el correo", EventLogEntryType.Information);





           
        }
        private void EnviarDocVencidosJueves()
        {
            // Revisamos si es dia jueves y 9.30 am

            if (true)
            {
               
            }
        }
        private void EnviarPagos()
        {
            //Revisamos si se ha insertado un nuevo pago en la tabla CoagroAttCliente que venga de Auxiliar CC y que tenga un tipo Deposito o Recibo

            string connectionString = ConfigurationManager.ConnectionStrings["sConexion"].ToString();
            SqlConnection conn = new SqlConnection(connectionString);
            string sql = @"select aux.FECHA,aux.DEBITO as NumFactura, doc.monto as MontoFactura, aux.CREDITO as NumRecibo,aux.MONTO_CREDITO as MontoRecibo ,doc.SALDO,cli.ALIAS,cli.E_MAIL,cli.TELEFONO2,ven.E_MAIL as correovendedor,ven.telefono as telefonovendedor from 
                            CINCOH.AUXILIAR_CC aux,
                            CINCOH.DOCUMENTOS_CC doc,
                            CINCOH.cliente cli,
                            CINCOH.VENDEDOR ven	 
                            where 
                            doc.CLIENTE_ORIGEN=cli.CLIENTE and
                            aux.DEBITO=doc.DOCUMENTO and
                            doc.VENDEDOR=ven.VENDEDOR and
                            aux.TIPO_CREDITO in ('DEP','REC')  and 
                            aux.CREDITO  in(select recibo from CINCOH.CoagroAttCliente where modulo='NEWPAGO' and Procesado='N')";
            SqlCommand command = new SqlCommand(sql, conn);
            conn.Open();
            try
            {
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    //Revisamos datos por cada cliente y enviamos correo y mensaje al cliente y vendedor
                    string _doc = reader[0].ToString();
                    string _cliente = reader[1].ToString();
                    string _nombre = reader[2].ToString();
                    string _telefono2 = reader[3].ToString();
                    string _email = reader[4].ToString();
                    DateTime _fecha = Convert.ToDateTime(reader[5]);
                    DateTime _fechaVence = Convert.ToDateTime(reader[6]);
                    var _monto = string.Format("{0:#.##}", reader[7]);
                    var _saldo = string.Format("{0:#.##}", reader[8]);

                    string _vendedor = reader[9].ToString();
                    string _emailvendedor = reader[10].ToString();
                    string _tipodoc = reader[12].ToString();
                    //Enviamos Correo Por documentos Vencidos
                    EnviarFactura(_doc, _tipodoc);
                    EnviarEstadoCuenta(_cliente, _doc);
                    GenerarCorreoDocvencidos(_cliente, _doc, _nombre, _email, _fecha, _fechaVence, _monto, _saldo, _emailvendedor, _tipodoc);




                }

            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                conn.Close();
            }


        }

    }
}
