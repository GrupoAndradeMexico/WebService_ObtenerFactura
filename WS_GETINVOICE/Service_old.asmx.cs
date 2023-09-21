using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;
using System.Configuration;
using System.Collections;

using System.Runtime.InteropServices; //DLLImport
using System.Security.Principal; //WindowsImpersonationContext
using System.Security.Permissions; //PermissionSetAttribute 

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WS_GETINVOICE
{
    /// <summary>
    /// Descripción breve de Service1
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente. 
    // [System.Web.Script.Services.ScriptService]
    public class Service1 : System.Web.Services.WebService
    {
       // Byte[] pdf = new Byte[0];
       // Byte[][] arrFacturas = new Byte[0][];
       //20160712 Se agrega código para que en el objeto serializado de retorno agregue un mensaje.
        //si el mensaje es vacio no hubo problema alguno.
        
        //TODO: Agregar la tabla de las series que no se muestran por agencia/marca consultar aquellas que tienen aún un Saldo por pagar y/o aquellas que son de Contado.
        // NO porque solo muestra el pdf y el que tiene valor oficial es el xml.

        //20191101  Huyyyy que miedo--> Cuando se hace un Depósito Referenciado de un cliente perteneciente a una cotizacion de una sucursal, el Anticipo que se crea en automático, se registra en la matríz.
        //   por lo que se tiene la cotizacion en la sucursal y el anticipo en la matriz. Al hacer la búsqueda del documento en las rutas de la sucursal, no se encuentra físicamente ahí y manda error, aunque el pdf se 
        //   encuentran en la ruta de la matriz ("los archivos se van a la matriz") . BPRO muestra (con una vista), registros tanto en la sucursal como en la matriz (en ambas bases de datos), estando únicamente en la matriz el documento.
        //   se agrega al query de consulta del documento, que la serie del documento buscado,  pertenezca a las series de la base de datos donde se está haciendo la consulta.

        string StringConnection = ConfigurationManager.AppSettings["ConnectionString"].ToString(); //"Data Source=192.168.20.59;Initial Catalog=PortalClientes;User ID=sa;Password=S0p0rt3";
        string StringConnectionBP = "Data Source={0};Initial Catalog={1};User ID={2};Password={3}";
        string strCarpetaLocal = ConfigurationManager.AppSettings["strCarpetaLocal"].ToString(); //"C:\\LB\\WS_GETINVOICE\\";

        ConexionBD objDB = null;
        private DiskFileDestinationOptions diskFileDestinationOptions;
        private ExportOptions exportOptions;

        #region Impersonacion en el servidor remoto
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool LogonUser(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

        //[DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        //private unsafe static extern int FormatMessage(int dwFlags, ref IntPtr lpSource, int dwMessageId, int dwLanguageId, ref String lpBuffer, int nSize, IntPtr* arguments);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public extern static bool DuplicateToken(IntPtr existingTokenHandle, int SECURITY_IMPERSONATION_LEVEL, ref IntPtr duplicateTokenHandle);

        // logon types
        const int LOGON32_LOGON_INTERACTIVE = 2;
        const int LOGON32_LOGON_NETWORK = 3;
        const int LOGON32_LOGON_NEW_CREDENTIALS = 9;

        // logon providers
        const int LOGON32_PROVIDER_DEFAULT = 0;
        const int LOGON32_PROVIDER_WINNT50 = 3;
        const int LOGON32_PROVIDER_WINNT40 = 2;
        const int LOGON32_PROVIDER_WINNT35 = 1;

        #region manejo de errores
        // GetErrorMessage formats and returns an error message
        // corresponding to the input errorCode.
        public static string GetErrorMessage(int errorCode)
        {
            int FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100;
            int FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200;
            int FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;

            int messageSize = 255;
            string lpMsgBuf = "";
            int dwFlags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;

            IntPtr ptrlpSource = IntPtr.Zero;
            IntPtr ptrArguments = IntPtr.Zero;

            int retVal = 1; //FormatMessage(dwFlags, ref ptrlpSource, errorCode, 0, ref lpMsgBuf, messageSize, &ptrArguments);
            if (retVal == 0)
            {
                throw new ApplicationException(string.Format("Failed to format message for error code '{0}'.", errorCode));
            }

            return lpMsgBuf;
        }

        private static void RaiseLastError()
        {
            int errorCode = Marshal.GetLastWin32Error();
            string errorMessage = "Error LB"; //GetErrorMessage(errorCode);

            throw new ApplicationException(errorMessage);
        }

        #endregion
        #endregion

        [WebMethod]
        //public Byte[][] BuscaFacturasXVin(string VIN)
        public Documento BuscaFacturasXVin(string VIN)
        { 
           Documento objRegresar = new Documento();                        

           var arrPdfs = new List<byte[]>();       
           ArrayList arrRutasPdf = new ArrayList(); 
           string RutaPdf="";
            string Q = "";
           bool Encontrado=false;

           if (VIN.Trim().Length == 17)
           {
               this.objDB = new ConexionBD(this.StringConnection);
               Q = "select * from TRANSMISION ";
               DataSet dsCon = this.objDB.Consulta(Q);
               foreach (DataRow reg in dsCon.Tables[0].Rows)
               {
                   if (Encontrado == false)
                   {
                       string StringConnectionBPaux = string.Format(this.StringConnectionBP, reg["ip"].ToString(), reg["bd_alterna"].ToString(), reg["usr_bd"].ToString(), reg["pass_bd"].ToString());
                       ConexionBD objDBBP = new ConexionBD(StringConnectionBPaux);
                       Q = " SELECT VDE_DOCTO FROM ADE_VTACFD WHERE VDE_VIN='" + VIN.ToUpper().Trim() + "'";
                       DataSet dsBP = objDBBP.Consulta(Q);
                       if (!objDBBP.EstaVacio(dsBP))
                       {
                           foreach (DataRow regD in dsBP.Tables[0].Rows)
                           {
                               //teniendo los datos vamos por el archivo    
                               //TODO: falta hacer lo mismo de la fecha para que recupere la factura del servidor auxiliar.

                               string RutaRemotaPDF = "\\\\" + reg["ip_almacen_archivos"].ToString() + reg["dir_remoto_pdf"].ToString() + "\\" + regD["vde_docto"].ToString().Trim() + ".pdf";
                               RutaPdf = TraeArchivo(reg["usr_remoto"].ToString(), reg["pass_remoto"].ToString(), reg["ip_almacen_archivos"].ToString(), RutaRemotaPDF);
                               if (RutaPdf.IndexOf("ERROR:") == -1)
                               {
                                   Encontrado = true;
                                   arrRutasPdf.Add(RutaPdf);
                               }
                           } //del for de cada documento.                                   
                       }//Si viene vacio el dataset
                   } //Del si ya fue encontrado 
                   if (Encontrado)
                       break;
               } //Del ForDeCada basededatos

               //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++/
               if (!Encontrado)
               {
                   RutaPdf = "ERROR: No se encontró registro en BD del V.I.N. = " + VIN.ToUpper().Trim();
                   arrRutasPdf.Add(RutaPdf);
               }
           }
           else {
               RutaPdf = "ERROR: La clave V.I.N  : " + VIN.ToUpper().Trim() + " proporcionada es inválida.";
               arrRutasPdf.Add(RutaPdf);
           }// 


            foreach (string Archivo in arrRutasPdf)
            {            
                if (Archivo.IndexOf("ERROR:") > -1)
                { //Quiere decir que hay un error:
                    RutaPdf = GeneraPdfMensaje(Archivo); //Siempre se llama Error.pdf                    
                }
                else
                {
                 RutaPdf = Archivo.Trim();
                }
                arrPdfs.Add(ConviertePDFtoArregloDeBytes(RutaPdf)); 
            }

            objRegresar.arrFacturas = arrPdfs.ToArray();   
            return objRegresar; 
        }

        [WebMethod]
        //public Byte[] MuestraFactura(string RFCEMISOR, string RFCRECEPTOR, string SERIE, string FOLIO)
        public Documento MuestraFactura(string RFCEMISOR, string RFCRECEPTOR, string SERIE, string FOLIO)
        {

          Documento objRegresar = new Documento();
          
          string RutaPdf = "";
          string TABLACONSULTA = "TRANSMISION";

          if (RFCEMISOR.Trim() == "")
              RutaPdf = "ERROR: Proporcione el RFC EMISOR";
          if (SERIE.Trim() == "")
              RutaPdf = "ERROR: Proporcione la SERIE";

          if (RutaPdf == "")
          {
              this.objDB = new ConexionBD(this.StringConnection);

              string Q = "select * from TRANSMISION WHERE id_agencia in (select id_agencia from BP_SERIES_BUSQUEDA where id_serie='" + SERIE.ToUpper().Trim() + "' AND ID_AGENCIA in (select id_agencia from AGENCIAS where rfc='" + RFCEMISOR.ToUpper().Trim() + "'))";
              DataSet ds = this.objDB.Consulta(Q);
              if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
              {
                  foreach (DataRow reg in ds.Tables[0].Rows)
                  {
                      this.StringConnectionBP = string.Format(this.StringConnectionBP, reg["ip"].ToString(), reg["nombre_bd"].ToString(), reg["usr_bd"].ToString(), reg["pass_bd"].ToString());
                      ConexionBD objDBBP = new ConexionBD(this.StringConnectionBP);

                      //consultamos si existe el registro del documento en la base de datos local de Bpro
                      //Q = "select VDE_DOCTO as archivo, VDE_FECHOPE as fecha from ADE_VTACFD";
                      //Q += " where VDE_SERIE = '" + SERIE.Trim() + "'";
                      //Q += " and VDE_FOLIO = '" + FOLIO.Trim() + "'";

                      //20191101 para el problema de que el registro esté en la base de datos donde esta físicamente el documento.
                      Q = "select VDE_DOCTO as archivo, VDE_FECHOPE as fecha from ADE_VTACFD, ADE_CFDFOLIOS";
                      Q += " where VDE_SERIE = '" + SERIE.Trim() + "'";
                      Q += " and VDE_FOLIO = '" + FOLIO.Trim() + "'";
                      Q += " and VDE_SERIE = FCF_SERIE";                      
                      DataSet ds_vde_docto = objDBBP.Consulta(Q);
                      if (!objDBBP.EstaVacio(ds_vde_docto))
                      {
                          foreach (DataRow registro in ds_vde_docto.Tables[0].Rows)
                          {
                              string documento = registro["archivo"].ToString().Trim();
                              string fecha = registro["fecha"].ToString().Trim(); //dd/MM/yyyy
                              fecha = fecha.Substring(6, 4) + fecha.Substring(3, 2) + fecha.Substring(0, 2);

                              Q = "select Convert(char(8),fecha_hasta,112) from TRANSMISION_AUX where id_agencia = " + reg["id_agencia"].ToString().Trim();
                              string fechau_transaux = objDB.ConsultaUnSoloCampo(Q);
                              fechau_transaux = fechau_transaux.Trim() == "" ? "20000101" : fechau_transaux.Trim();
                              //Por la fecha vemos si está en el servidor alterno
                              if (Convert.ToDouble(fecha) <= Convert.ToDouble(fechau_transaux))
                                  TABLACONSULTA = "TRANSMISION_AUX"; //ESTA EN EL SERVIDOR ALTERNO
                              else
                                  TABLACONSULTA = "TRANSMISION"; //ESTA EN EL SERVIDOR ORIGINAL

                              if (TABLACONSULTA != "TRANSMISION_AUX")
                                {
                                    Q = "Select ip,usr_bd,pass_bd,nombre_bd, dir_remoto_xml, dir_remoto_pdf,usr_remoto,pass_remoto, ip_almacen_archivos, smtpserverhost, smtpport, usrcredential, usrpassword ";
                                    Q += " From " + TABLACONSULTA + " where id_agencia='" + reg["id_agencia"].ToString().Trim() + "'";
                                }
                              else {
                                    Q = "Select 'ip' as ip, 'usr_bd' as usr_bd, 'pass_bd' as pass_bd, 'nombre_bd' as nombre_bd, dir_remoto_xml, dir_remoto_pdf,usr_remoto,pass_remoto, ip_almacen_archivos, 'smtpserverhost' as smtpserverhost, 'smtpport' as smtpport, 'usrcredential' as usrcredential , 'usrpassword' as usrpassword";
                                    Q += " From " + TABLACONSULTA + " where id_agencia='" + reg["id_agencia"].ToString().Trim()  + "'";            
                                   }

                            DataSet ds_aux = objDB.Consulta(Q);
                            if (!objDBBP.EstaVacio(ds_aux))
                            {
                                DataRow regConexion = ds_aux.Tables[0].Rows[0];                                
                                //strUsrRemoto = regConexion["usr_remoto"].ToString().Trim();
                                //strPassRemoto = regConexion["pass_remoto"].ToString().Trim();
                                //strDirectorioRemotoXML = regConexion["dir_remoto_xml"].ToString().Trim();
                                //strDirectorioRemotoPDF = regConexion["dir_remoto_pdf"].ToString().Trim();
                                //strIPFileStorage = regConexion["ip_almacen_archivos"].ToString().Trim();}
                                // --- para envio de correo ---
                                //smtpserverhost = regConexion["smtpserverhost"].ToString().Trim();
                                //smtpport = regConexion["smtpport"].ToString().Trim();
                                //usrcredential = regConexion["usrcredential"].ToString().Trim();
                                //usrpassword = regConexion["usrpassword"].ToString().Trim();


                                //teniendo los datos vamos por el archivo    
                                string RutaRemotaPDF = "\\\\" + regConexion["ip_almacen_archivos"].ToString() + regConexion["dir_remoto_pdf"].ToString() + "\\" + documento.Trim() + ".pdf";
                                RutaPdf = TraeArchivo(regConexion["usr_remoto"].ToString(), regConexion["pass_remoto"].ToString(), regConexion["ip_almacen_archivos"].ToString(), RutaRemotaPDF);
                                if (RutaPdf.IndexOf("ERROR:") == -1 && File.Exists(RutaPdf))
                                {//Ya la encontramos ahora validamos si está cancelada
                                    Q = "Select Count(*) FROM ADE_CANCFD cfdscancelados where CDE_SERIE='" + SERIE.Trim() + "' and CDE_FOLIO = '" + FOLIO.Trim() + "'";
                                    if (objDBBP.ConsultaUnSoloCampo(Q).Trim() == "1")
                                    {
                                        PonMarcaDeAgua(RutaPdf, "CANCELADA");
                                    }
                                    else
                                    {
                                        PonMarcaDeAgua(RutaPdf, "SIN VALOR COMERCIAL");
                                    }
                                    break;
                                }
                            } //de que si tenemos datos de la tabla de TRANSFERENCIA 
                         }//del for de cada documento encontrado en la tabla de BPRO de facturas.                      
                        }//de si existe el documento en la tabla de BPRo.
                      else
                      {
                          RutaPdf = "ERROR: No se encontró registro en BD del documento: Serie = " + SERIE + " Folio = " + FOLIO;
                      }                      
                  }//delfor de todas las agencias con el RFC de empresa
                  
              }//de que tenemos datos para conectarnos a BPRo y consultar en la base local
              else
              {
                  RutaPdf = "ERROR: No se encontró agencia con RFC: " + RFCEMISOR + " SERIE BUSQUEDA: " + SERIE;
              }
          }

          if (RutaPdf.IndexOf("ERROR:") > -1)
          { //Quiere decir que hay un error:
              objRegresar.mensajeresultado = "El documento " + SERIE.Trim() + FOLIO.Trim() + " no está disponible para consulta"; //RutaPdf.Trim();
              RutaPdf = GeneraPdfMensaje(RutaPdf); //Siempre se llama Error.pdf              
          }

          if (RutaPdf.Trim() != "" && RutaPdf.IndexOf("ERROR:") == -1  && File.Exists(RutaPdf))
          {
              try
              {
                  FileStream foto = new FileStream(RutaPdf, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                  Byte[] arreglo = new Byte[foto.Length];
                  BinaryReader reader = new BinaryReader(foto);
                  arreglo = reader.ReadBytes(Convert.ToInt32(foto.Length));
                  objRegresar.pdf = arreglo;                  
                  foto.Flush();
                  foto.Close();
              }
              catch (Exception exLB)
              {
                  Debug.WriteLine(exLB.Message);
              }
              finally {
                  FileInfo fires = new FileInfo(RutaPdf);
                  if (fires.Name.ToUpper().IndexOf(".PDF") > -1)
                  {
                      fires.Delete();
                  }
              }
          }
            
        return objRegresar;
        }

        public class Documento
        {
           public  Byte[] pdf {get; set;}
           public  Byte[][] arrFacturas {get; set;} 
           public string mensajeresultado {get;set;}

           public Documento()
           { 
                this.pdf = new Byte[0];
                this.arrFacturas = new Byte[0][];
                this.mensajeresultado = "";
           }
        } //De la clase Documento.



        public string GeneraPdfMensaje(string Mensaje)
        {
            string res = "";

            try
            {
                CrystalDecisions.CrystalReports.Engine.ReportDocument report = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                report.Load(this.strCarpetaLocal  + "\\Mensaje.rpt");
                
                report.SetParameterValue("Mensaje", Mensaje);

                if (!System.IO.Directory.Exists(this.strCarpetaLocal))
                {
                    System.IO.Directory.CreateDirectory(this.strCarpetaLocal);
                }

                diskFileDestinationOptions = new DiskFileDestinationOptions();
                exportOptions = report.ExportOptions;
                exportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                exportOptions.FormatOptions = null;

                exportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;

                //string FechaHora = DateTime.Now.ToString("yyyyMMdd HH:mm:ss");
                //FechaHora = FechaHora.Replace(":", "");
                //FechaHora = FechaHora.Replace(" ", "");
                
                if (File.Exists(this.strCarpetaLocal + "\\Error.pdf"))
                {
                    File.Delete(this.strCarpetaLocal + "\\Error.pdf");
                }

                diskFileDestinationOptions.DiskFileName = this.strCarpetaLocal + "\\Error.pdf";
                exportOptions.DestinationOptions = diskFileDestinationOptions;
                report.Export();
                report.Close();
                res = this.strCarpetaLocal + "\\Error.pdf";
            }
            catch (Exception ex)
            {
                res = "ERROR: " + ex.Message.Trim(); 
            }

            return res;
        }


        public string TraeArchivo(string strUsrRemoto, string strPassRemoto, string strIPFileStorage, string RutaPDF)
        {
            string res = "";

            string strDominio = "";

            if (strUsrRemoto.IndexOf("\\") > -1)
            {   // DANDRADE\sistemas     DANDRADE = dominio sistemas=usuario
                strDominio = strUsrRemoto.Substring(0, strUsrRemoto.IndexOf("\\"));
                strUsrRemoto = strUsrRemoto.Substring(strUsrRemoto.IndexOf("\\") + 1);
            }

            #region funciones de logueo
            IntPtr token = IntPtr.Zero;
            IntPtr dupToken = IntPtr.Zero;
            //primero intentamos el logueo en el servidor remoto
            bool isSuccess = false;
            if (strDominio.Trim() != "") //cuando la impersonizacion es en un servidor que pertenece a un dominio entonces es necesario autenticarse haciendo uso del dominio y no de la ip.
                isSuccess = LogonUser(strUsrRemoto, strDominio, strPassRemoto, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref token);
            else
                isSuccess = LogonUser(strUsrRemoto, strIPFileStorage, strPassRemoto, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref token);

            if (!isSuccess)
            {
                RaiseLastError();
            }

            isSuccess = DuplicateToken(token, 2, ref dupToken);
            if (!isSuccess)
            {
                RaiseLastError();
            }

            WindowsIdentity newIdentity = new WindowsIdentity(dupToken);
            #endregion


            //En este punto ya debemos tener acceso al servidor remoto para poder traer los archivos;
            using (newIdentity.Impersonate())
            {
                try
                {
                        #region Copia del Archivo
                        try
                        {
                            FileInfo fi = new FileInfo(RutaPDF);                            
                            string archdestino = "";
                            archdestino = this.strCarpetaLocal + fi.Name.Trim();                            
                            if (File.Exists(archdestino))
                                File.Delete(archdestino);
                            //nos traemos el archivo                            
                            fi.CopyTo(archdestino);
                            res = archdestino.Trim();                                                                                  
                            //this.Invoke(new UpdateProgessCallback(this.UpdateProgress), new object[] { descargo, TotalArchivosxDescargar });
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                            res = "ERROR:" + ex.Message.Trim();
                        }
                        #endregion                    
                }//del try
                catch (Exception exguia)
                {
                    Debug.WriteLine(exguia.Message);
                    res = "ERROR:" + exguia.Message.Trim();
                }
                finally
                {
                    isSuccess = CloseHandle(token);
                    if (!isSuccess)
                    {
                        RaiseLastError();
                    }
                }
            }//del using del usuario autenticado   

            return res;
        }

        public string ConvierteArregloDeBytesEnArchivo(MemoryStream ms, string RutaPdf)
        {
            string res="";
            using (FileStream file = new FileStream(RutaPdf, FileMode.Create, System.IO.FileAccess.Write))
            {
                byte[] bytes = new byte[ms.Length];
                ms.Read(bytes, 0, (int)ms.Length);
                file.Write(bytes, 0, bytes.Length);
                ms.Close();
            }
            if (File.Exists(RutaPdf))
                res = RutaPdf.Trim();

            return res;
        }

        public byte[] ConviertePDFtoArregloDeBytes(string RutaPdf)
        {
         byte[] res = new byte[0];
         try
              {
                  FileStream foto = new FileStream(RutaPdf, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                  Byte[] arreglo = new Byte[foto.Length];
                  BinaryReader reader = new BinaryReader(foto);
                  arreglo = reader.ReadBytes(Convert.ToInt32(foto.Length));
                  res = arreglo;
                  foto.Flush();
                  foto.Close();
              }
              catch (Exception exLB)
              {
                  Debug.WriteLine(exLB.Message);
              }
              finally {
                  FileInfo fires = new FileInfo(RutaPdf);
                  if (fires.Name.ToUpper().IndexOf(".PDF") > -1)
                  {
                      fires.Delete();
                  }
              }
          return res;
        }
        
        #region Poner Marca De Agua al Pdf.

        public bool PonMarcaDeAgua(string RutaPdf, string TextoMarca)
        {
            bool res = false;
            try
            {
                //ConviertePDFtoArregloDeBytes borra el archivo pdf despues de convertirlo a bytes[]
                Byte[] pdf = AddWatermark(ConviertePDFtoArregloDeBytes(RutaPdf),TextoMarca);                
                File.WriteAllBytes(RutaPdf, pdf);
                if (File.Exists(RutaPdf))
                    res = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            finally {
            }
            return res;
        }


        private static void WriteTextToDocument(BaseFont bf, Rectangle tamPagina, PdfContentByte over, PdfGState gs, string texto)
        {
            over.SetGState(gs);
            over.SetRGBColorFill(220, 220, 220);
            over.SetTextRenderingMode(PdfContentByte.TEXT_RENDER_MODE_STROKE);
            over.SetFontAndSize(bf, 46);
            Single anchoDiag = (Single)Math.Sqrt(Math.Pow((tamPagina.Height - 120), 2) + Math.Pow((tamPagina.Width - 60), 2));
            Single porc = (Single)100 * (anchoDiag / bf.GetWidthPoint(texto, 46));
            over.SetHorizontalScaling(porc);
            double angPage = (-1) * Math.Atan((tamPagina.Height - 60) / (tamPagina.Width - 60));
            over.SetTextMatrix((float)Math.Cos(angPage),(float)Math.Sin(angPage),(float)((-1F) * Math.Sin(angPage)),(float)Math.Cos(angPage),30F,(float)tamPagina.Height - 60);
            over.ShowText(texto);
        }

        //* http://stackoverflow.com/questions/2372041/c-sharp-itextsharp-pdf-creation-with-watermark-on-each-page

        private byte[] AddWatermark(byte[] bytes,string texto)
        {
            BaseFont bf = BaseFont.CreateFont(@"c:\windows\fonts\arial.ttf", BaseFont.CP1252, true);

            using (var ms = new MemoryStream(10 * 1024))
            {
                var reader = new PdfReader(bytes);
                var stamper = new PdfStamper(reader, ms);
                
                    int times = reader.NumberOfPages;
                    for (int i = 1; i <= times; i++)
                    {
                        var dc = stamper.GetOverContent(i);
                        AddWaterMarkaux(dc, texto, bf, 60, 35, new iTextSharp.text.Color(250, 0, 0), reader.GetPageSizeWithRotation(i));
                    }
                    stamper.Close();
                
                reader.Close();

                return ms.ToArray();
            }
        }

        public static void AddWaterMarkaux(PdfContentByte dc, string text, BaseFont font, float fontSize, float angle, iTextSharp.text.Color color, Rectangle realPageSize, Rectangle rect = null)
        {
            var gstate = new PdfGState { FillOpacity = 0.1f, StrokeOpacity = 0.3f };
            dc.SaveState();
            dc.SetGState(gstate);
            dc.SetColorFill(color);
            dc.BeginText();
            dc.SetFontAndSize(font, fontSize);
            var ps = rect ?? realPageSize; /*dc.PdfDocument.PageSize is not always correct*/
            var x = (ps.Right + ps.Left) / 2;
            var y = (ps.Bottom + ps.Top) / 2;
            dc.ShowTextAligned(Element.ALIGN_CENTER, text, x, y, angle);
            dc.EndText();
            dc.RestoreState();
        }

        #endregion
    }
}