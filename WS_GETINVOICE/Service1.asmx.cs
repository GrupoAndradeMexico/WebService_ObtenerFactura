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

using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;

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

        //20201122 Se pide desarrollo que proteja con iText un documento pdf. Se protegen al crearse las series de AutosNuevos y las Series de Autos seminuevos. Este desarrollo impacta la puesta de la marca de agua.
        //puesto que el archivo ya está protegido no se puede alterar (agregar la marca de Agua), hay que desarrollar para abrir el documento protegido sin la marca de agua haciendo uso del password, agregar la marca de agua y volver a proteger.
        //en su defecto, traer la copia desprotegida...
        //Mientras se encuentra solucion, se comentaría la marca de agua.

        //20220726 se agrega && RutaPdf.IndexOf("ERROR:") == -1 para que cuando no encuentre el archivo en la primer ruta siga buscando exaustivamente en las rutas de las demás agencias de la marca.
        //20220804 en esta consulta: string Q = "select * from TRANSMISION WHERE id_agencia in (select id_agencia from BP_SERIES_BUSQUEDA where id_serie='" + SERIE.ToUpper().Trim() + "' AND ID_AGENCIA in (select id_agencia from AGENCIAS where rfc='" + RFCEMISOR.ToUpper().Trim() + "'))";
        // está involucrada la tabla de BP_SERIES_BUSQUEDA en ocasiones esta tabla no tiene todas las series de una agencia y por ese motivo no muestra el documento.
        // se agregan con esta consulta:
        /*
                insert into BP_SERIES_BUSQUEDA(id_agencia,id_serie)   
                select '38', FCF_Serie
                --, PAR_DESCRIP1 
                from [192.168.20.29].GAAA_Tepepan.dbo.ADE_CFDFOLIOS, [192.168.20.29].GAAA_Tepepan.dbo.PNC_PARAMETR
                WHERE PAR_IDENPARA = FCF_IDFOLIOREL
                AND PAR_TIPOPARA = 'FF'
                and FCF_Serie Collate SQL_Latin1_General_CP1_CI_AS not in  (select id_serie from BP_SERIES_BUSQUEDA where id_agencia = '38')
        */
        //20221223 SE solicita que para CDRMBritanica no agregue la marca de agua: "Sin Valor Comercial"
        //20221228 SE agregan metodos para recuperar el xml de la factura y ambos pdf y xml en ObtenerFactura
        //20230412 se registra xml y pdf en la subcarpeta de cada AñoMes este comportamiento es a partir de la fecha configurada en TRANSMISION ! fechasubclasif 
        //20230621 se crea el metodo SISCO para identificar la procedencia del método y si el consumo del web service viene por este método no colocar marca de agua al pdf.

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
          string StringConnectionBP_AUX = this.StringConnectionBP.Trim(); 


          string RutaPdf = "";

          if (RFCEMISOR.Trim() == "")
              RutaPdf = "ERROR: Proporcione el RFC EMISOR";
          if (SERIE.Trim() == "")
              RutaPdf = "ERROR: Proporcione la SERIE";

          if (RutaPdf == "")
          {
              
              try
              {
                  this.objDB = new ConexionBD(this.StringConnection);

                  //string Q = "select * from TRANSMISION WHERE id_agencia in (select id_agencia from BP_SERIES_BUSQUEDA where id_serie='" + SERIE.ToUpper().Trim() + "' AND ID_AGENCIA in (select id_agencia from AGENCIAS where rfc='" + RFCEMISOR.ToUpper().Trim() + "'))";
                  string Q = "select id_agencia from AGENCIAS where rfc='" + RFCEMISOR.ToUpper().Trim() + "'";
                  DataSet ds = this.objDB.Consulta(Q);
                  //if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0) //20221116
                  if (!this.objDB.EstaVacio(ds))
                  {
                      foreach (DataRow reg1 in ds.Tables[0].Rows)
                      {
                          //Q = "Select * from TRANSMISION where id_agencia='" + reg1["id_agencia"].ToString().Trim() + "'";
                          
                          Q = "select ";                          
                          Q += " id_agencia, ip, usr_bd, pass_bd, nombre_bd, ip_almacen_archivos, dir_remoto_xml, dir_remoto_pdf, usr_remoto, pass_remoto, smtpserverhost, smtpport, usrcredential, usrpassword, visible, bd_alterna, plantillaHTML, logo, cuenta_from, enable_ssl, isnull(Convert(char(8),fechasubclasif,112),'19000101') as fechasubclasif ";
                          Q += " from TRANSMISION WHERE  id_agencia='" + reg1["id_agencia"].ToString().Trim() + "'";

                          DataSet ds1 = this.objDB.Consulta(Q);
                          if (!this.objDB.EstaVacio(ds1))
                          {
                              DataRow reg = ds1.Tables[0].Rows[0];
                              if (reg != null)
                              {
                                  StringConnectionBP_AUX = this.StringConnectionBP.Trim();
                                  StringConnectionBP_AUX = string.Format(StringConnectionBP_AUX, reg["ip"].ToString(), reg["nombre_bd"].ToString(), reg["usr_bd"].ToString(), reg["pass_bd"].ToString());
                                  ConexionBD objDBBP = new ConexionBD(StringConnectionBP_AUX);

                                  string id_agencia = reg["id_agencia"].ToString().Trim();

                                  //consultamos si existe el registro del documento en la base de datos local de Bpro
                                  Q = "select VDE_DOCTO as archivo, VDE_FECHOPE as fecha, Convert(char(8),Convert(datetime,VDE_FECHOPE),112) as fechaopeu from ADE_VTACFD";
                                  Q += " where VDE_SERIE = '" + SERIE.Trim() + "'";
                                  Q += " and VDE_FOLIO = '" + FOLIO.Trim() + "'";

                                  //20191101 para el problema de que el registro esté en la base de datos donde esta físicamente el documento.
                                  //Q = "select VDE_DOCTO as archivo, VDE_FECHOPE as fecha from ADE_VTACFD, ADE_CFDFOLIOS";
                                  //Q += " where VDE_SERIE = '" + SERIE.Trim() + "'";
                                  //Q += " and VDE_FOLIO = '" + FOLIO.Trim() + "'";
                                  //Q += " and VDE_SERIE = FCF_SERIE";

                                  string vde_docto = "";
                                  string VDE_FECHOPE = "";
                                  string fechaopeu = "";

                                  //inicia
                                  DataSet ds_vde_docto = objDBBP.Consulta(Q);
                                  if (!objDBBP.EstaVacio(ds_vde_docto))
                                  {
                                      DataRow regFactura = ds_vde_docto.Tables[0].Rows[0];
                                      vde_docto = regFactura["archivo"].ToString().Trim();
                                      VDE_FECHOPE = regFactura["fecha"].ToString().Trim();
                                      fechaopeu = regFactura["fechaopeu"].ToString().Trim();
                                  }

                                  string strUsrRemoto = reg["usr_remoto"].ToString().Trim();
                                  string strPassRemoto = reg["pass_remoto"].ToString().Trim();
                                  string strDirectorioRemotoXML = reg["dir_remoto_xml"].ToString().Trim();
                                  string strDirectorioRemotoPDF = reg["dir_remoto_pdf"].ToString().Trim();
                                  string strIPFileStorage = reg["ip_almacen_archivos"].ToString().Trim();
                                  string fechasubclasif = reg["fechasubclasif"].ToString().Trim();

                                  Q = "select Convert(char(8),fecha_hasta,112) from TRANSMISION_AUX where id_agencia = " + id_agencia;

                                  string fechau_transaux = objDB.ConsultaUnSoloCampo(Q);
                                  fechau_transaux = fechau_transaux.Trim() == "" ? "20000101" : fechau_transaux.Trim();
                                  //Por la fecha vemos si está en el servidor alterno

                                  if (Convert.ToDouble(fechaopeu) <= Convert.ToDouble(fechau_transaux))  //20150101,20160101 
                                  {   //esta en el servidor alterno este maneja rangos de fechas.
                                      Q = "Select ip_almacen_archivos,dir_remoto_xml,dir_remoto_pdf,usr_remoto,pass_remoto From TRANSMISION_AUX where id_agencia='" + id_agencia + "'";
                                      DataSet dsaux = objDB.Consulta(Q);
                                      if (dsaux != null && dsaux.Tables.Count > 0 && dsaux.Tables[0].Rows.Count > 0)
                                      {
                                          DataRow regConexion = dsaux.Tables[0].Rows[0];
                                          strUsrRemoto = regConexion["usr_remoto"].ToString().Trim();
                                          strPassRemoto = regConexion["pass_remoto"].ToString().Trim();
                                          strDirectorioRemotoXML = regConexion["dir_remoto_xml"].ToString().Trim();
                                          strDirectorioRemotoPDF = regConexion["dir_remoto_pdf"].ToString().Trim();
                                          strIPFileStorage = regConexion["ip_almacen_archivos"].ToString().Trim();
                                      }
                                  }
                                  //else
                                  //{ //esta en el servidor original
                                  //Q = "Select ip_almacen_archivos,dir_remoto_xml,dir_remoto_pdf,usr_remoto,pass_remoto From TRANSMISION where id_agencia='" + id_agencia + "'";
                                  //}                           
                                  //termina
                                  //string vde_docto = objDBBP.ConsultaUnSoloCampo(Q).Trim();

                                  /*
                                   * Buscamos en la bitacora si el archivo esta protegido
                                   */
                                  var passFile = this.objDB.getPasswordPdf(vde_docto); //objDBBP.getPasswordPdf(vde_docto); //20221116

                                  if (vde_docto.Trim() != "")
                                  {
                                      //teniendo los datos vamos por el archivo    
                                      string RutaRemotaPDF = "\\\\" + strIPFileStorage.Trim() + "\\" + strDirectorioRemotoPDF.Trim() + "\\" + vde_docto.Trim() + ".pdf";
                                      
                                      if (Convert.ToDouble(fechaopeu) >= Convert.ToDouble(fechasubclasif) && Convert.ToDouble(fechasubclasif) != 19000101) //yyyyMMdd
                                      { //20230412 se registra en la subcarpeta de cada AñoMes 
                                          string AnioFactura = fechaopeu.Substring(0, 4);
                                          string MesFactura = fechaopeu.Substring(4, 2);
                                          RutaRemotaPDF = string.Format("\\\\{0}\\{1}\\{3}{4}\\{2}.pdf", strIPFileStorage, strDirectorioRemotoPDF, vde_docto.Trim(), AnioFactura, MesFactura);
                                      }

                                      RutaPdf = TraeArchivo(strUsrRemoto.Trim(), strPassRemoto.Trim(), strIPFileStorage.Trim(), RutaRemotaPDF);

                                      //20220726 if (passFile != "0" && passFile != "") se agrega && RutaPdf.IndexOf("ERROR:") == -1 para que cuando no encuentre el archivo en la primer ruta siga buscando exaustivamente en las rutas de las demás agencias de la marca.
                                      if (passFile != "0" && passFile != "" && RutaPdf.IndexOf("ERROR:") == -1)
                                      {

                                          PdfSharp.Pdf.PdfDocument document = PdfSharp.Pdf.IO.PdfReader.Open(RutaPdf, passFile, PdfDocumentOpenMode.Modify, null);
                                          bool hasOwnerAccess = document.SecuritySettings.HasOwnerPermissions;
                                          document.Save(RutaPdf);

                                      }

                                      if (RutaPdf.IndexOf("ERROR:") == -1 && File.Exists(RutaPdf))
                                      {//Ya la encontramos ahora validamos si está cancelada
                                          Q = "Select Count(*) FROM ADE_CANCFD cfdscancelados where CDE_SERIE='" + SERIE.Trim() + "' and CDE_FOLIO = '" + FOLIO.Trim() + "'";
                                          if (objDBBP.ConsultaUnSoloCampo(Q).Trim() == "1")
                                          {
                                              //20201122 
                                              PonMarcaDeAgua(RutaPdf, "CANCELADA");
                                              objRegresar.mensajeresultado = "El documento " + SERIE.Trim() + FOLIO.Trim() + " está cancelado ";
                                          }
                                          else
                                          {
                                              //20221223 
                                              if (RFCEMISOR.ToUpper().Trim() != "CBR080923A2A" && RFCEMISOR.ToUpper().Trim() != "FGA161114294" && RFCRECEPTOR.ToUpper().Trim() != "SISCO")
                                                    PonMarcaDeAgua(RutaPdf, "SIN VALOR COMERCIAL");
                                          }

                                          /*colocamos nuevamnete la seguridad*/
                                          if (passFile != "0" && passFile != "")
                                          {
                                              PdfSharp.Pdf.PdfDocument document = PdfSharp.Pdf.IO.PdfReader.Open(RutaPdf);
                                              PdfSecuritySettings securitySettings = document.SecuritySettings;
                                              securitySettings.OwnerPassword = passFile;
                                              securitySettings.PermitAccessibilityExtractContent = false;
                                              securitySettings.PermitAnnotations = false;
                                              securitySettings.PermitAssembleDocument = false;
                                              securitySettings.PermitExtractContent = false;
                                              securitySettings.PermitFormsFill = false;
                                              securitySettings.PermitFullQualityPrint = false;
                                              securitySettings.PermitModifyDocument = false;
                                              securitySettings.PermitPrint = true;
                                              document.Save(RutaPdf);
                                          }

                                          break;
                                      }

                                  }//de si existe el documento en la tabla de BPRo.
                                  else
                                  {
                                      RutaPdf = "ERROR: No se encontró registro en BD del documento: Serie = " + SERIE + " Folio = " + FOLIO;
                                  }
                              } //del for de cada registro de TRANSMISION
                          } //del if de la consulta sobre TRANSMISION si viene vacio.
                          else {
                              RutaPdf = "ERROR: No se encontró registro en TRANSMISION para la agencia: " + reg1["id_agencia"].ToString().Trim();
                          }
                      }//delfor de todas las agencias con el RFC de empresa
                  }//de que tenemos datos para conectarnos a BPRo y consultar en la base local
                  else
                  {
                      RutaPdf = "ERROR: No se encontró agencia con RFC: " + RFCEMISOR + " SERIE BUSQUEDA: " + SERIE;
                  }
              }
              catch (Exception exlb)
              {
                  RutaPdf = "ERROR: " + exlb.Message; 
              }
          }

          if (RutaPdf.IndexOf("ERROR:") > -1)
          { //Quiere decir que hay un error:
              objRegresar.mensajeresultado = "El documento " + SERIE.Trim() + FOLIO.Trim() + " no está disponible para consulta "  + RutaPdf.Trim();
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
           public  Byte[] xml { get; set; }
           public  Byte[][] arrFacturas {get; set;} 
           public string mensajeresultado {get;set;}

           public Documento()
           { 
                this.pdf = new Byte[0];
                this.xml = new Byte[0];
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
        #region PonMarcaDeAgua1

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
#endregion


        private byte[] AddWatermark(byte[] bytes,string texto)
        {
            BaseFont bf = BaseFont.CreateFont(@"c:\windows\fonts\arial.ttf", BaseFont.CP1252, true);

            using (var ms = new MemoryStream(10 * 1024))
            {
                var reader = new iTextSharp.text.pdf.PdfReader(bytes);
                //20201122 si se quiere seguir colocando la marca de agua, hay que invesigar la manera de abrir con el password la factura                 
                //que ya viene protegida.  https://blog.pdfsam.org/pdf-merge/if-you-have-the-error-message-pdfreader-not-opened-with-owner-password/323/
                //SELECT * FROM Tramites.dbo.seguridadPDF p where nombreArchivo like '%AG%7436%'

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



        [WebMethod]
        public Documento MuestraXML(string RFCEMISOR, string RFCRECEPTOR, string SERIE, string FOLIO)
        {
            Documento objRegresar = new Documento();
            string StringConnectionBP_AUX = StringConnectionBP.Trim();
            string RutaXml = "";
            if (RFCEMISOR.Trim() == "")
            {
                RutaXml = "ERROR: Proporcione el RFC EMISOR";
            }
            if (SERIE.Trim() == "")
            {
                RutaXml = "ERROR: Proporcione la SERIE";
            }
            if (RutaXml == "")
            {
                objDB = new ConexionBD(StringConnection);
                try
                {
                    //string Q = "select * from TRANSMISION WHERE id_agencia in (select id_agencia from BP_SERIES_BUSQUEDA where id_serie='" + SERIE.ToUpper().Trim() + "' AND ID_AGENCIA in (select id_agencia from AGENCIAS where rfc='" + RFCEMISOR.ToUpper().Trim() + "'))";
                    string Q = "select ";
                    Q += " id_agencia, ip, usr_bd, pass_bd, nombre_bd, ip_almacen_archivos, dir_remoto_xml, dir_remoto_pdf, usr_remoto, pass_remoto, smtpserverhost, smtpport, usrcredential, usrpassword, visible, bd_alterna, plantillaHTML, logo, cuenta_from, enable_ssl, isnull(Convert(char(8),fechasubclasif,112),'19000101') as fechasubclasif ";
                    Q += " from TRANSMISION WHERE id_agencia in (select id_agencia from BP_SERIES_BUSQUEDA where id_serie='" + SERIE.ToUpper().Trim() + "' AND ID_AGENCIA in (select id_agencia from AGENCIAS where rfc='" + RFCEMISOR.ToUpper().Trim() + "'))";
                    
                    DataSet ds = objDB.Consulta(Q);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow reg in ds.Tables[0].Rows)
                        {
                            StringConnectionBP_AUX = StringConnectionBP.Trim();
                            StringConnectionBP_AUX = string.Format(StringConnectionBP_AUX, reg["ip"].ToString(), reg["nombre_bd"].ToString(), reg["usr_bd"].ToString(), reg["pass_bd"].ToString());
                            ConexionBD objDBBP = new ConexionBD(StringConnectionBP_AUX);
                            string id_agencia = reg["id_agencia"].ToString().Trim();
                            string fechasubclasif = reg["fechasubclasif"].ToString().Trim();

                            Q = "select VDE_DOCTO as archivo, VDE_FECHOPE as fecha, Convert(char(8),Convert(datetime,VDE_FECHOPE),112) as fechaopeu from ADE_VTACFD";
                            Q = Q + " where VDE_SERIE = '" + SERIE.Trim() + "'";
                            Q = Q + " and VDE_FOLIO = '" + FOLIO.Trim() + "'";
                            string vde_docto = "";
                            string VDE_FECHOPE = "";
                            string fechaopeu = "";
                            
                            DataSet ds_vde_docto = objDBBP.Consulta(Q);
                            if (!objDBBP.EstaVacio(ds_vde_docto))
                            {
                                DataRow regFactura = ds_vde_docto.Tables[0].Rows[0];
                                vde_docto = regFactura["archivo"].ToString().Trim();
                                VDE_FECHOPE = regFactura["fecha"].ToString().Trim();
                                fechaopeu = regFactura["fechaopeu"].ToString().Trim();
                            }
                            string strUsrRemoto = reg["usr_remoto"].ToString().Trim();
                            string strPassRemoto = reg["pass_remoto"].ToString().Trim();
                            string strDirectorioRemotoXML = reg["dir_remoto_xml"].ToString().Trim();
                            string strDirectorioRemotoPDF = reg["dir_remoto_pdf"].ToString().Trim();
                            string strIPFileStorage = reg["ip_almacen_archivos"].ToString().Trim();
                            Q = "select Convert(char(8),fecha_hasta,112) from TRANSMISION_AUX where id_agencia = " + id_agencia;
                            string fechau_transaux = objDB.ConsultaUnSoloCampo(Q);
                            fechau_transaux = ((fechau_transaux.Trim() == "") ? "20000101" : fechau_transaux.Trim());
                            if (Convert.ToDouble(fechaopeu) <= Convert.ToDouble(fechau_transaux))
                            {
                                Q = "Select ip_almacen_archivos,dir_remoto_xml,dir_remoto_pdf,usr_remoto,pass_remoto From TRANSMISION_AUX where id_agencia='" + id_agencia + "'";
                                DataSet dsaux = objDB.Consulta(Q);
                                if (dsaux != null && dsaux.Tables.Count > 0 && dsaux.Tables[0].Rows.Count > 0)
                                {
                                    DataRow regConexion = dsaux.Tables[0].Rows[0];
                                    strUsrRemoto = regConexion["usr_remoto"].ToString().Trim();
                                    strPassRemoto = regConexion["pass_remoto"].ToString().Trim();
                                    strDirectorioRemotoXML = regConexion["dir_remoto_xml"].ToString().Trim();
                                    strDirectorioRemotoPDF = regConexion["dir_remoto_pdf"].ToString().Trim();
                                    strIPFileStorage = regConexion["ip_almacen_archivos"].ToString().Trim();
                                }
                            }
                            string passFile = objDBBP.getPasswordPdf(vde_docto);
                            if (vde_docto.Trim() != "")
                            {
                                string RutaRemotaXML = "\\\\" + strIPFileStorage.Trim() + "\\" + strDirectorioRemotoXML.Trim() + "\\" + vde_docto.Trim() + ".xml";
                                
                                if (Convert.ToDouble(fechasubclasif) >= Convert.ToDouble(fechaopeu)) //yyyyMMdd
                                { //20230412 se registra en la subcarpeta de cada AñoMes 
                                    string AnioFactura = fechaopeu.Substring(1, 4);
                                    string MesFactura = fechaopeu.Substring(4, 2);
                                    RutaRemotaXML = string.Format("\\\\{0}\\{1}\\{3}{4}\\{2}.pdf", strIPFileStorage, strDirectorioRemotoXML, vde_docto.Trim(), AnioFactura, MesFactura);
                                }

                                RutaXml = TraeArchivo(strUsrRemoto.Trim(), strPassRemoto.Trim(), strIPFileStorage.Trim(), RutaRemotaXML);
                                break;
                            }
                            RutaXml = "ERROR: No se encontró registro en BD del documento: Serie = " + SERIE + " Folio = " + FOLIO;
                        }
                    }
                    else
                    {
                        RutaXml = "ERROR: No se encontró agencia con RFC: " + RFCEMISOR + " SERIE BUSQUEDA: " + SERIE;
                    }
                }
                catch (Exception exlb)
                {
                    RutaXml = "ERROR: " + exlb.Message;
                }
            }
            if (RutaXml.IndexOf("ERROR:") > -1)
            {
                objRegresar.mensajeresultado = "El documento " + SERIE.Trim() + FOLIO.Trim() + " no está disponible para consulta";
                RutaXml = GeneraPdfMensaje(RutaXml);
            }
            if (RutaXml.Trim() != "" && RutaXml.IndexOf("ERROR:") == -1 && File.Exists(RutaXml))
            {
                try
                {
                    FileStream foto = new FileStream(RutaXml, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    byte[] arreglo = new byte[foto.Length];
                    BinaryReader reader = new BinaryReader(foto);
                    arreglo = (objRegresar.xml = reader.ReadBytes(Convert.ToInt32(foto.Length)));
                    foto.Flush();
                    foto.Close();
                }
                catch (Exception exLBXml)
                {
                    Debug.WriteLine(exLBXml.Message);
                }
                finally
                {
                    FileInfo fires = new FileInfo(RutaXml);
                    if (fires.Name.ToUpper().IndexOf(".XML") > -1)
                    {
                        fires.Delete();
                    }
                }
            }
            return objRegresar;
        }

        [WebMethod]
        public Documento ObtenerFacturaSISCO(string RFCEMISOR, string RFCRECEPTOR, string SERIE, string FOLIO)
        {
            return ObtenerFactura(RFCEMISOR, "SISCO", SERIE, FOLIO);
        }
        [WebMethod]
        public Documento MuestraFacturaSISCO(string RFCEMISOR, string RFCRECEPTOR, string SERIE, string FOLIO)
        {
            return MuestraFactura(RFCEMISOR, "SISCO", SERIE, FOLIO);
        }

        [WebMethod]
        public Documento ObtenerFactura(string RFCEMISOR, string RFCRECEPTOR, string SERIE, string FOLIO)
        {
            Documento objRegresar = new Documento();
            string StringConnectionBP_AUX = StringConnectionBP.Trim();
            string RutaPdf = "";
            string RutaXml = "";
            if (RFCEMISOR.Trim() == "")
            {
                RutaPdf = "ERROR: Proporcione el RFC EMISOR";
            }
            if (SERIE.Trim() == "")
            {
                RutaPdf = "ERROR: Proporcione la SERIE";
            }
            if (RutaPdf == "")
            {
                objDB = new ConexionBD(StringConnection);
                try
                {
                    string Q = "select ";
                    Q += " id_agencia, ip, usr_bd, pass_bd, nombre_bd, ip_almacen_archivos, dir_remoto_xml, dir_remoto_pdf, usr_remoto, pass_remoto, smtpserverhost, smtpport, usrcredential, usrpassword, visible, bd_alterna, plantillaHTML, logo, cuenta_from, enable_ssl, isnull(Convert(char(8),fechasubclasif,112),'19000101') as fechasubclasif ";
                    Q += " from TRANSMISION WHERE id_agencia in (select id_agencia from BP_SERIES_BUSQUEDA where id_serie='" + SERIE.ToUpper().Trim() + "' AND ID_AGENCIA in (select id_agencia from AGENCIAS where rfc='" + RFCEMISOR.ToUpper().Trim() + "'))";
                    DataSet ds = objDB.Consulta(Q);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow reg in ds.Tables[0].Rows)
                        {
                            StringConnectionBP_AUX = StringConnectionBP.Trim();
                            StringConnectionBP_AUX = string.Format(StringConnectionBP_AUX, reg["ip"].ToString(), reg["nombre_bd"].ToString(), reg["usr_bd"].ToString(), reg["pass_bd"].ToString());
                            ConexionBD objDBBP = new ConexionBD(StringConnectionBP_AUX);
                            string id_agencia = reg["id_agencia"].ToString().Trim();
                            Q = "select VDE_DOCTO as archivo, VDE_FECHOPE as fecha, Convert(char(8),Convert(datetime,VDE_FECHOPE),112) as fechaopeu from ADE_VTACFD";
                            Q = Q + " where VDE_SERIE = '" + SERIE.Trim() + "'";
                            Q = Q + " and VDE_FOLIO = '" + FOLIO.Trim() + "'";
                            string vde_docto = "";
                            string VDE_FECHOPE = "";
                            string fechaopeu = "";

                            DataSet ds_vde_docto = objDBBP.Consulta(Q);
                            if (!objDBBP.EstaVacio(ds_vde_docto))
                            {
                                DataRow regFactura = ds_vde_docto.Tables[0].Rows[0];
                                vde_docto = regFactura["archivo"].ToString().Trim();
                                VDE_FECHOPE = regFactura["fecha"].ToString().Trim();
                                fechaopeu = regFactura["fechaopeu"].ToString().Trim();
                            }
                            string strUsrRemoto = reg["usr_remoto"].ToString().Trim();
                            string strPassRemoto = reg["pass_remoto"].ToString().Trim();
                            string strDirectorioRemotoXML = reg["dir_remoto_xml"].ToString().Trim();
                            string strDirectorioRemotoPDF = reg["dir_remoto_pdf"].ToString().Trim();
                            string strIPFileStorage = reg["ip_almacen_archivos"].ToString().Trim();
                            string fechasubclasif = reg["fechasubclasif"].ToString().Trim();


                            Q = "select Convert(char(8),fecha_hasta,112) from TRANSMISION_AUX where id_agencia = " + id_agencia;
                            string fechau_transaux = objDB.ConsultaUnSoloCampo(Q);
                            fechau_transaux = ((fechau_transaux.Trim() == "") ? "20000101" : fechau_transaux.Trim());
                            if (Convert.ToDouble(fechaopeu) <= Convert.ToDouble(fechau_transaux))
                            {
                                Q = "Select ip_almacen_archivos,dir_remoto_xml,dir_remoto_pdf,usr_remoto,pass_remoto From TRANSMISION_AUX where id_agencia='" + id_agencia + "'";
                                DataSet dsaux = objDB.Consulta(Q);
                                if (dsaux != null && dsaux.Tables.Count > 0 && dsaux.Tables[0].Rows.Count > 0)
                                {
                                    DataRow regConexion = dsaux.Tables[0].Rows[0];
                                    strUsrRemoto = regConexion["usr_remoto"].ToString().Trim();
                                    strPassRemoto = regConexion["pass_remoto"].ToString().Trim();
                                    strDirectorioRemotoXML = regConexion["dir_remoto_xml"].ToString().Trim();
                                    strDirectorioRemotoPDF = regConexion["dir_remoto_pdf"].ToString().Trim();
                                    strIPFileStorage = regConexion["ip_almacen_archivos"].ToString().Trim();
                                }
                            }
                            string passFile = "0";
                            if (vde_docto.Trim() != "")
                            {
                                string RutaRemotaPDF = "\\\\" + strIPFileStorage.Trim() + "\\" + strDirectorioRemotoPDF.Trim() + "\\" + vde_docto.Trim() + ".pdf";

                                if (Convert.ToDouble(fechaopeu) >= Convert.ToDouble(fechasubclasif) && Convert.ToDouble(fechasubclasif) != 19000101) //yyyyMMdd
                                { //20230412 se registra en la subcarpeta de cada AñoMes 
                                    string AnioFactura = fechaopeu.Substring(0, 4);
                                    string MesFactura = fechaopeu.Substring(4, 2);
                                    RutaRemotaPDF = string.Format("\\\\{0}\\{1}\\{3}{4}\\{2}.pdf", strIPFileStorage, strDirectorioRemotoPDF, vde_docto.Trim(), AnioFactura, MesFactura);
                                }

                                RutaPdf = TraeArchivo(strUsrRemoto.Trim(), strPassRemoto.Trim(), strIPFileStorage.Trim(), RutaRemotaPDF);
                                if (passFile != "0" && passFile != "")
                                {
                                    PdfSharp.Pdf.PdfDocument document = PdfSharp.Pdf.IO.PdfReader.Open(RutaPdf, passFile, PdfDocumentOpenMode.Modify, null);
                                    bool hasOwnerAccess = document.SecuritySettings.HasOwnerPermissions;
                                    document.Save(RutaPdf);
                                }
                                if (RutaPdf.IndexOf("ERROR:") == -1 && File.Exists(RutaPdf))
                                {
                                    Q = "Select Count(*) FROM ADE_CANCFD cfdscancelados where CDE_SERIE='" + SERIE.Trim() + "' and CDE_FOLIO = '" + FOLIO.Trim() + "'";
                                    if (objDBBP.ConsultaUnSoloCampo(Q).Trim() == "1")
                                    {
                                        PonMarcaDeAgua(RutaPdf, "CANCELADA");
                                        objRegresar.mensajeresultado = "El documento " + SERIE.Trim() + FOLIO.Trim() + " está cancelado ";
                                    }
                                    else
                                    {
                                        //20221223 
                                        if (RFCEMISOR.ToUpper().Trim() != "CBR080923A2A" && RFCEMISOR.ToUpper().Trim() != "FGA161114294" && RFCRECEPTOR.ToUpper().Trim() != "SISCO")
                                            PonMarcaDeAgua(RutaPdf, "SIN VALOR COMERCIAL");
                                    }

                                    if (passFile != "0" && passFile != "")
                                    {
                                        PdfSharp.Pdf.PdfDocument document = PdfSharp.Pdf.IO.PdfReader.Open(RutaPdf);
                                        PdfSecuritySettings securitySettings = document.SecuritySettings;
                                        securitySettings.OwnerPassword = passFile;
                                        securitySettings.PermitAccessibilityExtractContent = false;
                                        securitySettings.PermitAnnotations = false;
                                        securitySettings.PermitAssembleDocument = false;
                                        securitySettings.PermitExtractContent = false;
                                        securitySettings.PermitFormsFill = false;
                                        securitySettings.PermitFullQualityPrint = false;
                                        securitySettings.PermitModifyDocument = false;
                                        securitySettings.PermitPrint = true;
                                        document.Save(RutaPdf);
                                    }
                                    //break;
                                }

                                string RutaRemotaXML = "\\\\" + strIPFileStorage.Trim() + "\\" + strDirectorioRemotoXML.Trim() + "\\" + vde_docto.Trim() + ".xml";

                                if (Convert.ToDouble(fechaopeu) >= Convert.ToDouble(fechasubclasif) && Convert.ToDouble(fechasubclasif) != 19000101)
                                { //20230412 se registra en la subcarpeta de cada AñoMes 
                                    string AnioFactura = fechaopeu.Substring(0, 4);
                                    string MesFactura = fechaopeu.Substring(4, 2);
                                    RutaRemotaXML = string.Format("\\\\{0}\\{1}\\{3}{4}\\{2}.xml", strIPFileStorage, strDirectorioRemotoXML, vde_docto.Trim(), AnioFactura, MesFactura);
                                }                                
                                
                                RutaXml = TraeArchivo(strUsrRemoto.Trim(), strPassRemoto.Trim(), strIPFileStorage.Trim(), RutaRemotaXML);
                                break;
                            }
                            else
                            {
                                RutaPdf = "ERROR: No se encontró registro en BD del documento: Serie = " + SERIE + " Folio = " + FOLIO;
                            }
                        }
                    }
                    else
                    {
                        RutaPdf = "ERROR: No se encontró agencia con RFC: " + RFCEMISOR + " SERIE BUSQUEDA: " + SERIE;
                    }
                }
                catch (Exception exlb)
                {
                    RutaPdf = "ERROR: " + exlb.Message;
                }
            }
            if (RutaPdf.IndexOf("ERROR:") > -1)
            {
                objRegresar.mensajeresultado = "El documento " + SERIE.Trim() + FOLIO.Trim() + " no está disponible para consulta";
                RutaPdf = GeneraPdfMensaje(RutaPdf);
            }
            if (RutaPdf.Trim() != "" && RutaPdf.IndexOf("ERROR:") == -1 && File.Exists(RutaPdf))
            {
                try
                {
                    FileStream foto = new FileStream(RutaPdf, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    byte[] arreglo = new byte[foto.Length];
                    BinaryReader reader = new BinaryReader(foto);
                    arreglo = (objRegresar.pdf = reader.ReadBytes(Convert.ToInt32(foto.Length)));
                    foto.Flush();
                    foto.Close();
                }
                catch (Exception exLB)
                {
                    Debug.WriteLine(exLB.Message);
                }
                finally
                {
                    FileInfo fires = new FileInfo(RutaPdf);
                    if (fires.Name.ToUpper().IndexOf(".PDF") > -1)
                    {
                        fires.Delete();
                    }
                }
            }

            if (RutaXml.Trim() != "" && RutaXml.IndexOf("ERROR:") == -1 && File.Exists(RutaXml))
            {
                try
                {
                    FileStream foto = new FileStream(RutaXml, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    byte[] arreglo = new byte[foto.Length];
                    BinaryReader reader = new BinaryReader(foto);
                    arreglo = (objRegresar.xml = reader.ReadBytes(Convert.ToInt32(foto.Length)));
                    foto.Flush();
                    foto.Close();
                }
                catch (Exception exLBXml)
                {
                    Debug.WriteLine(exLBXml.Message);
                }
                finally
                {
                    FileInfo fires = new FileInfo(RutaXml);
                    if (fires.Name.ToUpper().IndexOf(".XML") > -1)
                    {
                        fires.Delete();
                    }
                }
            }

            return objRegresar;
        }


}
}