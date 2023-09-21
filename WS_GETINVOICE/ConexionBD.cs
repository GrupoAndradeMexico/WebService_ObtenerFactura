using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
 
namespace WS_GETINVOICE
{    
    class ConexionBD
    {
        string strConBDC = "";
        //string TipoConexion = "BDLOCAL";
        public static string MensajeError = "";
        private System.Data.SqlClient.SqlConnection sqlConnection1;
        private System.Data.SqlClient.SqlCommand sqlCommand1;
        //WSCG.SWparaControlGastos Conexion;
        public static int RegAfec = 0;

        public string DameCadenaConexion
        {
            get {
                return strConBDC; 
            }        
        }

        public bool EstaVacio(DataSet ds)
        {
            bool res = true;
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                res = false;
            return res;
        }

        public ConexionBD(string CadenaBD)
        {            
            this.strConBDC = CadenaBD.Trim();      
            MensajeError = "";
            RegAfec = 0; 
            
                this.sqlConnection1 = new System.Data.SqlClient.SqlConnection();
                this.sqlCommand1 = new System.Data.SqlClient.SqlCommand();            
        }

        public DataSet Consulta(string Q)
        {
            DataSet ds = new DataSet();
            try
            {
                    if (this.sqlConnection1.State.ToString().ToUpper().Trim() != "OPEN")
                    {
                        this.sqlConnection1.ConnectionString = this.strConBDC.Trim();
                        this.sqlConnection1.Open();
                    }
                    System.Data.SqlClient.SqlDataAdapter objAdaptador = new System.Data.SqlClient.SqlDataAdapter(Q, this.strConBDC);
                    objAdaptador.Fill(ds, "Command");
                    if (ds.Tables.Count > 0)
                        RegAfec = ds.Tables[0].Rows.Count;
                    else
                        RegAfec = 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            finally 
            {   
                this.sqlConnection1.Close(); 
            }
            return ds;
        }

        public string ConsultaUnSoloCampo(string Q)
        {
            string res = String.Empty;
            DataSet ds = new DataSet();
            try
            {
                    if (this.sqlConnection1.State.ToString().ToUpper().Trim() != "OPEN")
                    {
                        this.sqlConnection1.ConnectionString = this.strConBDC.Trim();
                        this.sqlConnection1.Open();
                    }
                    System.Data.SqlClient.SqlDataAdapter objAdaptador = new System.Data.SqlClient.SqlDataAdapter(Q, this.strConBDC);
                    objAdaptador.Fill(ds, "Command");
                    if (ds.Tables.Count > 0)
                    {
                        if (ds.Tables[0].Rows.Count > 0)
                        {//no importa cuantos registros traiga siempre regresará solo la primer columna y del primer registro. 
                            res = ds.Tables[0].Rows[0][0].ToString(); 
                        }
                    }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            finally
            {   //LJBA: 20101108 la siguiente línea se agregó para ver si es solución al arranamiento de la barra
            this.sqlConnection1.Close();
            }
            return res;
        }

        public string getPasswordPdf(string fileName)
        {
            const string servidor = "192.168.20.29";
            const string usuario = "sa";
            const string contrasenia = "S0p0rt3";
            const string nombreBD = "Tramites";
            const string storedProcedure = "consultaPDF";

            DataTable infoPDF = new DataTable();
            string passFile = "";

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = servidor;
            builder.InitialCatalog = nombreBD;
            builder.UserID = usuario;
            builder.Password = contrasenia;

            using(SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                try
                {
                    if(connection.State != ConnectionState.Open)
                    {
                        connection.Open();
                    }

                    using(SqlCommand command = new SqlCommand())
                    {
                        command.Connection = connection;
                        command.CommandText = storedProcedure;
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@nombre", fileName).SqlDbType = SqlDbType.VarChar;

                        using (SqlDataReader objSqlDataReader = command.ExecuteReader(CommandBehavior.CloseConnection))
                        {
                            infoPDF.Load(objSqlDataReader);
                            if(infoPDF.Rows.Count > 0)
                            {
                                passFile = infoPDF.Rows[0][3].ToString();
                            }
                            else
                            {
                                passFile = "";
                            }
                            
                            return passFile;
                        }
                    }
                }
                catch(Exception ex)
                {
                    return "0";
                }
            }

        }



        //Dependiendo del tipo de conexion ejecuta un store Procedure, siempre regresa un 
        //arreglo de 10 lugares. En la posicion 0 si tuvo éxito o no y en las demás posiciones
        //el resultado de los parámetros de salida. Todos los arreglos son requeridos y deben de 
        //ser del mismo tamaño.
        public string[] EjecStoreProc(string NomSP, string[] NombresP, string[] TiposP, int[] TamanosP, string[] DireccionesP, string[] ValoresP)
        {
            string[] res = { "", "", "", "", "", "", "", "", "", "" };

            //if (TipoConexion == "SERVICIOWEB")
            //{
            //     res = Conexion.EjecutarStoreProc(NomSP, NombresP, TiposP, TamanosP, DireccionesP, ValoresP);
            //}
            //else 
 
            //Verificar el correcto paso de los parámetros y como se le agregan al sp.
               //pagina 512 del libro C# La biblia.
 
                if (this.sqlConnection1.State.ToString().ToUpper().Trim() != "OPEN")
                {
                    this.sqlConnection1.ConnectionString = this.strConBDC.Trim();
                    this.sqlConnection1.Open();
                }                
                string Q = "[" + NomSP + "]";
                this.sqlCommand1 = new System.Data.SqlClient.SqlCommand(Q,this.sqlConnection1);   
                this.sqlCommand1.CommandType = CommandType.StoredProcedure; 
                //System.Data.SqlClient.SqlDataAdapter objAdaptador = new System.Data.SqlClient.SqlDataAdapter();
                //DataSet ds = new DataSet(); 
                SqlDbType tipoparam = SqlDbType.VarChar;
                int tamano = 50; //agregamos un parametro por cada uno que llegue
                for (int i = 0; i < NombresP.Length; i++)
                {
                    string tipo = TiposP[i].Trim();
                    tamano = TamanosP[i];
                    string direccion = DireccionesP[i].Trim();
                    string valor = ValoresP[i];
                    string nomparametro = NombresP[i].Trim();
                    nomparametro = nomparametro.Replace("@", "");
                    nomparametro = "@" + nomparametro.Trim();

                    switch (tipo)
                    {
                        case "VarChar": tipoparam = SqlDbType.VarChar; break;
                        case "Int": tipoparam = SqlDbType.Int; break;
                        case "Double": tipoparam = SqlDbType.Float; break;
                        case "Real": tipoparam = SqlDbType.Real; break;
                        case "DateTime": tipoparam = SqlDbType.DateTime; break;
                        case "Date": tipoparam = SqlDbType.DateTime; break;
                        case "Decimal": tipoparam = SqlDbType.Decimal; break;
                        case "Char": tipoparam = SqlDbType.Char; break;
                        case "Text": tipoparam = SqlDbType.Text; break;
                        default: tipoparam = SqlDbType.VarChar; break;
                    }
                   SqlParameter param = new SqlParameter(nomparametro, tipoparam, tamano);

                    switch (direccion)
                    {
                        case "Input": param.Direction = ParameterDirection.Input; break;
                        case "InputOutput": param.Direction = ParameterDirection.InputOutput; break;
                        case "Output": param.Direction = ParameterDirection.Output; break;
                        case "ReturnValue": param.Direction = ParameterDirection.ReturnValue; break;
                        default: param.Direction = ParameterDirection.Input; break;
                    }
                    //dependiendo del tipo se convierte el valor;
                    if (tipoparam == SqlDbType.Int)
                    {
                        param.Value = valor.ToUpper().Trim() == "NULL" ? param.Value  : Convert.ToInt32(valor);
                    }

                    if (tipoparam == SqlDbType.Float || tipoparam == SqlDbType.Real)
                    {
                        param.Value = Convert.ToDouble(valor);
                    }
                    if (tipoparam == SqlDbType.Decimal)
                    {
                        param.Value = Convert.ToDecimal(valor);
                    }
                    if (tipoparam == SqlDbType.VarChar || tipoparam == SqlDbType.Char || tipoparam == SqlDbType.Text)
                    {
                        param.Value = valor.ToUpper().Trim()=="NULL" ? DBNull.Value.ToString().Trim() : valor.Trim();
                    }
                    if (tipoparam == SqlDbType.DateTime || tipoparam == SqlDbType.DateTime)
                    {
                        param.Value = Convert.ToDateTime(valor);
                    }

                    //OdbcParameter param = odbcCommand1.Parameters.Add(nomparametro,tipoparam,tamano);
                    this.sqlCommand1.Parameters.Add(param);
                }
                //ejecutamos el store procedure.
                int numRows = this.sqlCommand1.ExecuteNonQuery(); 
                
                res[0] = "Ejecución exitosa";
                //ahora llenamos el arreglo resultante
                //ahora por cada parámetro que haya sido definido como OUTPUT 
                //agregamos su resultado en el orden en que llegarón al arreglo de resultados
                int contparametrossalida = 1;
                for (int j = 0; j < this.sqlCommand1.Parameters.Count - 1; j++)
                {
                    if (this.sqlCommand1.Parameters[j].Direction == ParameterDirection.Output)
                    {
                        res[contparametrossalida] = this.sqlCommand1.Parameters[j].Value.ToString();
                        contparametrossalida++;
                    }
                }

            
        
         return res;    
        }


        //Dependiendo del tipo de conexion ejecuta una instrucción.
        //regresa el número de registros afectados.
        public int EjecUnaInstruccion(string Q)
        {
            int res = 0;
            if (Q.Trim() != "")
            {                
                try
                    {
                        if (this.sqlConnection1.State.ToString().ToUpper().Trim() == "CLOSED")
                        {
                            this.sqlConnection1.ConnectionString = this.strConBDC.Trim();                            
                            this.sqlConnection1.Open();
                        }
                        this.sqlCommand1.Connection = this.sqlConnection1;
                        this.sqlCommand1.CommandType = CommandType.Text;
                        this.sqlCommand1.CommandText = Q.Trim();
                        res = this.sqlCommand1.ExecuteNonQuery();
                    }
                    catch (Exception e)
                    {
                        Q = e.Message;
                    }
                    finally
                    {
                        this.sqlConnection1.Close();
                    }                               
            }            
            return res;
        }


    }
}
