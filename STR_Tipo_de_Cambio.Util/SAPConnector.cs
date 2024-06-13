using SAPbobsCOM;
using System;
using System.Configuration;

namespace STR_Tipo_de_Cambio.Util
{
    public static class SAPConnector
    {
        public static Company SboCompany { get; set; }

        static SAPConnector()
        {
            try
            {
                /*
                SboCompany = new Company();
                SboCompany.Server = ConfigurationManager.AppSettings["SAP_SERVIDOR"];
                SboCompany.CompanyDB = ConfigurationManager.AppSettings["SAP_BASE"];
                SboCompany.DbServerType = getTypeDB(ConfigurationManager.AppSettings["SAP_TIPO_BASE"]);
                SboCompany.DbUserName = ConfigurationManager.AppSettings["SAP_DBUSUARIO"];
                SboCompany.DbPassword = ConfigurationManager.AppSettings["SAP_DBPASSWORD"];
                SboCompany.UserName = ConfigurationManager.AppSettings["SAP_USUARIO"];
                SboCompany.Password = ConfigurationManager.AppSettings["SAP_PASSWORD"];
                SboCompany.language = BoSuppLangs.ln_Spanish_La;
                */
            }
            catch (Exception ex)
            {
                string mensaje = ex.Message;
                throw;
            }
        }

        public static BoDataServerTypes getTypeDB(string tipoDB) {
            switch (tipoDB)
            {
                case "HANA":
                    return BoDataServerTypes.dst_HANADB;
                case "SQL19":
                    return BoDataServerTypes.dst_MSSQL2019;
                case "SQL17":
                    return BoDataServerTypes.dst_MSSQL2017;
                case "SQL16":
                    return BoDataServerTypes.dst_MSSQL2016;
                case "SQL14":
                    return BoDataServerTypes.dst_MSSQL2014;
                case "SQL12":
                    return BoDataServerTypes.dst_MSSQL2012;
                default:
                    throw new Exception("No existe tipo de base de datos");
            }
        }

        public static void Conectar(SBO sbo)
        {
            try
            {

                SboCompany = new Company();
                SboCompany.Server = sbo.SAP_SERVIDOR;
                SboCompany.CompanyDB = sbo.SAP_BASE;
                SboCompany.DbServerType = getTypeDB(sbo.SAP_TIPO_BASE);
                SboCompany.DbUserName = sbo.SAP_DBUSUARIO;
                SboCompany.DbPassword = sbo.SAP_DBPASSWORD;
                SboCompany.UserName = sbo.SAP_USUARIO;
                SboCompany.Password = sbo.SAP_PASSWORD;
                SboCompany.language = BoSuppLangs.ln_Spanish_La;

                if (SboCompany.Connect() != 0)
                {
                    Log.WriteToFile("CONEXION-SAPConnector:" + SboCompany.GetLastErrorDescription());

                }
                //else
                //ServicioIntegracion.WriteToFile("CONEXION EXITOSA");

            }
            catch (Exception ex)
            {

                Log.WriteToFile("CONEXION :" + ex.Message);

            }

        }
        public static void Desconectar()
        {
            try
            {

                SboCompany.Disconnect();
                //else
                //ServicioIntegracion.WriteToFile("CONEXION EXITOSA");

            }
            catch (Exception ex)
            {

                //ServicioIntegracion.WriteToFile("CONEXION :" + ex.Message);
                throw new Exception(ex.Message);
            }

        }
    }
}
