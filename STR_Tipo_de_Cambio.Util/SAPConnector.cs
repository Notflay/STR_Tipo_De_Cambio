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
                SboCompany = new Company();
                SboCompany.Server = ConfigurationManager.AppSettings["SAP_SERVIDOR"];
                SboCompany.CompanyDB = ConfigurationManager.AppSettings["SAP_BASE"];
                SboCompany.DbServerType = ConfigurationManager.AppSettings["SAP_TIPO_BASE"] == "HANA" ? BoDataServerTypes.dst_HANADB : BoDataServerTypes.dst_MSSQL2017;
                SboCompany.DbUserName = ConfigurationManager.AppSettings["SAP_DBUSUARIO"];
                SboCompany.DbPassword = ConfigurationManager.AppSettings["SAP_DBPASSWORD"];
                SboCompany.UserName = ConfigurationManager.AppSettings["SAP_USUARIO"];
                SboCompany.Password = ConfigurationManager.AppSettings["SAP_PASSWORD"];
                SboCompany.language = BoSuppLangs.ln_Spanish_La;

            }
            catch (Exception ex)
            {
                string mensaje = ex.Message;
                throw;
            }
        }

        public static void Conectar()
        {
            try
            {

                string a = SboCompany.Server;
                string a1 = SboCompany.CompanyDB;
                string a2 = SboCompany.DbUserName;
                string a3 = SboCompany.DbPassword;
                string a4 = SboCompany.UserName;
                string a5 = SboCompany.Password;
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
