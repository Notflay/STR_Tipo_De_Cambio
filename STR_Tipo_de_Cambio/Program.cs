using STR_Tipo_de_Cambio.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using HtmlAgilityPack;
using ScrapySharp.Extensions;

namespace STR_Tipo_de_Cambio
{
    class Program
    {
        static void Main(string[] args)
        {

            SAPConnector.Conectar();
            IntegrarTipodeCambio();
            SAPConnector.Desconectar();

        }

        private static void IntegrarTipodeCambio()
        {

            try
            {
                string date = System.DateTime.Now.ToString("yyyy-MM-dd");

                #region Conexión a SBS
                List<string> TipoCambio = new List<string>();

                HtmlWeb oWeb = new HtmlWeb();
                HtmlDocument doc = oWeb.Load("https://www.sbs.gob.pe/app/pp/sistip_portal/paginas/publicacion/tipocambiopromedio.aspx");

                HtmlNode Body = doc.DocumentNode.CssSelect("body").First();
                string sbody = Body.InnerHtml;

                foreach (var Node in doc.DocumentNode.CssSelect(".APLI_fila2"))
                {
                    TipoCambio.Add(Node.InnerHtml);
                }
                if (string.IsNullOrEmpty(TipoCambio[1]))
                    throw new Exception("Aun no se ha actualizado el tipo de Cambio SBS");
                double tipoCambio = Convert.ToDouble(TipoCambio[1]);
                #endregion

                SAPbobsCOM.SBObob bo = SAPConnector.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                Log.WriteToFile("Actualizando Tipo de Cambio " + date + ".....");
                bo.SetCurrencyRate("USD", System.DateTime.Now, tipoCambio, true);
                Log.WriteToFile("Tipo de Cambio del dia " + date + " : " + tipoCambio.ToString("F2"));

            }
            catch (Exception ex)
            {
                Log.WriteToFile("Error al Actualizar en SAP :" + ex.Message);
            }

        }
    }
}
