using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonFNR
{
    public class Datos
    {
        #region TRANSFERENCIAS Y ENTRADAS DE MERCANCIA

        /// <summary>
        /// Código del artículo
        /// </summary>
        public string itemCode { get; set; }
        /// <summary>
        /// Serie inicio
        /// </summary>
        public string serieInial { get; set; }
        /// <summary>
        /// Serie fin
        /// </summary>
        public string serieFinal { get; set; }
        /// <summary>
        /// Numero de línea
        /// </summary>
        public int noLinea { get; set; } 
 
        #endregion     

        #region TRASPASOS

        /// <summary>
        /// No de traspaso Docnum
        /// </summary>
        public string TrasNoDocumento { get; set; }
        /// <summary>
        /// Número de serie
        /// </summary>
        public string TrasSerie { get; set; }
        /// <summary>
        /// Número de serie interna (Solicitud)
        /// </summary>
        public string TrasSerieInterna { get; set; }
        /// <summary>
        /// Código del articulo o plan
        /// </summary>
        public string TrasCodigoPlan { get; set; }
        /// <summary>
        /// Nombre del articulo o plan
        /// </summary>
        public string TrasNombrePlan { get; set; }
        /// <summary>
        /// Importe del pago inicial
        /// </summary>
        public double TrasImportePagoInicial { get; set; }
        /// <summary>
        /// Importe de la comisión
        /// </summary>
        public double TrasImporteComision { get; set; }
        /// <summary>
        /// Importe de la papelería
        /// </summary>
        public double TrasImportePapeleria { get; set; }
        /// <summary>
        /// Importe recibido
        /// </summary>
        public double TrasImporteRecibido { get; set; }
        /// <summary>
        /// Importe del excedente de la inversión inicial
        /// </summary>
        public double TrasExcedenteInvInicial { get; set; }
        /// <summary>
        /// Importe del Bono
        /// </summary>
        public double TrasImporteBono { get; set; }


        #endregion     
    }
}
