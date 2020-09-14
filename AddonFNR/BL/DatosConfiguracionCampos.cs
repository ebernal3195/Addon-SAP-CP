using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonFNR.BL
{
    public class DatosConfiguracionCampos
    {
        #region OBTIENE LA CONFIGURACION INICIAL

        /// <summary>
        /// Nombre de usuario para autorizar
        /// </summary>
        public string usuario { get; set; }

        /// <summary>
        /// Nombre del campo para autorizar
        /// </summary>
        public string campo { get; set; }

        /// <summary>
        /// Valor de autorización
        /// </summary>
        public bool activo { get; set; }

        #endregion      
    } 
}
