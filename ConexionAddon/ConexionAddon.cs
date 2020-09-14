using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConexionAddon
{
    public sealed class ConexionAddon
    {
        #region VARIABLES

        public static SAPbouiCOM.Application SBO_Application;
        public static SAPbobsCOM.Company _oCompany;

        #endregion

        #region METODOS

        /// <summary>
        /// Método para realizar la conexión
        /// </summary>
        /// <returns>true/false</returns>
        public bool Conectar()
        {
            EstablecerAplicacion();
            if (EstablecerContextoConexion() != 0)
            {
                return false;
            }
            else
            {
                SBO_Application.StatusBar.SetText("Addon SAP-CP conectado con éxito",SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                return true;
            }
            
        }

        /// <summary>
        /// Establece la aplicación
        /// </summary>
        private void EstablecerAplicacion()
        {
            SAPbouiCOM.SboGuiApi _oSboGuiApi;
            string _oConnectionString;
            _oSboGuiApi = new SAPbouiCOM.SboGuiApi();
            try
            {
                _oConnectionString = Environment.GetCommandLineArgs().GetValue(1).ToString();
                _oSboGuiApi.Connect(_oConnectionString);
                SBO_Application = _oSboGuiApi.GetApplication();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }



        }

        /// <summary>
        /// Establece el contexto de la conexión
        /// </summary>
        /// <returns>Conexión</returns>
        private int EstablecerContextoConexion()
        {
            int _oRetCode = -1;
            try
            {
                _oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
                _oRetCode = 0;
                return _oRetCode;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }

        #endregion    
    }
}
