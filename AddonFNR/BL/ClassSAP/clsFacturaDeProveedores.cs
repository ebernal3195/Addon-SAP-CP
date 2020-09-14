using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonFNR.BL
{
    class clsFacturaDeProveedores : ComportaForm
    {
        #region CONSTANTES

        private const int FRM_FACTURA_DE_PROVEEDOR = 141;

        //ENCABEZADO
        private const string LBL_NUMERO_FOLIO = "84";
        private const string LBL_GUION = "210";

        private const string TXT_FOLIO_PREFIJO = "208";
        private const string TXT_FOLIO_NUM = "211";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForm = null;
        private static bool _oFacturaProveedores = false;
        private int _oContadorFormas = 0;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de la factura de proveedor
        /// </summary>
        /// <param name="_Application">Objeto de la conexión de SAP</param>
        /// <param name="_Company">Objeto de la empresa</param>
        /// <param name="form">Nombre de la forma</param>
        public clsFacturaDeProveedores(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oFacturaProveedores == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                setEventos();
                _oFacturaProveedores = true;
            }
        }

        #endregion

        #region EVENTOS

        /// <summary>
        /// Eventos de la forma activa
        /// </summary>
        /// <param name="FormUID">Id de la forma</param>
        /// <param name="pVal">Propiedades de la forma</param>
        /// <param name="BubbleEvent">true/false</param>
        public void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                eventos(FormUID, ref pVal, out BubbleEvent);
            }
            catch (Exception ex)
            {
                _Application.MessageBox("Ocurrió un error en ItemEvent: " + ex.Message);
            }
        }

        /// <summary>
        /// Eventos de la forma activa
        /// </summary>
        /// <param name="FormUID">Id de la forma</param>
        /// <param name="pVal">Propiedades de la forma</param>
        /// <param name="BubbleEvent">Evento true/false</param>
        public override void eventos(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                if (pVal.BeforeAction == false && pVal.FormType == FRM_FACTURA_DE_PROVEEDOR)
                {
                    if (pVal.EventType == BoEventTypes.et_FORM_RESIZE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        OcultarControlesVentana(_oForm);
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
                    {
                        if (_oContadorFormas == 1)
                        {
                            _Application.ItemEvent -= new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                            Dispose();
                            application = null;
                            company = null;
                            _oFacturaProveedores = false;
                            Addon.typeList.RemoveAll(p => p._forma == formID);
                            return;
                        }
                        else
                        {
                            _oContadorFormas -= 1;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en método 'eventos' *clsFacturaDeProveedores* : " + ex.Message);
            }
        }

        #endregion

        #region METODOS

        /// <summary>
        /// Liberar recursos
        /// </summary>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Inicializa los eventos de la forma
        /// </summary>
        private void setEventos()
        {
            _Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
        }


        /// <summary>
        /// Oculta los controles de la ventana activa
        /// </summary>
        /// <param name="_oForm">Forma activa</param>
        private void OcultarControlesVentana(Form _oForm)
        {
            try
            {
                SAPbouiCOM.Item oItem = null;

                //LABELS
                oItem = _oForm.Items.Item(LBL_NUMERO_FOLIO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_GUION);
                oItem.Visible = false;

                //TEXBOX
                oItem = _oForm.Items.Item(TXT_FOLIO_PREFIJO);
                oItem.Visible = false;
                
                oItem = _oForm.Items.Item(TXT_FOLIO_NUM);
                oItem.Visible = false;
            
                _oContadorFormas += 1;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al ocultar controles *OcultarControlesVentana* : " + ex.Message);
            }
        }
        
        #endregion
    }
}
