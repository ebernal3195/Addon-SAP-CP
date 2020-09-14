using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonFNR.BL
{
    class clsFacturaClientes : ComportaForm
    {
        #region CONSTANTES

        private const int FRM_FACTURA_CLIENTES = 133;
        private const string OBJETO_FACTURA_CLIENTES = "13";
        private const string GRID_ARTICULOS = "38";
        private const string COLUMNA_CLAVE_ARTICULO = "1";
        private const string COLUMNA_SERIE = "U_Serie";
        private const string BTN_CREAR = "1";
        
        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForm = null;
        private static bool _oFacturaClientes = false;      
        private SAPbobsCOM.Recordset _oRec = null;

        private static List<Datos> lDatos = new List<Datos>();
        private static Datos itemDatos = new Datos();
        private SAPbouiCOM.Matrix _oMatrixArticulos = null;

        private SAPbouiCOM.EditText oItemCode = null;       
        private SAPbouiCOM.EditText oSerie = null;      

        private int _oContadorFormas = 0;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de la factura de clientes
        /// </summary>
        /// <param name="_Application">Objeto de la conexión de SAP</param>
        /// <param name="_Company">Objeto de la empresa</param>
        /// <param name="form">Nombre de la forma</param>
        public clsFacturaClientes(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oFacturaClientes == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                setEventos();
                _oFacturaClientes = true;
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
                if (pVal.BeforeAction == false && pVal.FormType == FRM_FACTURA_CLIENTES)
                {          
                    if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
                    {
                        if (_oContadorFormas == 1)
                        {
                            _Application.ItemEvent -= new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                            _Application.StatusBarEvent -= new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(SBO_Application_StatusBarEvent);
                            Dispose();
                            application = null;
                            company = null;
                            _oFacturaClientes = false;
                            Addon.typeList.RemoveAll(p => p._forma == formID);
                            return;
                        }
                        else
                        {
                            _oContadorFormas -= 1;
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_RESIZE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_ACTIVATE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    }
                }                     
            }
            catch (Exception ex)
            {
                throw new Exception("Error en método 'eventos' *clsFacturaClientes* : " + ex.Message);
            }
        }

        /// <summary>
        /// Evento que se ejecuta cuando se encuentra un mensaje de status
        /// </summary>
        /// <param name="Text">Texto que se esta mostrando</param>
        /// <param name="messageType">Tipo de mensaje</param>
        private void SBO_Application_StatusBarEvent(string Text, BoStatusBarMessageType messageType)
        {
            try
            {
                if (Text.Contains("No se puede añadir fila") && messageType == BoStatusBarMessageType.smt_Warning)
                {
                    if (_Application.Forms.ActiveForm.TypeEx == FRM_FACTURA_CLIENTES.ToString())
                    {
                        if (_oForm.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            lDatos.Clear();
                            _oMatrixArticulos = _oForm.Items.Item(GRID_ARTICULOS).Specific;
                            int x = 1;
                            for (int noLinea = 1; noLinea < _oMatrixArticulos.RowCount; noLinea++)
                            {
                                oItemCode = (SAPbouiCOM.EditText)_oMatrixArticulos.Columns.Item(COLUMNA_CLAVE_ARTICULO).Cells.Item(noLinea).Specific;
                                if (oItemCode.Value.Substring(0, 2).ToString() == "PL")
                                {
                                    oSerie = (SAPbouiCOM.EditText)_oMatrixArticulos.Columns.Item(COLUMNA_SERIE).Cells.Item(noLinea).Specific;
                                    if (!string.IsNullOrEmpty(oSerie.Value.ToString()))
                                    {
                                        itemDatos = new Datos();
                                        itemDatos.itemCode = oItemCode.Value.ToString();
                                        itemDatos.serieInial = oSerie.Value.ToString();
                                        itemDatos.serieFinal = oSerie.Value.ToString();
                                        itemDatos.noLinea = x;
                                        lDatos.Add(itemDatos);
                                        x += 1;
                                    }
                                }
                            }
                            if (lDatos.Count != 0)
                            {
                                Addon.Instance.Ejecutaclase("25", lDatos);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _Application.MessageBox("Error en StatusBarEvent *clsFacturaClientes* : " + ex.Message);
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
            _Application.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(SBO_Application_StatusBarEvent);
        }              

        #endregion
    }
}
