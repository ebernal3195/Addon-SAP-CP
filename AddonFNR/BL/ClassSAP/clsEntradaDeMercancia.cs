using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonFNR.BL
{
    class clsEntradaDeMercancia : ComportaForm
    {

        #region CONSTANTES

        private const int FRM_ENTRADA_DE_MERCANCIA = 143;
        private const string OBJETO_ENTRADA_MERCANCIA = "20";
        private const string GRID_ARTICULOS = "38";
        private const string COLUMNA_CLAVE_ARTICULO = "1";
        private const string COLUMNA_SERIE_INICIO = "U_SerieIni";
        private const string COLUMNA_SERIE_FIN = "U_SerieFin";
        private const int CHAR_PRESS_ENTER = 13;
        private const string VENTANA_EMERGENTE = "0";

        //ENCABEZADO
        private const string LBL_PERSONA_CONTACTO = "83";
        private const string LBL_NUMERO_REFERENCIA = "15";
        private const string TXT_NUMERO_REFERENCIA = "14";
        private const string CMB_PERSONA_CONTACTO = "85";
        private const string ICO_PERSONA_CONTACTO = "80";

        //PESTAÑAS
        private const string FLD_ANEXOS = "1320002137";
        private const string FLD_LOGISTICA = "114";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForm = null;
        private static bool _oEntradaDeMercancia = false;
        private static List<Datos> lDatos = new List<Datos>();
        private static Datos itemDatos = new Datos();
        private SAPbouiCOM.Matrix _oMatrixArticulos = null;

        private SAPbouiCOM.EditText oItemCode = null;
        private SAPbouiCOM.EditText oSerieInicio = null;
        private SAPbouiCOM.EditText oSerieFin = null;
        private int _oContadorFormas = 0;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de la entrada de mercancía
        /// </summary>
        /// <param name="_Application">Objeto de la conexión de SAP</param>
        /// <param name="_Company">Objeto de la empresa</param>
        /// <param name="form">Nombre de la forma</param>
        public clsEntradaDeMercancia(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oEntradaDeMercancia == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                setEventos();
                _oEntradaDeMercancia = true;
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
                if (pVal.BeforeAction == false && pVal.FormType == FRM_ENTRADA_DE_MERCANCIA)
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
                            _Application.FormDataEvent -= new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormEvent);
                            Dispose();
                            application = null;
                            company = null;
                            _oEntradaDeMercancia = false;
                            Addon.typeList.RemoveAll(p => p._forma == formID);
                            return;
                        }
                        else
                        {
                            _oContadorFormas -= 1;
                        }
                    }
                }

                if(pVal.BeforeAction == true && pVal.FormType == FRM_ENTRADA_DE_MERCANCIA)
                {
                    if(pVal.EventType == BoEventTypes.et_KEY_DOWN && pVal.ItemUID == GRID_ARTICULOS && pVal.ColUID == COLUMNA_SERIE_INICIO && pVal.CharPressed == CHAR_PRESS_ENTER)
                    {
                        bubbleEvent = false;
                        return;
                    }
                    if (pVal.EventType == BoEventTypes.et_KEY_DOWN && pVal.ItemUID == GRID_ARTICULOS && pVal.ColUID == COLUMNA_SERIE_FIN && pVal.CharPressed == CHAR_PRESS_ENTER)
                    {
                        bubbleEvent = false;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en método 'eventos' *clsEntradaDeMercancia* : " + ex.Message);
            }
        }

        /// <summary>
        ///Se producen cuando la aplicación realiza las acciones siguientes en formularios conectados a objetos de negocio:
        ///- Añadir
        ///- Actualizar
        ///- Borrar      
        /// </summary>
        /// <param name="BusinessObjectInfo">
        /// Información del objeto aplicado
        /// </param>
        /// <param name="BubbleEvent">
        /// true/false
        /// </param>
        private void SBO_Application_FormEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.BeforeAction == true && BusinessObjectInfo.FormTypeEx == FRM_ENTRADA_DE_MERCANCIA.ToString())
                {
                    if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        && BusinessObjectInfo.Type == OBJETO_ENTRADA_MERCANCIA && BusinessObjectInfo.ActionSuccess == false)
                    {
                        lDatos.Clear();
                        SAPbouiCOM.Form _oNuevaForm = _Application.Forms.GetForm(BusinessObjectInfo.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);
                        if (BusinessObjectInfo.FormTypeEx == FRM_ENTRADA_DE_MERCANCIA.ToString())
                        {
                            _oMatrixArticulos = _oNuevaForm.Items.Item(GRID_ARTICULOS).Specific;

                            for (int noLinea = 1; noLinea < _oMatrixArticulos.RowCount; noLinea++)
                            {
                                oItemCode = (SAPbouiCOM.EditText)_oMatrixArticulos.Columns.Item(COLUMNA_CLAVE_ARTICULO).Cells.Item(noLinea).Specific;
                                if (oItemCode.Value.Substring(0, 2).ToString() == "PL")
                                {
                                    oSerieInicio = (SAPbouiCOM.EditText)_oMatrixArticulos.Columns.Item(COLUMNA_SERIE_INICIO).Cells.Item(noLinea).Specific;
                                    oSerieFin = (SAPbouiCOM.EditText)_oMatrixArticulos.Columns.Item(COLUMNA_SERIE_FIN).Cells.Item(noLinea).Specific;
                                    if (!string.IsNullOrEmpty(oSerieInicio.Value.ToString()) && !string.IsNullOrEmpty(oSerieFin.Value.ToString()))
                                    {
                                        itemDatos = new Datos();
                                        itemDatos.itemCode = oItemCode.Value.ToString();
                                        itemDatos.serieInial = oSerieInicio.Value.ToString();
                                        itemDatos.serieFinal = oSerieFin.Value.ToString();
                                        lDatos.Add(itemDatos);
                                    }
                                }
                            }
                            if (lDatos.Count != 0)
                            {
                                Addon.Instance.Ejecutaclase("21", lDatos);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _Application.MessageBox("Error en FormEvent *clsEntradaDeMercancia* : " + ex.Message);
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
            _Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormEvent);
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
                oItem = _oForm.Items.Item(LBL_PERSONA_CONTACTO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_NUMERO_REFERENCIA);
                oItem.Visible = false;

                //TEXBOX
                oItem = _oForm.Items.Item(TXT_NUMERO_REFERENCIA);
                oItem.Visible = false;

                //COMBOS
                oItem = _oForm.Items.Item(CMB_PERSONA_CONTACTO);
                oItem.Visible = false;

                //ICONOS CHOOSE
                oItem = _oForm.Items.Item(ICO_PERSONA_CONTACTO);
                oItem.Visible = false;

                //FOLDER
                oItem = _oForm.Items.Item(FLD_ANEXOS);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(FLD_LOGISTICA);
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
