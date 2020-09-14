using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonFNR.BL
{
    class clsSolicitudDeTraslado : ComportaForm
    {
        #region CONSTANTES

        private const int FRM_SOLICITUD_DE_TRASLADO = 1250000940;

        private const string GRID_ARTICULOS = "23";
        private const string COLUMNA_SERIE_INICIO = "U_SerieIni";
        private const string COLUMNA_SERIE_FIN = "U_SerieFin";
        private const int CHAR_PRESS_ENTER = 13;

        //ENCABEZADO
        private const string LBL_SOCIO_NEGOCIO = "5";
        private const string LBL_NOMBRE_SOCIO_NEGOCIO = "8";
        private const string LBL_PERSONA_CONTACTO = "33";
        private const string LBL_DESTINATARIO = "10";
        private const string LBL_FECHA_VENCIMIENTO = "1250000071";
        private const string LBL_FECHA_DOCUMENTO = "17";
        private const string LBL_LISTA_DE_PRECIOS = "37";
        private const string LBL_OFICINA = "27";

        private const string TXT_SOCIO_NEGOCIO = "3";
        private const string TXT_NOMBRE_SOCIO_NEGOCIO ="7";
        private const string TXT_PERSONA_CONTACTO = "31";
        private const string TXT_DESTINATARIO = "9";
        private const string TXT_FECHA_VENCIMIENTO = "1250000072";
        private const string TXT_FECHA_DOCUMENTO = "16";
        private const string TXT_OFICINA = "25";

        private const string CMB_LISTA_PRECIOS = "36";
        private const string CMB_SHIP_TO_CODE = "254000004";

        private const string TIPO_MOVIMIENTO = "U_TipoMov";
      
        //CAMPOS NUEVOS DE VENTANA
        private const string LBL_NOMBRE_PROVEEDOR = "lblNomPro";
        private const string LBL_SERIE = "lblSerie";

        private const string TXT_NOMBRE_PROVEEDOR = "NombreP";
        private const string TXT_SERIE = "Serie";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForm = null;
        private static bool _oSolicitudDeTraslado = false;

        private SAPbouiCOM.StaticText _oLblNombreProveedor = null;
        private SAPbouiCOM.StaticText _oLblSerie = null;
        private SAPbouiCOM.EditText _oTxtNombrePromotor = null;
        private SAPbouiCOM.EditText _oTxtSerie = null;
        private SAPbouiCOM.EditText _oTxtNombreOficina = null;
        private SAPbouiCOM.ComboBox _oCmbTipoMovimiento = null;
        private string TipoMovimiento = null;

        private int _oContadorFormas = 0;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de la solicitud de traslado
        /// </summary>
        /// <param name="_Application">Objeto de la conexión de SAP</param>
        /// <param name="_Company">Objeto de la empresa</param>
        /// <param name="form">Nombre de la forma</param>
        public clsSolicitudDeTraslado(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oSolicitudDeTraslado == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                setEventos();
                _oSolicitudDeTraslado = true;
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
                if (pVal.BeforeAction == false && pVal.FormType == FRM_SOLICITUD_DE_TRASLADO)
                {
                    if (pVal.EventType == BoEventTypes.et_FORM_RESIZE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        OcultarControlesVentana(_oForm);
                        CrearCamposDeUsuario(_oForm);
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
                    {
                        if (_oContadorFormas == 1)
                        {
                            _Application.ItemEvent -= new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                            Dispose();
                            application = null;
                            company = null;
                            _oSolicitudDeTraslado = false;
                            Addon.typeList.RemoveAll(p => p._forma == formID);
                            return;
                        }
                        else
                        {
                            _oContadorFormas -= 1;
                        }
                    }

                    if(pVal.EventType == BoEventTypes.et_FORM_ACTIVATE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        CrearCamposDeUsuario(_oForm);
                    }                   
                }

                if (pVal.BeforeAction == false && pVal.FormType == -FRM_SOLICITUD_DE_TRASLADO)
                {
                    if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == TIPO_MOVIMIENTO)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        _oForm.Freeze(true);
                        SAPbouiCOM.Form F1 = _Application.Forms.GetFormByTypeAndCount(Convert.ToInt32(_oForm.TypeEx.TrimStart('-')), _oForm.TypeCount);
                        F1.Freeze(true);
                        _oTxtNombreOficina = F1.Items.Item(TXT_OFICINA).Specific;
                        _oTxtNombrePromotor = F1.Items.Item(TXT_NOMBRE_PROVEEDOR).Specific;
                        _oCmbTipoMovimiento = _oForm.Items.Item(TIPO_MOVIMIENTO).Specific;

                        if (!string.IsNullOrEmpty(_oCmbTipoMovimiento.Value.ToString()))
                        {
                            TipoMovimiento = _oCmbTipoMovimiento.Value.ToString().TrimEnd(' ');
                            if (TipoMovimiento != "PROMOTORES - OFICINAS")
                            {
                                if (_oForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    _oTxtNombreOficina.Value = Extensor.ObtenerSecretaria(_Company, "U_codigo_secretaria");
                                    _oTxtNombrePromotor.Value = Extensor.ObtenerSecretaria(_Company, "T0.U_nombre_secretaria");
                                }
                            }
                            else
                            {
                                if (_oForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    _oTxtNombreOficina.Value = "";
                                    _oTxtNombrePromotor.Value = "";
                                }
                            }
                        }
                        _oForm.Freeze(false);
                        F1.Freeze(false);
                    }
                }

                if (pVal.BeforeAction == true && pVal.FormType == FRM_SOLICITUD_DE_TRASLADO)
                {
                    if (pVal.EventType == BoEventTypes.et_KEY_DOWN && pVal.ItemUID == GRID_ARTICULOS && pVal.ColUID == COLUMNA_SERIE_INICIO && pVal.CharPressed == CHAR_PRESS_ENTER)
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
                throw new Exception("Error en método 'eventos' *clsSolicitudDeTraslado* : " + ex.Message);
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
                SAPbouiCOM.Item oItemOficina = null;

                //LABELS
                oItem = _oForm.Items.Item(LBL_SOCIO_NEGOCIO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_NOMBRE_SOCIO_NEGOCIO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_PERSONA_CONTACTO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_DESTINATARIO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_FECHA_VENCIMIENTO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_FECHA_DOCUMENTO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_LISTA_DE_PRECIOS);
                oItem.Visible = false;

                //TEXBOX
                oItem = _oForm.Items.Item(TXT_SOCIO_NEGOCIO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(TXT_NOMBRE_SOCIO_NEGOCIO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(TXT_PERSONA_CONTACTO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(TXT_DESTINATARIO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(TXT_FECHA_VENCIMIENTO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(TXT_FECHA_DOCUMENTO);
                oItem.Visible = false;

                //COMBOS
                oItem = _oForm.Items.Item(CMB_LISTA_PRECIOS);
                oItem.Visible = false;

                oItem = _oForm.Items.Item(CMB_SHIP_TO_CODE);
                oItem.Visible = false;
     
                //CAMPOS OFICINA
                oItem = _oForm.Items.Item(LBL_OFICINA);
                oItem.Top = 10;
                oItem.Left = 5;
                oItem.Width = 70;
                oItem.TextStyle = 0;

                oItemOficina = _oForm.Items.Item(TXT_OFICINA);
                oItemOficina.Top = oItem.Top;
                oItemOficina.Left = oItem.Left + 80;

            }
            catch (Exception ex)
            {
                throw new Exception("Error al ocultar controles *OcultarControlesVentana* : " + ex.Message);
            }
        }
        
        /// <summary>
        /// Crea los campos definido por el usuario 
        /// </summary>
        /// <param name="_oForma">Forma activa</param>
        private void CrearCamposDeUsuario(Form _oForma)
        {
            SAPbouiCOM.Item newItem = null;
            try
            {
                try
                {
                    string s = _oForma.Items.Item(TXT_NOMBRE_PROVEEDOR).UniqueID;
                }
                catch (Exception)
                {
                    _oForma.Freeze(true);                 

                    //Label 'Nombre proveedor' ligado al campo de Label 'oficina'.
                    SAPbouiCOM.Item _olblOficina = null;
                    _olblOficina = _oForma.Items.Item(LBL_OFICINA);
                    newItem = _oForma.Items.Add(LBL_NOMBRE_PROVEEDOR, BoFormItemTypes.it_STATIC);
                    newItem.Left = _olblOficina.Left;
                    newItem.Top = _olblOficina.Top + 18;
                    newItem.Width = 70;
                    newItem.ToPane = 0;
                    _oLblNombreProveedor = newItem.Specific;
                    _oLblNombreProveedor.Caption = "Nombre";                  

                    //Campo texto 'Nombre proveedor' ligado al campo de Label 'Nombre proveedor'.
                    SAPbouiCOM.Item _oLblNP = null;
                    _oLblNP = _oForma.Items.Item(LBL_NOMBRE_PROVEEDOR);
                    newItem = _oForma.Items.Add(TXT_NOMBRE_PROVEEDOR, BoFormItemTypes.it_EDIT);
                    newItem.Left = _oLblNP.Left + 80;
                    newItem.Top = _oLblNP.Top;
                    newItem.Width = 141;
                    newItem.Height = 15;
                    newItem.ToPane = 0;
                    _oTxtNombrePromotor = (SAPbouiCOM.EditText)newItem.Specific;
                    _oTxtNombrePromotor.DataBind.SetBound(true, "OWTQ", "U_NombreP");
                    _oLblNP.LinkTo = newItem.UniqueID;
                  
                    _oContadorFormas += 1;
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Error al crear campos de usuario *CrearCamposDeUsuario* : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        #endregion
    }
}
