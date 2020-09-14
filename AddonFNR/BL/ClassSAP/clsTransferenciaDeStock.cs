using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonFNR.BL
{
    class clsTransferenciaDeStock : ComportaForm
    {
        #region CONSTANTES

        private const int FRM_TRANSFERENCIA_DE_STOCK = 940;
        private const string OBJETO_TRANSFERNECIA_DE_STOCK = "1250000001";
        private const string GRID_ARTICULOS = "23";
        private const string COLUMNA_CLAVE_ARTICULO = "1";
        private const string COLUMNA_SERIE_INICIO = "U_SerieIni";
        private const string COLUMNA_SERIE_FIN = "U_SerieFin";
        private const string COLUMNA_SERIE = "U_Serie";
        private const string COLUMNA_IVERSION_INICIAL = "U_InvInicial";
        private const string COLUMNA_COMISION = "U_Comision";
        private const string COLUMNA_STATUS_DE_SOLICITUD = "U_StatusSolicitud";
        private const int CHAR_PRESS_ENTER = 13;              
        private const string BTN_CREAR = "1";

        //ENCABEZADO
        private const string LBL_SOCIO_NEGOCIO = "5";
        private const string LBL_NOMBRE_SOCIO_NEGOCIO = "8";
        private const string LBL_PERSONA_CONTACTO = "33";
        private const string LBL_DESTINATARIO = "10";
        private const string LBL_LISTA_DE_PRECIOS = "37";
        private const string LBL_PROMOTOR = "27";
        private const string TXT_FECHA_CONTABILIZACION = "14";
        private const string TXT_FECHA_DOCUMENTO = "16";
        private const string TXT_SOCIO_NEGOCIO = "3";
        private const string TXT_NOMBRE_SOCIO_NEGOCIO = "7";
        private const string TXT_PERSONA_CONTACTO = "31";
        private const string TXT_DESTINATARIO = "9";
        private const string TXT_PROMOTOR = "25";
        private const string TXT_SERIE_DOCUMENTO = "40";
        private const string CMB_LISTA_PRECIOS = "36";
        private const string TIPO_MOVIMIENTO = "U_TipoMov";
        private const string CMB_SHIP_TO_CODE = "254000004";

        //CAMPOS NUEVOS DE VENTANA
        private const string LBL_NOMBRE_PROMOTOR = "lblNomPro";
        private const string LBL_SERIE = "lblSerie";
        private const string TXT_NOMBRE_PROMOTOR = "NombreP";
        private const string TXT_SERIE = "Serie";
        private const string TXT_ALMACEN_ORIGEN = "18";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForm = null;
        private static bool _oTransferenciaDeStock = false;

        private SAPbouiCOM.StaticText _oLblNombrePromotor = null;
        private SAPbouiCOM.StaticText _oLblSerie = null;
        private SAPbouiCOM.EditText _oTxtNombrePromotor = null;
        private SAPbouiCOM.EditText _oTxtSerie = null;
        private SAPbouiCOM.EditText _oAlmacenOrigen = null;
        private SAPbouiCOM.EditText _oTxtPromotor = null;
        private SAPbouiCOM.ComboBox _oCmbTipoMovimiento = null;
        private SAPbouiCOM.ComboBox _oComboSerie = null;
        private SAPbobsCOM.Recordset _oRec = null;
        private SAPbouiCOM.EditText _oTxtFechaContabilizacion = null;
        private SAPbouiCOM.EditText _oTxtFechaDocumento = null;

        private static List<Datos> lDatos = new List<Datos>();
        private static Datos itemDatos = new Datos();
        private SAPbouiCOM.Matrix _oMatrixArticulos = null;

        private SAPbouiCOM.EditText oItemCode = null;
        private SAPbouiCOM.EditText oSerieInicio = null;
        private SAPbouiCOM.EditText oSerieFin = null;
        private SAPbouiCOM.EditText oSerie = null;
        private SAPbouiCOM.EditText oInvInicial = null;
        private SAPbouiCOM.EditText oComision = null;
        private SAPbouiCOM.Button oBtnCrearSAP = null;

        private int _oContadorFormas = 0;
        private string TipoMovimiento = null;
        private bool PresionoBotonCrear = false;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de la transferencia de stock
        /// </summary>
        /// <param name="_Application">Objeto de la conexión de SAP</param>
        /// <param name="_Company">Objeto de la empresa</param>
        /// <param name="form">Nombre de la forma</param>
        public clsTransferenciaDeStock(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oTransferenciaDeStock == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                setEventos();
                _oTransferenciaDeStock = true;
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
                if (pVal.BeforeAction == false && pVal.FormType == FRM_TRANSFERENCIA_DE_STOCK)
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
                            _Application.StatusBarEvent -= new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(SBO_Application_StatusBarEvent);
                            Dispose();
                            application = null;
                            company = null;
                            _oTransferenciaDeStock = false;
                            Addon.typeList.RemoveAll(p => p._forma == formID);
                            return;
                        }
                        else
                        {
                            _oContadorFormas -= 1;
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_ACTIVATE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        CrearCamposDeUsuario(_oForm);
                    }                  
                }

                if (pVal.BeforeAction == true && pVal.FormType == FRM_TRANSFERENCIA_DE_STOCK)
                {                   

                    if (pVal.ItemUID == TXT_SERIE && pVal.CharPressed == CHAR_PRESS_ENTER && pVal.EventType == BoEventTypes.et_KEY_DOWN)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        _oTxtSerie = _oForm.Items.Item(TXT_SERIE).Specific;
                        _oAlmacenOrigen = _oForm.Items.Item(TXT_ALMACEN_ORIGEN).Specific;
                        _oForm.Select();
                        if (!string.IsNullOrEmpty(_oTxtSerie.Value.ToString()))
                        {
                            _oForm.Select();
                            AgregarArticulo(_oForm, _oTxtSerie.Value.ToString(), _oAlmacenOrigen.Value.ToString());
                            _oForm.Select();
                        }
                        else
                        {
                            _Application.MessageBox("Capture el número de serie");
                        }
                        bubbleEvent = false;
                        return;
                    }

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

                    if (pVal.EventType == BoEventTypes.et_KEY_DOWN && pVal.ItemUID == GRID_ARTICULOS && pVal.ColUID == COLUMNA_SERIE && pVal.CharPressed == CHAR_PRESS_ENTER)
                    {
                        bubbleEvent = false;
                        return;
                    }

                    if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == BTN_CREAR && pVal.InnerEvent == true)
                    {
                        bubbleEvent = false;
                        return;
                    }

                    if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == BTN_CREAR && pVal.InnerEvent == false)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        oBtnCrearSAP = _oForm.Items.Item(BTN_CREAR).Specific;

                        if (oBtnCrearSAP.Caption == "Crear")
                        {
                            string msgError = ValidarStatusSolicitudGrid(_oForm);

                            if (!string.IsNullOrEmpty(msgError))
                            {
                                _Application.MessageBox(msgError);
                                bubbleEvent = false;
                                bubbleEvent = false;
                                return;
                            }
                        }
                    }
                }

                if (pVal.BeforeAction == true && pVal.FormType == -FRM_TRANSFERENCIA_DE_STOCK)
                {
                    if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == TIPO_MOVIMIENTO)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        _oCmbTipoMovimiento = _oForm.Items.Item(TIPO_MOVIMIENTO).Specific;

                        if (!string.IsNullOrEmpty(_oCmbTipoMovimiento.Value.ToString()))
                        {
                            if (_oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                if (_Application.MessageBox("Estas seguro de cambiar el tipo de movimiento sin crear." + Environment.NewLine + "¿Desea continuar?", 2, "Si", "No") == 1)
                                {
                                    try
                                    {
                                        _oTxtPromotor.Value = "";
                                    }
                                    catch (Exception)
                                    {
                                    }                                    
                                }
                                else
                                {
                                    bubbleEvent = false;
                                }
                            }
                        }
                    }
                }



                if (pVal.BeforeAction == false && pVal.FormType == -FRM_TRANSFERENCIA_DE_STOCK)
                {
                    if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == TIPO_MOVIMIENTO)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        _oForm.Freeze(true);
                        SAPbouiCOM.Form F1 = _Application.Forms.GetFormByTypeAndCount(Convert.ToInt32(_oForm.TypeEx.TrimStart('-')), _oForm.TypeCount);
                        F1.Freeze(true);
                        _oTxtPromotor = F1.Items.Item(TXT_PROMOTOR).Specific;
                        _oTxtNombrePromotor = F1.Items.Item(TXT_NOMBRE_PROMOTOR).Specific;
                        _oTxtSerie = F1.Items.Item(TXT_SERIE).Specific;
                        _oCmbTipoMovimiento = _oForm.Items.Item(TIPO_MOVIMIENTO).Specific;


                        if (!string.IsNullOrEmpty(_oCmbTipoMovimiento.Value.ToString()))
                        {
                            F1.Items.Item(TXT_SERIE).Enabled = true;
                            TipoMovimiento = _oCmbTipoMovimiento.Value.ToString().TrimEnd(' ');
                            _oTxtFechaContabilizacion = F1.Items.Item(TXT_FECHA_CONTABILIZACION).Specific;
                            _oTxtFechaDocumento = F1.Items.Item(TXT_FECHA_DOCUMENTO).Specific;

                            if (TipoMovimiento != "PROMOTORES - OFICINAS" && TipoMovimiento != "OFICINAS - PROMOTORES")
                            {
                                if (_oForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    _oTxtPromotor.Value = Extensor.ObtenerSecretaria(_Company, "U_codigo_secretaria");
                                    _oTxtNombrePromotor.Value = Extensor.ObtenerSecretaria(_Company, "T0.U_nombre_secretaria");
                                }
                            }
                            else
                            {
                                if (_oForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    _oTxtPromotor.Value = "";
                                    _oTxtNombrePromotor.Value = "";
                                }
                            }

                            if (TipoMovimiento == "PROMOTORES - OFICINAS" || TipoMovimiento == "OFICINAS - PROMOTORES")
                            {
                                if (_oForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    //if (Extensor.ValidarImpresionCorteSolicitudes(_Company))
                                    //{
                                    if (F1.Items.Item(TXT_FECHA_DOCUMENTO).Enabled == true)
                                    {
                                        _oTxtFechaContabilizacion.Value = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                                        _oTxtFechaDocumento.Value = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                                        F1.Items.Item(TXT_FECHA_CONTABILIZACION).Enabled = false;
                                        F1.Items.Item(TXT_PROMOTOR).Click();
                                        F1.Items.Item(TXT_FECHA_DOCUMENTO).Enabled = false;
                                    }
                                    //}
                                    //else
                                    //{
                                    //    if (F1.Items.Item(TXT_FECHA_DOCUMENTO).Enabled == true)
                                    //    {
                                    //        DateTime hoy = DateTime.Now;
                                    //        DateTime mañana = hoy.AddDays(1);
                                    //        _oTxtFechaContabilizacion.Value = hoy.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString();
                                    //        _oTxtFechaDocumento.Value = mañana.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString();
                                    //        F1.Items.Item(TXT_FECHA_CONTABILIZACION).Enabled = false;
                                    //        F1.Items.Item(TXT_PROMOTOR).Click();
                                    //        F1.Items.Item(TXT_FECHA_DOCUMENTO).Enabled = false;
                                    //    }
                                    //}
                                }
                            }
                            else
                            {
                                if (_oForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    F1.Items.Item(TXT_PROMOTOR).Click();
                                    F1.Items.Item(TXT_FECHA_CONTABILIZACION).Enabled = true;
                                    _oTxtFechaContabilizacion.Value = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                                    F1.Items.Item(TXT_PROMOTOR).Click();
                                    F1.Items.Item(TXT_FECHA_DOCUMENTO).Enabled = true;
                                    _oTxtFechaDocumento.Value = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                                }
                            }
                        }

                        _oForm.Freeze(false);
                        F1.Freeze(false);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en método 'eventos' *clsTransferenciaDeStock* : " + ex.Message);
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
                        if (_Application.Forms.ActiveForm.TypeEx == FRM_TRANSFERENCIA_DE_STOCK.ToString())
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
                                        oSerieInicio = (SAPbouiCOM.EditText)_oMatrixArticulos.Columns.Item(COLUMNA_SERIE_INICIO).Cells.Item(noLinea).Specific;
                                        oSerieFin = (SAPbouiCOM.EditText)_oMatrixArticulos.Columns.Item(COLUMNA_SERIE_FIN).Cells.Item(noLinea).Specific;
                                        oSerie = (SAPbouiCOM.EditText)_oMatrixArticulos.Columns.Item(COLUMNA_SERIE).Cells.Item(noLinea).Specific;
                                        if (!string.IsNullOrEmpty(oSerieInicio.Value.ToString()) && !string.IsNullOrEmpty(oSerieFin.Value.ToString()))
                                        {
                                            itemDatos = new Datos();
                                            itemDatos.itemCode = oItemCode.Value.ToString();
                                            itemDatos.serieInial = oSerieInicio.Value.ToString();
                                            itemDatos.serieFinal = oSerieFin.Value.ToString();
                                            itemDatos.noLinea = x;
                                            lDatos.Add(itemDatos);
                                            x += 1;
                                        }
                                        else if (!string.IsNullOrEmpty(oSerie.Value.ToString()))
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
                _Application.MessageBox("Error en StatusBarEvent *clsTransferenciaDeStock* : " + ex.Message);
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

                //COMBOS
                oItem = _oForm.Items.Item(CMB_LISTA_PRECIOS);
                oItem.Visible = false;

                oItem = _oForm.Items.Item(CMB_SHIP_TO_CODE);
                oItem.Visible = false;

                //CAMPOS PROMOTOR
                oItem = _oForm.Items.Item(LBL_PROMOTOR);
                oItem.Top = 10;
                oItem.Left = 5;
                oItem.Width = 70;
                oItem.TextStyle = 0;

                oItemOficina = _oForm.Items.Item(TXT_PROMOTOR);
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
                    string s = _oForma.Items.Item(TXT_NOMBRE_PROMOTOR).UniqueID;
                }
                catch (Exception)
                {
                    _oForma.Freeze(true);

                    //Label 'Nombre promotor' ligado al campo de Label 'promotor'.
                    SAPbouiCOM.Item _olblPro = null;
                    _olblPro = _oForma.Items.Item(LBL_PROMOTOR);
                    newItem = _oForma.Items.Add(LBL_NOMBRE_PROMOTOR, BoFormItemTypes.it_STATIC);
                    newItem.Left = _olblPro.Left;
                    newItem.Top = _olblPro.Top + 18;
                    newItem.Width = 70;
                    newItem.ToPane = 0;
                    _oLblNombrePromotor = newItem.Specific;
                    _oLblNombrePromotor.Caption = "Nombre";

                    //Label 'Label serie' ligado al campo de Label 'Nombre promotor'.
                    SAPbouiCOM.Item _oLblNombreProm = null;
                    _oLblNombreProm = _oForma.Items.Item(LBL_NOMBRE_PROMOTOR);
                    newItem = _oForma.Items.Add(LBL_SERIE, BoFormItemTypes.it_STATIC);
                    newItem.Left = _oLblNombreProm.Left;
                    newItem.Top = _oLblNombreProm.Top + 20;
                    newItem.Width = 70;
                    newItem.ToPane = 0;
                    _oLblSerie = newItem.Specific;
                    _oLblSerie.Caption = "Serie";

                    //Campo texto 'Nombre promotor' ligado al campo de Label 'Nombre promotor'.
                    SAPbouiCOM.Item _oLblNP = null;
                    _oLblNP = _oForma.Items.Item(LBL_NOMBRE_PROMOTOR);
                    newItem = _oForma.Items.Add(TXT_NOMBRE_PROMOTOR, BoFormItemTypes.it_EDIT);
                    newItem.Left = _oLblNP.Left + 80;
                    newItem.Top = _oLblNP.Top;
                    newItem.Width = 141;
                    newItem.Height = 15;
                    newItem.ToPane = 0;
                    _oTxtNombrePromotor = (SAPbouiCOM.EditText)newItem.Specific;
                    _oTxtNombrePromotor.DataBind.SetBound(true, "OWTR", "U_NombreP");
                    _oLblNP.LinkTo = newItem.UniqueID;

                    //Campo texto 'Serie' ligado al campo de Label 'Serie'.
                    SAPbouiCOM.Item _oLblSer = null;
                    _oLblSer = _oForma.Items.Item(LBL_SERIE);
                    newItem = _oForma.Items.Add(TXT_SERIE, BoFormItemTypes.it_EDIT);
                    newItem.Left = _oLblSer.Left + 80;
                    newItem.Top = _oLblSer.Top;
                    newItem.Width = 141;
                    newItem.Height = 15;
                    newItem.ToPane = 0;
                    _oTxtSerie = (SAPbouiCOM.EditText)newItem.Specific;
                    _oLblSer.LinkTo = newItem.UniqueID;

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

        /// <summary>
        /// Agrega el artículo que corresponde al número de serie
        /// </summary>
        /// <param name="_oForm">Forma activa</param>
        /// <param name="numeroSerie">Número de serie capturado</param>
        private void AgregarArticulo(Form _oForm, string numeroSerie, string almacenOrigen)
        {
            try
            {
                bool existe = false;
                _oForm.Freeze(true);
                _oForm.Select();
                string msgError = ValidarEstatusSolicitud(numeroSerie, _oForm);

                if (string.IsNullOrEmpty(msgError))
                {
                    _oForm.Select();
                    _oRec = null;
                    _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    _oRec.DoQuery(@"SELECT  T0.ItemCode
                                    FROM    dbo.OSRN T0
                                            INNER JOIN dbo.OSRI T1 ON T1.SysSerial = T0.SysNumber
                                                                    AND T1.IntrSerial = T0.DistNumber 
                                                                    AND T1.Status = 0
                                    WHERE   T0.DistNumber = '" + numeroSerie + "' " +
                                           " AND T1.WhsCode = '" + almacenOrigen + "'");
                    _oForm.Select();
                    if (!string.IsNullOrEmpty(_oRec.Fields.Item("ItemCode").Value.ToString()))
                    {
                        _oMatrixArticulos = _oForm.Items.Item(GRID_ARTICULOS).Specific;
                        _oForm.Select();
                        for (int noLinea = 1; noLinea <= _oMatrixArticulos.RowCount; noLinea++)
                        {
                            _oForm.Select();
                            oSerieInicio = (SAPbouiCOM.EditText)_oMatrixArticulos.Columns.Item(COLUMNA_SERIE).Cells.Item(noLinea).Specific;
                            if (numeroSerie == oSerieInicio.Value.ToString())
                            {
                                _oForm.Select();
                                existe = true;
                                _Application.MessageBox("La serie ya se encuentra registrada");
                                _oTxtSerie = _oForm.Items.Item(TXT_SERIE).Specific;
                                _oTxtSerie.Value = "";
                                _oTxtSerie.Active = true;
                                _oForm.Select();
                            }
                        }

                        if (existe == false)
                        {
                            if (_oForm.Items.Item(GRID_ARTICULOS).Enabled == true)
                            {
                                _oForm.Select();
                                int linea = _oMatrixArticulos.RowCount;
                                _oMatrixArticulos.Columns.Item(COLUMNA_CLAVE_ARTICULO).Cells.Item(linea).Specific.Value = _oRec.Fields.Item("ItemCode").Value.ToString();
                                _oForm.Select();
                                _oMatrixArticulos.Columns.Item(COLUMNA_SERIE).Cells.Item(linea).Specific.Value = numeroSerie;
                                _oForm.Select();
                                _oTxtSerie = _oForm.Items.Item(TXT_SERIE).Specific;
                                _oTxtSerie.Value = "";
                                _oTxtSerie.Active = true;
                                _oForm.Select();                              
                            }
                            else
                            {
                                _oForm.Select();
                                _Application.MessageBox("La transferencia ya esta creada, no se pueden agregar mas series");
                            }
                        }
                    }
                    else
                    {
                        _oForm.Select();
                        _Application.MessageBox("No se encontró el número de serie: " + numeroSerie + " o no corresponde al almacén origen");
                        _oTxtSerie = _oForm.Items.Item(TXT_SERIE).Specific;
                        _oTxtSerie.Value = "";
                        _oTxtSerie.Active = true;
                    }
                }
                else
                {
                    _oForm.Select();
                    _Application.MessageBox(msgError);
                    _oTxtSerie = _oForm.Items.Item(TXT_SERIE).Specific;
                    _oTxtSerie.Value = "";
                    _oTxtSerie.Active = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al agregar el artículo *AgregarArticulo* : " + ex.Message);
            }
            finally
            {
                try
                {
                    if (_oRec != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oRec);
                    _oForm.Freeze(false);
                }
                catch (Exception)
                {
                }
             
            }
        }



        /// <summary>
        /// Valida si la serie se encuentra con estatus 
        /// C - Cancelado
        /// A - Atraco
        /// E - Extravió
        /// N - Erróneo        
        /// </summary>
        /// <param name="numeroSerie">Numero de la serie del contrato</param>
        /// <returns>Mensaje de error</returns>
        private string ValidarEstatusSolicitud(string numeroSerie, Form _oForm)
        {
            try
            {
                _oForm.Select();
                string msgError = null;
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT TOP 1
                                        U_StatusSolicitud
                                FROM    dbo.WTR1
                                WHERE   U_Serie = '" + numeroSerie + "' " +
                                "ORDER BY DocEntry DESC");

                _oForm.Select();
                if (_oRec.Fields.Item("U_StatusSolicitud").Value.ToString().Contains("N"))
                {
                    return msgError = "No se puede agregar la solicitud: " + numeroSerie + Environment.NewLine + " Estatus: ERRONEO";
                }
                _oForm.Select();
                if (_oRec.Fields.Item("U_StatusSolicitud").Value.ToString().Contains("C"))
                {
                    return msgError = "No se puede agregar la solicitud: " + numeroSerie + Environment.NewLine + " Estatus: CANCELADO";
                }
                _oForm.Select();
                if (_oRec.Fields.Item("U_StatusSolicitud").Value.ToString().Contains("A"))
                {
                    return msgError = "No se puede agregar la solicitud: " + numeroSerie + Environment.NewLine + " Estatus: ATRACO";
                }
                _oForm.Select();
                if (_oRec.Fields.Item("U_StatusSolicitud").Value.ToString().Contains("E"))
                {
                    return msgError = "No se puede agregar la solicitud: " + numeroSerie + Environment.NewLine + " Estatus: EXTRAVIO";
                }
                _oForm.Select();
                return "";
            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar estatus de la solicitud *ValidarEstatusSolicitud* : " + ex.Message);
            }
            finally
            {
                try
                {
                    if (_oRec != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oRec);
                }
                catch (Exception)
                {
                }
              
            }

        }

        /// <summary>
        /// Valida que los planes que cancelados tengan la inversión inicial en ceros
        /// </summary>
        /// <param name="_oForm">Forma activa</param>
        /// <returns>Mensaje de error</returns>
        private string ValidarStatusSolicitudGrid(Form _oForm)
        {
            try
            {
                string msgError = null;
                _oMatrixArticulos = _oForm.Items.Item(GRID_ARTICULOS).Specific;
                SAPbouiCOM.ComboBox _oCmbStatusSolicitud = null;
                SAPbouiCOM.EditText _oInversionIni = null;
                SAPbouiCOM.EditText _oSolicitud = null;

                for (int noLinea = 1; noLinea <= _oMatrixArticulos.RowCount; noLinea++)
                {
                    _oCmbStatusSolicitud = (SAPbouiCOM.ComboBox)_oMatrixArticulos.Columns.Item(COLUMNA_STATUS_DE_SOLICITUD).Cells.Item(noLinea).Specific;
                    _oInversionIni = (SAPbouiCOM.EditText)_oMatrixArticulos.Columns.Item(COLUMNA_IVERSION_INICIAL).Cells.Item(noLinea).Specific;
                    _oSolicitud = (SAPbouiCOM.EditText)_oMatrixArticulos.Columns.Item(COLUMNA_SERIE).Cells.Item(noLinea).Specific;

                    if (_oCmbStatusSolicitud.Selected.Value.Contains('N'))
                    {
                        if (Convert.ToDouble(_oInversionIni.Value) != 0)
                        {
                            return msgError = "Monto de la inversión inicial debe estar en ceros " + Environment.NewLine +
                                         "Solicitud: " + _oSolicitud.Value.ToString() + Environment.NewLine + " Estatus: ERRONEO";
                        }
                    }

                    if (_oCmbStatusSolicitud.Selected.Value.Contains('C'))
                    {
                        if (Convert.ToDouble(_oInversionIni.Value) != 0)
                        {
                            return msgError = "Monto de la inversión inicial deber estar en ceros " + Environment.NewLine +
                                         "Solicitud: " + _oSolicitud.Value.ToString() + Environment.NewLine + " Estatus: CANCELADO";
                        }
                    }

                    if (_oCmbStatusSolicitud.Selected.Value.Contains('A'))
                    {
                        if (Convert.ToDouble(_oInversionIni.Value) != 0)
                        {
                            return msgError = "Monto de la inversión inicial deber estar en ceros " + Environment.NewLine +
                                         "Solicitud: " + _oSolicitud.Value.ToString() + Environment.NewLine + " Estatus: ATRACO";
                        }
                    }

                    if (_oCmbStatusSolicitud.Selected.Value.Contains('E'))
                    {
                        if (Convert.ToDouble(_oInversionIni.Value) != 0)
                        {
                            return msgError = "Monto de la inversión inicial deber estar en ceros " + Environment.NewLine +
                                         "Solicitud: " + _oSolicitud.Value.ToString() + Environment.NewLine + " Estatus: EXTRAVIO";
                        }
                    }
                }
                return msgError;

            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar estatus de solicitudes del grid *ValidarStatusSolicitudGrid* : " + ex.Message);
            }
        }

        #endregion
    }
}
