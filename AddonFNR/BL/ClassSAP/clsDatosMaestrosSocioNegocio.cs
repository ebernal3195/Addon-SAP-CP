using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonFNR.BL
{
    class clsDatosMaestrosSocioNegocio : ComportaForm
    {
        #region CONSTANTES

        private const int FRM_DATOS_MAESTROS_SOCIO = 134;
        private const string OBJETO_DMSN = "2";
        private const string FRM_CALCULADORA = "10000076";
        private const string TXT_COLONIA = "178";
        private const string VENTANA_EMERGENTE = "0";
        private const string MENSAJE_CAMBIO_GRUPO = "1000001";

        #region CAMPOS DE USUARIO PESTAÑA GENERAL

        //LABELS
        private const string LBL_FECHA_NACIMIENTO = "lblFecNac";
        private const string LBL_DIA = "lblDia";
        private const string LBL_MES = "lblMes";
        private const string LBL_YEAR = "lblYear";
        private const string LBL_PROMOTOR_SOLICITUD = "lblProSol";
        private const string LBL_PROMOTOR_COMISIONA = "lblProCom";
        private const string LBL_PROMOTOR_DERECHO = "lblProDer";
        private const string LBL_OBSERVACIONES = "lblObser";
        private const string LBL_ADMON_RECIBE = "lblAdmRec";
        private const string LBL_ASISTENTE_RECIBE = "lblAsiRec";
        private const string LBL_FECHA_REGRESA_COPIA = "lblFecCop";
        private const string LBL_MOTIVO_SOCIO_INACTIVO = "lblMotivo";
        private const string LBL_FECHA_INACTIVO = "lblFecIna";
        private const string LBL_COMENTARIOS_MOTIVO = "lblComenM";
        private const string LBL_DIAS_SIN_ABONAR_RECIBO = "lblDateAR";
        private const string LBL_MONTO_ATRASADO = "lblMonAtr";
        private const string LBL_VENCIMIENTO_CONTRATO = "lblVenCon";
        private const string LBL_FECHA_PRIMER_ABONO = "lblFecPri";
        private const string LBL_FORMA_DE_PAGO = "lblFPago";
        private const string LBL_STATUS_SOLICITUD = "lblStaSol";
        private const string LBL_MONTO_DE_PAGO = "lblMPago";
        //private const string LBL_DIA_VISITA_1 = "lblVisit1";
        //private const string LBL_DIA_VISITA_2 = "lblVisit2";

        //EDITTEXTS
        private const string TXT_DIA = "U_Dia";
        private const string TXT_MES = "U_Mes";
        private const string TXT_YEAR = "U_Year";
        private const string TXT_PROMOTOR_SOLICITUD = "U_PromotorSolicitud"; //"txtProSol";
        private const string TXT_PROMOTOR_COMISIONA = "U_PromotorComisiona"; //"txtProCom";
        private const string TXT_PROMOTOR_DERECHO = "U_PromotorDerecho"; //"txtProDer";
        private const string TXT_OBSERVACIONES = "txtObser";
        private const string TXT_ADMON_RECIBE_FECHA = "txtAdmReF";
        private const string TXT_ADMON_RECIBE_NOMBRE = "txtAdmReN";
        private const string TXT_ASISTENTE_RECIBE_FECHA = "txtAsisF";
        private const string TXT_ASISTENTE_RECIBE_NOMBRE = "txtAsisN";
        private const string TXT_FECHA_REGRESA_COPIA = "txtFecCop";
        private const string TXT_MOTIVO_SOCIO_INACTIVO = "txtMotivo";
        private const string TXT_FECHA_INACTIVO = "txtFecIna";
        private const string TXT_COMENTARIOS_MOTIVO = "txtComenM";
        private const string TXT_DIAS_SIN_ABONAR_RECIBO = "txtDateAR";
        private const string TXT_MONTO_ATRASADO = "txtMonAtr";
        private const string TXT_VENCIMIENTO_CONTRATO = "txtVenCon";
        private const string TXT_FECHA_PRIMER_ABONO = "txtFecPri";
        private const string TXT_FECHA_CREACION = "U_FechaCreacion";
        private const string TXT_FORMA_DE_PAGO = "txtFPago";
        private const string TXT_STATUS_SOLICITUD = "U_StatusSolicitud"; //"txtStaSol";
        private const string TXT_MONTO_DE_PAGO = "txtMPago";
        private const string TXT_ESTADO_CIVIL = "U_EstadoCivil";
        //private const string TXT_DIA_VISITA_1 = "txtVisit1";
        //private const string TXT_DIA_VISITA_2 = "txtVisit2";

        #endregion

        #region CAMPOS DE SAP ENCABEZADO

        //LABELS
        private const string LBL_NOMBRE_EXTRANJERO = "129";
        private const string LBL_PEDIDOS = "32";
        private const string LBL_OPORTUNIDADES = "148";

        //EDITTEXTS
        private const string TXT_NOMBRE_EXTRANJERO = "128";
        private const string TXT_PEDIDOS = "31";
        private const string TXT_OPORTUNIDADES = "149";
        private const string TXT_NAME_PROVEEDOR = "7";
        private const string TXT_RFC = "41";
        private const string TXT_CARDCODE = "5";

        //COMBOBOX
        private const string CMB_MONEDA = "38";
        private const string CMB_GRUPO = "16";
        private const string CMB_CLASE_SN = "40";
        private const string CMB_SERIE_FOLIO_SAP = "1320002080";

        #endregion

        #region CAMPOS DE SAP FOLDER GENERAL

        //LABELS
        private const string LBL_DESDE = "10002052";
        private const string LBL_HASTA = "10002053";
        private const string LBL_COMENTARIOS_INACTIVO = "10002048";
        private const string LBL_TIPO_SOCIO_NEGOCIO = "358";
        private const string LBL_TEL1 = "44";
        private const string LBL_TEL2 = "46";
        private const string LBL_TEL_MOVIL = "48";
        private const string LBL_FAX = "50";
        private const string LBL_SITIO_WEB = "395";
        private const string LBL_CLASE_EXPEDICION = "101";
        private const string LBL_CLAVE_ACCESO = "184";
        private const string LBL_INDICADOR_FACTORIN = "22";
        private const string LBL_PROYECTO_SN = "223";
        private const string LBL_RAMO = "350001034";
        private const string LBL_NOMBRE_ALIAS = "2013";
        private const string LBL_PERSONA_CONTACTO = "118";
        private const string LBL_CURP = "365";
        private const string LBL_ID_FISCAL_FEDERAL = "62";
        private const string LBL_COMENTARIOS = "109";
        private const string LBL_CODIGO_CANAL = "333";
        private const string LBL_COBRADOR_ASIGNADO = "336";
        private const string LBL_TERRITORIO = "343";
        private const string LBL_FORMATO_EXPORTACION = "480002076";
        private const string LBL_ID_NUMERO_2 = "114";
        private const string LBL_CORREO_ELECTRONICO = "61";
        private const string LBL_PROMOTOR = "59";

        //EDITTEXTS       
        private const string TXT_DESDE = "10002051";
        private const string TXT_HASTA = "10002054";
        private const string TXT_COMENTARIOS_INACTIVO = "10002047";
        private const string TXT_TEL1 = "43";
        private const string TXT_TEL2 = "45";
        private const string TXT_TEL_MOVIL = "51";
        private const string TXT_FAX = "47";
        private const string TXT_SITIO_WEB = "394";
        private const string TXT_CLAVE_ACCESO = "185";
        private const string TXT_INDICADOR_FACTORIN_1 = "39";
        private const string TXT_INDICADOR_FACTORIN_2 = "49";
        private const string TXT_PROYECTO_SN = "222";
        private const string TXT_NOMBRE_ALIAS = "2014";
        private const string TXT_PERSONA_CONTACTO = "117";
        private const string TXT_CURP = "113";
        private const string TXT_ID_FISCAL_FEDERAL = "73";
        private const string TXT_COMENTARIOS = "108";
        private const string TXT_COMISION = "57";
        private const string TXT_CODIGO_CANAL = "335";
        private const string TXT_COBRADOR_ASIGNADO = "338";
        private const string TXT_TERRITORIO = "345";

        //CHECKS
        private const string CHK_VALIDO = "10002044";
        private const string CHK_INACTIVO = "10002045";
        private const string CHK_AVANZADO = "10002046";

        //COMBOBOXS
        private const string CMB_CLASE_EXPEDICION = "102";
        private const string CMB_RAMO = "350001035";
        private const string CMB_COMISION_DEFINIDA = "55";
        private const string CMB_FORMA_EXPORTACION = "480002077";

        //ICONOS
        private const string ICO_CLASE_EXPEDICION = "111";

        #endregion

        #region BOTONES NATIVOS DE SAP

        private const string BTN_CREAR_SAP = "1282";
        private const string BTN_BUSCAR_SAP = "1281";
        private const string BTN_BUSCAR_VENTANA = "1";
        private const string BTN_CANCELAR_DATOS_MAESTROS = "2";

        #endregion

        #region BOTONES DE USUARIO

        private const string BTN_VALIDAR_SOLICITUD = "btnValSol";
        //private const string BTN_GENERAR_CODIGO_ACTIVACION = "btnCodAct";

        #endregion

        #region FOLDERS A OCULTAR

        private const string FLD_EJECUCION_PAGO = "214";
        private const string FLD_PROPIEDADES = "10";

        #endregion

        #region CAMPOS DEFINIDOS POR USUARIO

        private const string CDU_TRASPASOREL = "U_TraspasoRel";
        private const string CDU_SOLICITUD = "U_Solicitud";
        private const string CDU_SOLICITUD_INTERNA = "U_SolicitudInt";
        private const string CDU_CODIGO_PLAN_PREVISION = "U_NumArt_";
        private const string CDU_NOMBRE_PLAN_PREVISION = "U_Dsciption";
        private const string CDU_INVINICIAL = "U_InvIncial";
        private const string CDU_COMISION = "U_Comision";
        private const string CDU_PAPELERIA = "U_Papeleria";
        private const string CDU_IMPORTE_RECIBIDO = "U_Importe_Recibido";
        private const string CDU_EXCEDENTE_INVINI = "U_Excedente_InvIni";
        private const string CDU_BONO = "U_Bono";
        private const string CDU_CODCOB = "U_CodCob";
        private const string CDU_NOMCOB = "U_NomCob";
        private const string CDU_CMB_CATEGORIA = "9";
        private const string CDU_BENEFICIARIO_PAGO_RECIBIDO = "U_BeneficiarioPagoRe";
        private const string CDU_CODIGO_ASISTENTE = "U_CodigoPromotor";
        private const string CDU_NOMBRE_ASISTENTE = "U_NomProm";
        private const string CDU_PREFIJO_PLAN = "U_PrefijoPlan";
        private const string CDU_ESQUEMA_PAGO = "U_Esquema_pago";

        #endregion

        #endregion

        #region ENUMERADOS

        /// <summary>
        /// Enumerado que se ejecuta al momento de presionar click derecho
        /// </summary>
        enum ClickDerecho
        {
            Pegar = 773,
            Cortar = 771,
            Borrar = 774
        }

        #endregion

        #region VARIABLES

        private string userFielsType = "";
        private string FormType = "";
        private List<string> userFields = Extensor.Configuracion.SOCIOSNEGOCIOS.CamposSocioNegocios.Split(',').ToList();
        private List<string> formTypes = new List<string> { "-134" };
        private static bool _oDatosMaestrosSociosNegocio = false;
        private static List<Datos> lDatosContratos = null;
        private string[] camposArr = { TXT_PROMOTOR_SOLICITUD, TXT_PROMOTOR_COMISIONA, TXT_PROMOTOR_DERECHO, TXT_OBSERVACIONES, TXT_ADMON_RECIBE_FECHA, TXT_ADMON_RECIBE_NOMBRE,
                                        TXT_ASISTENTE_RECIBE_FECHA, TXT_ASISTENTE_RECIBE_NOMBRE,TXT_FECHA_REGRESA_COPIA,TXT_MOTIVO_SOCIO_INACTIVO,TXT_COMENTARIOS_MOTIVO,
                                        TXT_DIAS_SIN_ABONAR_RECIBO,TXT_MONTO_ATRASADO,TXT_VENCIMIENTO_CONTRATO,TXT_FORMA_DE_PAGO,TXT_MONTO_DE_PAGO,
                                        TXT_DIA, TXT_MES, TXT_YEAR, TXT_FECHA_INACTIVO, TXT_FECHA_PRIMER_ABONO, TXT_STATUS_SOLICITUD };

        private SAPbouiCOM.Form _oForm = null;

        private SAPbouiCOM.EditText _oNombreSocioNegocio = null;
        private SAPbouiCOM.EditText _oRFCSocioNegocio = null;
        private SAPbouiCOM.ComboBox _oDia = null;
        private SAPbouiCOM.ComboBox _oMes = null;
        private SAPbouiCOM.ComboBox _oYear = null;
        private SAPbouiCOM.EditText _oColonia = null;

        private SAPbouiCOM.StaticText _oLblFechaNacimiento = null;
        private SAPbouiCOM.StaticText _oLblDia = null;
        private SAPbouiCOM.StaticText _oLblMes = null;
        private SAPbouiCOM.StaticText _oLblYear = null;
        private SAPbouiCOM.StaticText _oLblObservaciones = null;
        private SAPbouiCOM.StaticText _oLblPromotorSolicitud = null;
        private SAPbouiCOM.StaticText _oLblPromotorComisiona = null;
        private SAPbouiCOM.StaticText _oLblPromotorDerecho = null;
        private SAPbouiCOM.StaticText _oLblAdmonRecibe = null;
        private SAPbouiCOM.StaticText _oLblMotivo_Inactivo = null;
        private SAPbouiCOM.StaticText _oLblFecha_Inactivo = null;
        private SAPbouiCOM.StaticText _oLblComentarios_Inactivo = null;
        private SAPbouiCOM.StaticText _oLblDiasSinAbonar = null;
        private SAPbouiCOM.StaticText _oLblMontoAtrasado = null;
        private SAPbouiCOM.StaticText _oLblVencimientoContrato = null;
        private SAPbouiCOM.StaticText _oLblFechaPrimerAbono = null;
        private SAPbouiCOM.StaticText _oLblFormaPago = null;
        private SAPbouiCOM.StaticText _oLblStatusSolicitud = null;
        private SAPbouiCOM.StaticText _oLblMontoPago = null;
        private SAPbouiCOM.StaticText _oLblVisita_1 = null;
        private SAPbouiCOM.StaticText _oLblVisita_2 = null;

        private SAPbouiCOM.ComboBox _oTxtDia = null;
        private SAPbouiCOM.ComboBox _oTxtMes = null;
        private SAPbouiCOM.ComboBox _oTxtYear = null;
        private SAPbouiCOM.EditText _oTxtPromotorSolicitud = null;
        private SAPbouiCOM.EditText _oTxtPromotorComisiona = null;
        private SAPbouiCOM.EditText _oTxtPromotorDerecho = null;
        private SAPbouiCOM.EditText _oTxtObservaciones = null;
        private SAPbouiCOM.EditText _oTxtCardCode = null;

        private SAPbouiCOM.EditText _oTxtAdmonRecibeFecha = null;
        private SAPbouiCOM.EditText _oTxtAdmonRecibeNombre = null;
        private SAPbouiCOM.EditText _oTxtAsistenteRecibeFecha = null;
        private SAPbouiCOM.EditText _oTxtAsistenteRecibeNombre = null;
        private SAPbouiCOM.EditText _oTxtFechaRegresaCopiaAsistente = null;
        private SAPbouiCOM.EditText _oTxtMotivo_Inactivo = null;
        private SAPbouiCOM.EditText _oTxtFecha_Inactivo = null;
        private SAPbouiCOM.EditText _oTxtComentarios_Inactivo = null;
        private SAPbouiCOM.EditText _oTxtDiasSinAbonar = null;
        private SAPbouiCOM.EditText _oTxtMontoAtrasado = null;
        private SAPbouiCOM.EditText _oTxtVencimientoContrato = null;
        private SAPbouiCOM.EditText _oTxtFechaPrimerAbono = null;
        private SAPbouiCOM.ComboBox _oCmbFormaPago = null;
        private SAPbouiCOM.ComboBox _oCmbStatusSolicitud = null;
        private SAPbouiCOM.EditText _oTxtMontoPago = null;
        private SAPbouiCOM.EditText _oTxtVisita_1 = null;
        private SAPbouiCOM.EditText _oTxtVisita_2 = null;

        private SAPbouiCOM.ComboBox _oCmbGrupo = null;
        private SAPbouiCOM.ComboBox _oCmbClaseSN = null;
        private SAPbouiCOM.ComboBox _oCodigoSerieSAP = null;
        private SAPbouiCOM.ComboBox _oCmbEsquemaPago = null;
        private SAPbobsCOM.Recordset _oRec = null;
        private int _oContadorFormas = 0;

        //CAMPOS DEFINIDOS POR EL USUARIO
        private SAPbouiCOM.EditText _oTxtTraspasoRel = null;
        //private SAPbouiCOM.EditText _oTxtSolicitud = null;
        private SAPbouiCOM.EditText _oTxtSolicitudInterna = null;
        private SAPbouiCOM.EditText _oTxtTrasp_CodPlan = null;
        private SAPbouiCOM.EditText _oTxtAct_Descripcion = null;
        private SAPbouiCOM.EditText _oTxtInvInicial = null;
        private SAPbouiCOM.EditText _oTxtComision = null;
        private SAPbouiCOM.EditText _oTxtPapaleria = null;
        private SAPbouiCOM.EditText _oTxtImporte_Recibido = null;
        private SAPbouiCOM.EditText _oTxtExcedente_InvIni = null;
        private SAPbouiCOM.EditText _oTxtBono = null;
        private SAPbouiCOM.EditText _oTxtCodCobrador = null;
        private SAPbouiCOM.EditText _oTxtNomCobrador = null;
        private SAPbouiCOM.ComboBox _oCmbCategoria = null;
        private SAPbouiCOM.EditText _oTxtBeneficiarioPagoRecibido = null;
        private SAPbouiCOM.EditText _oTxtCodigoAsisten = null;
        private SAPbouiCOM.EditText _oTxtNombreAsistente = null;
        private SAPbouiCOM.EditText _oTxtPrefijoPlan = null;
        private SAPbouiCOM.EditText _oTxtFechaCreacion = null;


        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de datos maestros de socio de negocio
        /// </summary>
        /// <param name="_Application">Objeto de la conexión de SAP</param>
        /// <param name="_Company">Objeto de la empresa</param>
        /// <param name="form">Nombre de la forma</param>
        public clsDatosMaestrosSocioNegocio(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oDatosMaestrosSociosNegocio == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                setEventos();
                _oDatosMaestrosSociosNegocio = true;
            }

            if (lDatos != null)
            {
                try
                {
                    lDatosContratos = new List<Datos>(lDatos);
                    _Application.ActivateMenuItem("2561");
                    lDatos.Clear();
                }
                catch (Exception)
                {
                }

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
                if (pVal.FormTypeEx == VENTANA_EMERGENTE && _Application.Forms.ActiveForm.TypeCount == 2 &&
                  pVal.EventType == BoEventTypes.et_FORM_ACTIVATE || pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == false)
                {
                    try
                    {
                        SAPbouiCOM.Form formaEmergente = null;
                        formaEmergente = _Application.Forms.GetForm(pVal.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);

                        if (formaEmergente.Title == "Datos maestros socio de negocios")
                        {
                            SAPbouiCOM.StaticText itemMensaje = formaEmergente.Items.Item(7).Specific;
                            var msgDescripcion = itemMensaje.Caption;
                            if (msgDescripcion == "¿Desea sustituir la lista de precios y el descuento en efectivo según el grupo de socios")
                            {
                                SAPbouiCOM.Button botonSi = formaEmergente.Items.Item("1").Specific;
                                botonSi.Item.Click();
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }

                }

                if (formTypes.Contains(pVal.FormTypeEx))
                {
                    FormType = pVal.FormTypeEx;
                }
                else
                {
                    FormType = "";
                }

                if (userFields.Contains(pVal.ItemUID))
                {
                    userFielsType = pVal.ItemUID;
                }
                else
                {
                    userFielsType = "";
                }

                if (pVal.BeforeAction == true && pVal.FormType == -FRM_DATOS_MAESTROS_SOCIO)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN || pVal.EventType == BoEventTypes.et_RIGHT_CLICK ||
                        pVal.EventType == BoEventTypes.et_DOUBLE_CLICK)
                    {

                        if (userFields.Contains(pVal.ItemUID))
                        {
                            _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                            if (pVal.ItemUID == "U_SolicitudInt" && _oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {

                            }
                            else
                                if (pVal.ItemUID.ToString() == "U_CodigoPromotor" || pVal.ItemUID.ToString() == "U_NomProm")
                                {
                                    if (_Company.UserName.ToString() == "gdlcon01" || _Company.UserName.ToString() == "gdlcon02" || _Company.UserName.ToString() == "tolcon01" || _Company.UserName.ToString() == "tolcon02")
                                    {
                                    }
                                    else
                                    {
                                        if (_oForm.Mode != BoFormMode.fm_FIND_MODE)
                                        {
                                            bubbleEvent = false;
                                        }
                                    }
                                }
                                else
                                    if (_oForm.Mode != BoFormMode.fm_FIND_MODE)
                                    {
                                        bubbleEvent = false;
                                    }
                        }

                        //if (pVal.ItemUID.ToString() == "U_CodigoPromotor" || pVal.ItemUID.ToString() == "U_NomProm")
                        //{
                        //    if (_Company.UserName.ToString() == "gdlcon01" || _Company.UserName.ToString() == "gdlcon02" || _Company.UserName.ToString() == "tolcon01" || _Company.UserName.ToString() == "tolcon02")
                        //    {
                        //    }
                        //    else
                        //    {
                        //        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        //        if (_oForm.Mode != BoFormMode.fm_FIND_MODE)
                        //        {
                        //            bubbleEvent = false;
                        //        }
                        //    }
                        //}
                        //else
                        //{
                        //    _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        //    if (_oForm.Mode != BoFormMode.fm_FIND_MODE)
                        //    {
                        //        bubbleEvent = false;
                        //    }
                        //}

                    }

                }

                if (pVal.FormTypeEx == FRM_CALCULADORA && pVal.EventType == BoEventTypes.et_FORM_LOAD)
                {
                    bubbleEvent = false;
                }

                if (pVal.BeforeAction == false && pVal.FormType == FRM_DATOS_MAESTROS_SOCIO)
                {


                    if (pVal.ItemUID == TXT_DIA && pVal.EventType == BoEventTypes.et_LOST_FOCUS)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        OcultarControlesVentana(_oForm);
                    }

                    if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == CMB_CLASE_SN)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        OcultarControlesVentana(_oForm);
                    }

                    if (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == TXT_RFC)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        _oNombreSocioNegocio = _oForm.Items.Item(TXT_NAME_PROVEEDOR).Specific;
                        _oRFCSocioNegocio = _oForm.Items.Item(TXT_RFC).Specific;
                        _oDia = _oForm.Items.Item(TXT_DIA).Specific;
                        _oMes = _oForm.Items.Item(TXT_MES).Specific;
                        _oYear = _oForm.Items.Item(TXT_YEAR).Specific;
                        _oCmbGrupo = _oForm.Items.Item(CMB_GRUPO).Specific;

                        string[] separarNombre = null;
                        string NombreMayusculas = null;
                        string RFCGenerado = null;

                        if (_oForm.Mode == BoFormMode.fm_ADD_MODE || _oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                            if (_oCmbGrupo.Selected.Description.ToString().Contains("PABS"))
                            {
                                if (!string.IsNullOrEmpty(_oNombreSocioNegocio.Value.ToString()))
                                {
                                    var nombreSeparado = SepararNombreApllidos_Fomato.FormatoTextoRFC(_oNombreSocioNegocio.Value.ToString());
                                    NombreMayusculas = SepararNombreApllidos_Fomato.FormatoMayusculas(_oNombreSocioNegocio.Value.ToString());
                                    _oNombreSocioNegocio.Value = NombreMayusculas;
                                    separarNombre = nombreSeparado.ToString().Split('@');

                                    string msgError = ValidarFechaNacimiento(_oDia, _oMes, _oYear);

                                    if (string.IsNullOrEmpty(msgError))
                                    {
                                        _oRFCSocioNegocio = _oForm.Items.Item(TXT_RFC).Specific;
                                        _oRFCSocioNegocio.Value = "";
                                        RFCGenerado = SepararNombreApllidos_Fomato.GenerarRFC(separarNombre, _oDia.Selected.Value, _oMes.Selected.Value, _oYear.Selected.Value);
                                        _oRFCSocioNegocio.Value = RFCGenerado;
                                    }
                                    else
                                    {
                                        _Application.MessageBox(msgError);
                                    }
                                }
                            }
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {

                        try
                        {
                            if (pVal.ItemUID == "uaf_0")
                            {
                                if (_oForm.Mode != BoFormMode.fm_FIND_MODE)
                                {
                                    _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                    bool checkValido = _oForm.Items.Item(CHK_VALIDO).Specific.Selected;
                                    if (checkValido == true)
                                    {
                                        _oForm.Items.Item(LBL_MOTIVO_SOCIO_INACTIVO).Visible = false;
                                        _oForm.Items.Item(TXT_MOTIVO_SOCIO_INACTIVO).Visible = false;
                                        _oForm.Items.Item(LBL_COMENTARIOS_MOTIVO).Visible = false;
                                        _oForm.Items.Item(TXT_COMENTARIOS_MOTIVO).Visible = false;
                                        _oForm.Items.Item(LBL_FECHA_INACTIVO).Visible = false;
                                        _oForm.Items.Item(TXT_FECHA_INACTIVO).Visible = false;
                                    }
                                    else
                                    {
                                        OcultarControlesVentana(_oForm);
                                        _oForm.Items.Item(LBL_MOTIVO_SOCIO_INACTIVO).Visible = true;
                                        _oForm.Items.Item(TXT_MOTIVO_SOCIO_INACTIVO).Visible = true;
                                        _oForm.Items.Item(LBL_COMENTARIOS_MOTIVO).Visible = true;
                                        _oForm.Items.Item(TXT_COMENTARIOS_MOTIVO).Visible = true;
                                        _oForm.Items.Item(LBL_FECHA_INACTIVO).Visible = true;
                                        _oForm.Items.Item(TXT_FECHA_INACTIVO).Visible = true;

                                        _oForm.Items.Item(LBL_DESDE).Left = 650;
                                        _oForm.Items.Item(TXT_DESDE).Left = 700;
                                        _oForm.Items.Item(LBL_HASTA).Left = 760;
                                        _oForm.Items.Item(TXT_HASTA).Left = 810;

                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }

                        if (pVal.ItemUID == BTN_VALIDAR_SOLICITUD)
                        {
                            _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                            if (_oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                SAPbouiCOM.Form _UDFForm = null;
                                string EmpresaSolicitud = null;
                                _UDFForm = _Application.Forms.GetForm("-" + pVal.FormType, _Application.Forms.ActiveForm.TypeCount);
                                _oTxtSolicitudInterna = _UDFForm.Items.Item(CDU_SOLICITUD_INTERNA).Specific;
                                _oTxtTrasp_CodPlan = _UDFForm.Items.Item(CDU_CODIGO_PLAN_PREVISION).Specific;
                                _oTxtAct_Descripcion = _UDFForm.Items.Item(CDU_NOMBRE_PLAN_PREVISION).Specific;
                                _oTxtCodigoAsisten = _UDFForm.Items.Item(CDU_CODIGO_ASISTENTE).Specific;
                                _oTxtNombreAsistente = _UDFForm.Items.Item(CDU_NOMBRE_ASISTENTE).Specific;
                                _oTxtPrefijoPlan = _UDFForm.Items.Item(CDU_PREFIJO_PLAN).Specific;
                                _oCodigoSerieSAP = _oForm.Items.Item(CMB_SERIE_FOLIO_SAP).Specific;
                                _oCmbGrupo = _oForm.Items.Item(CMB_GRUPO).Specific;
                                _oCmbClaseSN = _oForm.Items.Item(CMB_CLASE_SN).Specific;

                                if (!string.IsNullOrEmpty(_oTxtSolicitudInterna.Value.ToString()))
                                {
                                    if (Extensor.ValidarSiExisteSolicitud(_oTxtSolicitudInterna.Value.ToString(), _Company))
                                    {
                                        Extensor.DatosSolicitud DatosSolicitud = new Extensor.DatosSolicitud();
                                        DatosSolicitud = Extensor.ObtenerDatosSolicitud(_oTxtSolicitudInterna.Value.ToString(), _Company);

                                        if (!string.IsNullOrEmpty(DatosSolicitud.prefijoPlan.ToString()))
                                        {
                                            if (_oCmbClaseSN.Selected.Value == "L")
                                            {
                                                if (_oCmbGrupo.Selected.Description.Contains("PABS"))
                                                {
                                                    _oTxtTrasp_CodPlan.Value = DatosSolicitud.codigoPlan;
                                                    _oTxtAct_Descripcion.Value = DatosSolicitud.nombrePlan;
                                                    _oTxtCodigoAsisten.Value = DatosSolicitud.codigoAsistente;
                                                    _oTxtNombreAsistente.Value = DatosSolicitud.nombreAsistente;
                                                    _oTxtPrefijoPlan.Value = DatosSolicitud.prefijoPlan;
                                                    EmpresaSolicitud = DatosSolicitud.empresa;

                                                    if (EmpresaSolicitud == "APOYO")
                                                    {
                                                        _oCodigoSerieSAP.Select("Apoyo", BoSearchKey.psk_ByValue);
                                                    }
                                                    else if (EmpresaSolicitud == "COOPERATIVA")
                                                    {
                                                        _oCodigoSerieSAP.Select("Cooper", BoSearchKey.psk_ByValue);
                                                    }
                                                }
                                                else
                                                {
                                                    _Application.MessageBox("Seleccione el grupo PABS");
                                                }
                                            }
                                            else
                                            {
                                                _Application.MessageBox("Seleccione LEAD");
                                            }
                                        }
                                        else
                                        {
                                            _Application.MessageBox("No existe solicitud");
                                        }
                                        DatosSolicitud = null;
                                    }
                                    else
                                    {
                                        _Application.MessageBox("La solicitud ya se encuentra registrada");
                                    }

                                }
                                else
                                {
                                    _Application.MessageBox("Capture el número de solicitud");
                                    return;
                                }
                            }
                        }

                        if (pVal.ItemUID == BTN_BUSCAR_VENTANA)
                        {
                            _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                            SAPbouiCOM.Button boton = _oForm.Items.Item(BTN_BUSCAR_VENTANA).Specific;
                            if (boton.Caption == "Actualizar")
                            {
                                OcultarControlesVentana(_oForm);
                            }

                            if (boton.Caption == "OK")
                            {
                                OcultarControlesVentana(_oForm);
                            }
                        }

                        if (pVal.ItemUID == "18")
                        {
                            _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                            OcultarControlesVentana(_oForm);
                            if (_oForm.Mode == BoFormMode.fm_VIEW_MODE)
                            {
                                _oForm.Freeze(true);
                                InhabilitarControlesDefinipos(_oForm);
                                _oForm.Freeze(false);
                            }
                            _oForm.Freeze(false);
                            _oForm.Update();
                            _oForm.Refresh();
                        }

                        if (pVal.ItemUID == CHK_INACTIVO)
                        {
                            if (_oForm.Mode != BoFormMode.fm_FIND_MODE && _oForm.Mode != BoFormMode.fm_VIEW_MODE)
                            {
                                _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                OcultarControlesVentana(_oForm);
                                _oForm.Items.Item(TXT_NAME_PROVEEDOR).Click();
                                _oForm.Items.Item(LBL_MOTIVO_SOCIO_INACTIVO).Visible = true;
                                _oForm.Items.Item(TXT_MOTIVO_SOCIO_INACTIVO).Visible = true;
                                _oForm.Items.Item(LBL_COMENTARIOS_MOTIVO).Visible = true;
                                _oForm.Items.Item(TXT_COMENTARIOS_MOTIVO).Visible = true;
                                _oForm.Items.Item(LBL_FECHA_INACTIVO).Visible = true;
                                _oForm.Items.Item(TXT_FECHA_INACTIVO).Visible = true;

                                _oForm.Items.Item(LBL_DESDE).Left = 650;
                                _oForm.Items.Item(TXT_DESDE).Left = 700;
                                _oForm.Items.Item(LBL_HASTA).Left = 760;
                                _oForm.Items.Item(TXT_HASTA).Left = 810;


                            }
                        }
                        if (pVal.ItemUID == CHK_VALIDO)
                        {
                            if (_oForm.Mode != BoFormMode.fm_FIND_MODE && _oForm.Mode != BoFormMode.fm_VIEW_MODE)
                            {
                                _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                _oForm.Items.Item(TXT_NAME_PROVEEDOR).Click();
                                _oForm.Items.Item(LBL_MOTIVO_SOCIO_INACTIVO).Visible = false;
                                _oForm.Items.Item(TXT_MOTIVO_SOCIO_INACTIVO).Visible = false;
                                _oForm.Items.Item(LBL_COMENTARIOS_MOTIVO).Visible = false;
                                _oForm.Items.Item(TXT_COMENTARIOS_MOTIVO).Visible = false;
                                _oForm.Items.Item(LBL_FECHA_INACTIVO).Visible = false;
                                _oForm.Items.Item(TXT_FECHA_INACTIVO).Visible = false;
                            }
                        }
                        if (pVal.ItemUID == CHK_AVANZADO)
                        {
                            if (_oForm.Mode != BoFormMode.fm_FIND_MODE && _oForm.Mode != BoFormMode.fm_VIEW_MODE)
                            {
                                _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                _oForm.Items.Item(LBL_MOTIVO_SOCIO_INACTIVO).Visible = false;
                                _oForm.Items.Item(TXT_MOTIVO_SOCIO_INACTIVO).Visible = false;
                                _oForm.Items.Item(LBL_COMENTARIOS_MOTIVO).Visible = false;
                                _oForm.Items.Item(TXT_COMENTARIOS_MOTIVO).Visible = false;
                                _oForm.Items.Item(LBL_FECHA_INACTIVO).Visible = false;
                                _oForm.Items.Item(TXT_FECHA_INACTIVO).Visible = false;
                            }
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_GOT_FOCUS && pVal.ItemUID == TXT_CARDCODE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        if (_oForm.Visible == true)
                        {
                            if (_oForm.Mode != BoFormMode.fm_FIND_MODE && _oForm.Mode != BoFormMode.fm_VIEW_MODE)
                            {
                                OcultarControlesVentana(_oForm);
                            }
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_LOST_FOCUS && pVal.ItemUID == TXT_CARDCODE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        if (_oForm.Visible == true)
                        {
                            if (_oForm.Mode != BoFormMode.fm_FIND_MODE && _oForm.Mode != BoFormMode.fm_VIEW_MODE)
                            {
                                OcultarControlesVentana(_oForm);
                            }
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
                    {
                        if (_oContadorFormas == 1)
                        {
                            _Application.ItemEvent -= new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                            _Application.MenuEvent -= new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                            _Application.FormDataEvent -= new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormEvent);
                            _Application.StatusBarEvent -= new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(SBO_Application_StatusBarEvent);
                            Dispose();
                            application = null;
                            company = null;
                            _oDatosMaestrosSociosNegocio = false;
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
                        _oForm.Freeze(true);
                        OcultarControlesVentana(_oForm);
                        CrearCamposDeUsuario(_oForm);
                        _oForm.Freeze(false);

                        if (_oForm.Mode == BoFormMode.fm_FIND_MODE && lDatosContratos != null)
                        {
                            if (lDatosContratos.Count != 0)
                            {
                                if (_Application.Menus.Item(BTN_CREAR_SAP).Enabled == true)
                                {
                                    _oForm.Mode = BoFormMode.fm_ADD_MODE;
                                    SAPbouiCOM.Form _oFormaCDU = null;
                                    try
                                    {
                                        _oFormaCDU = _Application.Forms.GetFormByTypeAndCount(-pVal.FormType, pVal.FormTypeCount);
                                    }
                                    catch (Exception)
                                    {
                                        _Application.MessageBox("Favor de mostrar los campos definidos por el usuario");
                                        lDatosContratos.Clear();
                                        _oForm.Close();
                                        return;
                                    }

                                    AsignarDatosUDF(_oFormaCDU);
                                    string grupoUsuario = Extensor.ObtenerGrupoSocioNegocio(_oForm, _Company);
                                    if (!string.IsNullOrEmpty(grupoUsuario))
                                    {
                                        _oCmbGrupo = _oForm.Items.Item(CMB_GRUPO).Specific;
                                        _oCmbGrupo.Select(grupoUsuario, BoSearchKey.psk_ByValue);
                                        //_oForm.Items.Item(TXT_NAME_PROVEEDOR).Click();                                      
                                    }

                                    string claseSN = Extensor.ObtenerClaseSocioNegocio(_oForm, _Company);
                                    if (!string.IsNullOrEmpty(claseSN))
                                    {
                                        _oCmbClaseSN = _oForm.Items.Item(CMB_CLASE_SN).Specific;
                                        _oCmbClaseSN.Select(claseSN, BoSearchKey.psk_ByDescription);
                                    }

                                    OcultarControlesVentana(_oForm);
                                }
                            }
                            else
                            {
                                string grupoUsuario = Extensor.ObtenerGrupoSocioNegocio(_oForm, _Company);
                                if (!string.IsNullOrEmpty(grupoUsuario))
                                {
                                    _oCmbGrupo = _oForm.Items.Item(CMB_GRUPO).Specific;
                                    _oCmbGrupo.Select(grupoUsuario, BoSearchKey.psk_ByValue);
                                }
                            }
                        }
                        else if (_oForm.Mode == BoFormMode.fm_FIND_MODE)
                        {
                            string grupoUsuario = Extensor.ObtenerGrupoSocioNegocio(_oForm, _Company);
                            if (!string.IsNullOrEmpty(grupoUsuario))
                            {
                                _oCmbGrupo = _oForm.Items.Item(CMB_GRUPO).Specific;
                                _oCmbGrupo.Select(grupoUsuario, BoSearchKey.psk_ByValue);
                            }
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_ACTIVATE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        if (_oForm.Mode != BoFormMode.fm_FIND_MODE && _oForm.Mode != BoFormMode.fm_VIEW_MODE)
                        {

                        }
                    }
                }

                if (pVal.BeforeAction == false && pVal.FormType == FRM_DATOS_MAESTROS_SOCIO)
                {
                    if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == CMB_CLASE_SN)
                    {
                        SAPbouiCOM.Form _oFormCDU = null;
                        string solicitudInterna = null;
                        string prefijoSerie = null;
                        string prefijoComisiones = null;
                        string plan = null;

                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        _oCmbGrupo = _oForm.Items.Item(CMB_GRUPO).Specific;
                        _oCmbClaseSN = _oForm.Items.Item(CMB_CLASE_SN).Specific;
                        _oCodigoSerieSAP = _oForm.Items.Item(CMB_SERIE_FOLIO_SAP).Specific;
                        _oTxtCardCode = _oForm.Items.Item(TXT_CARDCODE).Specific;


                        if (Extensor.Configuracion.SERIECONTRATOS.SerieContratosAutomatica == "Y")
                        {
                            if (_oCmbClaseSN.Selected.Value == "C")
                            {
                                if (_oCmbGrupo.Selected != null)
                                {
                                    if (_oCmbGrupo.Selected.Description.Contains("PABS"))
                                    {
                                        try
                                        {
                                            _oForm.Items.Item(BTN_VALIDAR_SOLICITUD).Visible = false;
                                            _oFormCDU = _Application.Forms.GetFormByTypeAndCount(-pVal.FormType, pVal.FormTypeCount);
                                            _oTxtTrasp_CodPlan = _oFormCDU.Items.Item(CDU_CODIGO_PLAN_PREVISION).Specific;
                                            _oTxtSolicitudInterna = _oFormCDU.Items.Item(CDU_SOLICITUD_INTERNA).Specific;
                                            solicitudInterna = _oTxtSolicitudInterna.Value.ToString();
                                            plan = _oTxtTrasp_CodPlan.Value.ToString();
                                            if (!string.IsNullOrEmpty(solicitudInterna))
                                            {
                                                prefijoSerie = solicitudInterna.Substring(0, 6);
                                                prefijoComisiones = Extensor.ObtenerPrefijoSerie(prefijoSerie, plan, _Company);
                                                _oCodigoSerieSAP.Select(prefijoComisiones, BoSearchKey.psk_ByValue);
                                                _oForm.Items.Item(TXT_RFC).Click();
                                            }
                                        }
                                        catch (Exception)
                                        {
                                        }
                                    }
                                }
                            }
                            else if (_oCmbClaseSN.Selected.Value == "L")
                            {
                                if (_oCmbGrupo.Selected != null)
                                {
                                    if (_oCmbGrupo.Selected.Description.Contains("PABS"))
                                    {
                                        if (_oForm.Mode == BoFormMode.fm_ADD_MODE)
                                        {
                                            _oForm.Items.Item(BTN_VALIDAR_SOLICITUD).Visible = true;
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (_oCmbGrupo.Selected != null)
                                if (_oCmbGrupo.Selected.Description.Contains("PABS"))
                                    _oTxtCardCode.Value = "";
                        }
                    }
                }


                if (pVal.BeforeAction == true && pVal.FormType == FRM_DATOS_MAESTROS_SOCIO)
                {
                    if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == CMB_CLASE_SN)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        if (_oForm.Mode != BoFormMode.fm_FIND_MODE && _oForm.Mode != BoFormMode.fm_VIEW_MODE)
                        {
                            _oNombreSocioNegocio = _oForm.Items.Item(TXT_NAME_PROVEEDOR).Specific;
                            _oCmbGrupo = _oForm.Items.Item(CMB_GRUPO).Specific;
                            _oCmbClaseSN = _oForm.Items.Item(CMB_CLASE_SN).Specific;

                            string msgError = null;

                            if (_oCmbClaseSN.Selected.Value != "C")
                            {

                                if (_oCmbGrupo.Selected.Description.Contains("PABS"))
                                {
                                    msgError = Extensor.ValidarSiExisteSocioDeNegocio(_oNombreSocioNegocio.Value.ToString(), _Company);
                                }

                                if (!string.IsNullOrEmpty(msgError))
                                {
                                    _Application.MessageBox(msgError + Environment.NewLine + "Favor de validarlo");
                                }

                            }
                        }
                    }


                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if (pVal.ItemUID == BTN_BUSCAR_VENTANA)
                        {
                            _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                            SAPbouiCOM.Button boton = _oForm.Items.Item(BTN_BUSCAR_VENTANA).Specific;
                            if (boton.Caption == "Buscar")
                            {
                                OcultarControlesVentana(_oForm);
                                _oForm.Select();
                            }
                        }
                    }
                    if (pVal.EventType == BoEventTypes.et_FORM_ACTIVATE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        if (_oForm.Mode != BoFormMode.fm_FIND_MODE && _oForm.Mode != BoFormMode.fm_VIEW_MODE)
                        {

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en método 'eventos' *clsDatosMaestrosSocioNegocio* : " + ex.Message);
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
                if (BusinessObjectInfo.BeforeAction == true && BusinessObjectInfo.FormTypeEx == FRM_DATOS_MAESTROS_SOCIO.ToString() &&
                BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_UPDATE && BusinessObjectInfo.Type == OBJETO_DMSN &&
                BusinessObjectInfo.ActionSuccess == false)
                {
                    string msgError = null;
                    SAPbouiCOM.Form _UDFForm = null;
                    SAPbouiCOM.Form _Form = null;

                    try
                    {
                        _UDFForm = _Application.Forms.GetForm("-" + BusinessObjectInfo.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);
                        _Form = _Application.Forms.GetForm(BusinessObjectInfo.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);
                        _oCmbClaseSN = _Form.Items.Item(CMB_CLASE_SN).Specific;
                        _oRFCSocioNegocio = _Form.Items.Item(TXT_RFC).Specific;
                        _oCmbGrupo = _Form.Items.Item(CMB_GRUPO).Specific;

                        if (_oCmbClaseSN.Selected.Value == "C")
                        {
                            _oTxtCodCobrador = _UDFForm.Items.Item(CDU_CODCOB).Specific;
                            _oTxtNomCobrador = _UDFForm.Items.Item(CDU_NOMCOB).Specific;
                            _oTxtTraspasoRel = _UDFForm.Items.Item(CDU_TRASPASOREL).Specific;
                            _oTxtSolicitudInterna = _UDFForm.Items.Item(CDU_SOLICITUD_INTERNA).Specific;
                            _oTxtTrasp_CodPlan = _UDFForm.Items.Item(CDU_CODIGO_PLAN_PREVISION).Specific;
                            _oTxtAct_Descripcion = _UDFForm.Items.Item(CDU_NOMBRE_PLAN_PREVISION).Specific;
                            _oTxtInvInicial = _UDFForm.Items.Item(CDU_INVINICIAL).Specific;
                            _oTxtPapaleria = _UDFForm.Items.Item(CDU_PAPELERIA).Specific;
                            _oTxtExcedente_InvIni = _UDFForm.Items.Item(CDU_EXCEDENTE_INVINI).Specific;
                            _oTxtBono = _UDFForm.Items.Item(CDU_BONO).Specific;
                            _oTxtBeneficiarioPagoRecibido = _UDFForm.Items.Item(CDU_BENEFICIARIO_PAGO_RECIBIDO).Specific;

                            if (!string.IsNullOrEmpty(_oTxtTraspasoRel.Value.ToString()) &&
                                !string.IsNullOrEmpty(_oTxtSolicitudInterna.Value.ToString()) && !string.IsNullOrEmpty(_oTxtTrasp_CodPlan.Value.ToString()) &&
                                !string.IsNullOrEmpty(_oTxtAct_Descripcion.Value.ToString()) && Convert.ToDouble(_oTxtInvInicial.Value) != 0 &&
                                Convert.ToDouble(_oTxtPapaleria.Value) != 0)
                            {
                                if (string.IsNullOrEmpty(_oTxtBeneficiarioPagoRecibido.Value.ToString()))
                                {
                                    _Application.MessageBox("Capture el nombre del beneficiario");
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _Application.StatusBar.SetText("Error al insertar calculo comisiones DATA_UPDATE: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                }

                if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.FormTypeEx == FRM_DATOS_MAESTROS_SOCIO.ToString())
                {
                    if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_UPDATE && BusinessObjectInfo.ActionSuccess == true)
                    {
                        if (_Application != null)
                        {
                            string CardcodeRespaldo = null;
                            try
                            {
                                SAPbouiCOM.Form _UDFForm = null;
                                SAPbouiCOM.Form _oNuevaForm = _Application.Forms.GetForm(BusinessObjectInfo.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);
                                if (BusinessObjectInfo.FormTypeEx == FRM_DATOS_MAESTROS_SOCIO.ToString())
                                {
                                    _UDFForm = _Application.Forms.GetForm("-" + BusinessObjectInfo.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);
                                    _oNuevaForm = _Application.Forms.GetForm(BusinessObjectInfo.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);
                                    _oCmbClaseSN = _oNuevaForm.Items.Item(CMB_CLASE_SN).Specific;

                                    if (_oCmbClaseSN.Selected.Value == "C")
                                    {
                                        System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
                                        oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey);
                                        string DocEntryCardCode = oXmlDoc.SelectSingleNode("/BusinessPartnerParams/CardCode").InnerText;
                                        CardcodeRespaldo = DocEntryCardCode;
                                        _oTxtCodCobrador = _UDFForm.Items.Item(CDU_CODCOB).Specific;
                                        _oTxtNomCobrador = _UDFForm.Items.Item(CDU_NOMCOB).Specific;
                                        _oTxtTraspasoRel = _UDFForm.Items.Item(CDU_TRASPASOREL).Specific;
                                        _oTxtSolicitudInterna = _UDFForm.Items.Item(CDU_SOLICITUD_INTERNA).Specific;
                                        _oTxtTrasp_CodPlan = _UDFForm.Items.Item(CDU_CODIGO_PLAN_PREVISION).Specific;
                                        _oTxtAct_Descripcion = _UDFForm.Items.Item(CDU_NOMBRE_PLAN_PREVISION).Specific;
                                        _oTxtComision = _UDFForm.Items.Item(CDU_COMISION).Specific;
                                        _oTxtInvInicial = _UDFForm.Items.Item(CDU_INVINICIAL).Specific;
                                        _oTxtPapaleria = _UDFForm.Items.Item(CDU_PAPELERIA).Specific;
                                        _oTxtExcedente_InvIni = _UDFForm.Items.Item(CDU_EXCEDENTE_INVINI).Specific;
                                        _oTxtBono = _UDFForm.Items.Item(CDU_BONO).Specific;
                                        _oTxtBeneficiarioPagoRecibido = _UDFForm.Items.Item(CDU_BENEFICIARIO_PAGO_RECIBIDO).Specific;
                                        _oTxtCodigoAsisten = _UDFForm.Items.Item(CDU_CODIGO_ASISTENTE).Specific;
                                        _oTxtNombreAsistente = _UDFForm.Items.Item(CDU_NOMBRE_ASISTENTE).Specific;
                                        _oTxtFechaCreacion = _oNuevaForm.Items.Item(TXT_FECHA_CREACION).Specific;
                                        _oCmbEsquemaPago = _UDFForm.Items.Item(CDU_ESQUEMA_PAGO).Specific;


                                        if (!string.IsNullOrEmpty(_oTxtTraspasoRel.Value.ToString()) &&
                                            !string.IsNullOrEmpty(_oTxtSolicitudInterna.Value.ToString()) && !string.IsNullOrEmpty(_oTxtTrasp_CodPlan.Value.ToString()) &&
                                            !string.IsNullOrEmpty(_oTxtAct_Descripcion.Value.ToString()) && Convert.ToDouble(_oTxtInvInicial.Value) != 0 &&
                                            Convert.ToDouble(_oTxtPapaleria.Value) != 0)
                                        {
                                            if (Extensor.ValidarSiExistePlan(DocEntryCardCode, _Company))
                                            {
                                                InsertarCalculoComisiones(_oTxtTraspasoRel.Value, DocEntryCardCode, _oTxtSolicitudInterna.Value, _oTxtCodCobrador.Value.ToString(), _oTxtNomCobrador.Value.ToString(),
                                                                            _oTxtTrasp_CodPlan.Value, _oTxtInvInicial.Value, _oTxtBono.Value, _oTxtComision.Value, _oTxtCodigoAsisten.Value.ToString(),
                                                                            _oTxtNombreAsistente.Value.ToString(), _oCmbEsquemaPago.Value.ToString());

                                                string msgError = CrearFacturaPago(_oNuevaForm, DocEntryCardCode, _oTxtFechaCreacion.Value.ToString(), _oTxtCodigoAsisten.Value.ToString());
                                                if (string.IsNullOrEmpty(msgError))
                                                {
                                                    _Application.MessageBox("Se crea contrato No. : " + DocEntryCardCode + " correctamente");
                                                }
                                                else
                                                {                                                                                                       
                                                        _Application.MessageBox("Error: " + msgError + "  - Favor de actualizar nuevamente");
                                                }
                                            }
                                            else
                                            {
                                                if (Extensor.ValidarSiEstaCobrador(DocEntryCardCode, _Company))
                                                {
                                                    if (!string.IsNullOrEmpty(_oTxtCodCobrador.Value.ToString()))
                                                    {
                                                        ActualizarCobrador(DocEntryCardCode, _oTxtSolicitudInterna.Value, _oTxtCodCobrador.Value.ToString(), _oTxtNomCobrador.Value.ToString());
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    OcultarControlesVentana(_oNuevaForm);
                                }
                            }
                            catch (Exception ex)
                            {
                                _Application.StatusBar.SetText("Error al insertar calculo comisiones DATA_ADD_UPDATE: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);                              

                            }
                        }
                    }

                    if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD)
                    {
                        if (_Application != null)
                        {
                            if (BusinessObjectInfo.FormTypeEx == FRM_DATOS_MAESTROS_SOCIO.ToString())
                            {
                                if (BusinessObjectInfo.FormTypeEx == _Application.Forms.ActiveForm.TypeEx)
                                {
                                    SAPbouiCOM.Form _oNuevaForm = _Application.Forms.GetForm(BusinessObjectInfo.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);
                                    OcultarControlesVentana(_oNuevaForm);
                                    CrearCamposDeUsuario(_oNuevaForm);

                                    _oNuevaForm.Refresh();
                                    _oNuevaForm.Update();
                                }
                            }
                        }
                    }
                }

                if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.FormTypeEx == FRM_DATOS_MAESTROS_SOCIO.ToString() &&
                    BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.Type == OBJETO_DMSN &&
                    BusinessObjectInfo.ActionSuccess == true)
                {
                    SAPbouiCOM.Form _UDFForm = null;
                    SAPbouiCOM.Form _Form = null;

                    try
                    {
                        _UDFForm = _Application.Forms.GetForm("-" + BusinessObjectInfo.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);
                        _Form = _Application.Forms.GetForm(BusinessObjectInfo.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);
                        _oCmbClaseSN = _Form.Items.Item(CMB_CLASE_SN).Specific;
                        _oCmbGrupo = _Form.Items.Item(CMB_GRUPO).Specific;

                        if (_oCmbClaseSN.Selected.Value == "C")
                        {
                            System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
                            oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey);
                            string DocEntryCardCode = oXmlDoc.SelectSingleNode("/BusinessPartnerParams/CardCode").InnerText;

                            _oTxtCodCobrador = _UDFForm.Items.Item(CDU_CODCOB).Specific;
                            _oTxtNomCobrador = _UDFForm.Items.Item(CDU_NOMCOB).Specific;
                            _oTxtTraspasoRel = _UDFForm.Items.Item(CDU_TRASPASOREL).Specific;
                            _oTxtSolicitudInterna = _UDFForm.Items.Item(CDU_SOLICITUD_INTERNA).Specific;
                            _oTxtTrasp_CodPlan = _UDFForm.Items.Item(CDU_CODIGO_PLAN_PREVISION).Specific;
                            _oTxtAct_Descripcion = _UDFForm.Items.Item(CDU_NOMBRE_PLAN_PREVISION).Specific;
                            _oTxtInvInicial = _UDFForm.Items.Item(CDU_INVINICIAL).Specific;
                            _oTxtPapaleria = _UDFForm.Items.Item(CDU_PAPELERIA).Specific;
                            _oTxtExcedente_InvIni = _UDFForm.Items.Item(CDU_EXCEDENTE_INVINI).Specific;
                            _oTxtBono = _UDFForm.Items.Item(CDU_BONO).Specific;
                            _oTxtBeneficiarioPagoRecibido = _UDFForm.Items.Item(CDU_BENEFICIARIO_PAGO_RECIBIDO).Specific;
                            _oTxtComision = _UDFForm.Items.Item(CDU_COMISION).Specific;
                            _oTxtCodigoAsisten = _UDFForm.Items.Item(CDU_CODIGO_ASISTENTE).Specific;
                            _oTxtNombreAsistente = _UDFForm.Items.Item(CDU_NOMBRE_ASISTENTE).Specific;
                            _oTxtFechaCreacion = _Form.Items.Item(TXT_FECHA_CREACION).Specific;
                            _oCmbEsquemaPago = _UDFForm.Items.Item(CDU_ESQUEMA_PAGO).Specific;

                            if (!string.IsNullOrEmpty(_oTxtTraspasoRel.Value.ToString()) &&
                                !string.IsNullOrEmpty(_oTxtSolicitudInterna.Value.ToString()) && !string.IsNullOrEmpty(_oTxtTrasp_CodPlan.Value.ToString()) &&
                                !string.IsNullOrEmpty(_oTxtAct_Descripcion.Value.ToString()) && Convert.ToDouble(_oTxtInvInicial.Value) != 0 &&
                                Convert.ToDouble(_oTxtPapaleria.Value) != 0)
                            {
                                if (Extensor.ValidarSiExistePlan(DocEntryCardCode, _Company))
                                {
                                    InsertarCalculoComisiones(_oTxtTraspasoRel.Value, DocEntryCardCode, _oTxtSolicitudInterna.Value, _oTxtCodCobrador.Value.ToString(), _oTxtNomCobrador.Value.ToString(),
                                                                _oTxtTrasp_CodPlan.Value, _oTxtInvInicial.Value, _oTxtBono.Value, _oTxtComision.Value, _oTxtCodigoAsisten.Value.ToString(),
                                                                            _oTxtNombreAsistente.Value.ToString(), _oCmbEsquemaPago.Value.ToString());

                                    string msgError = CrearFacturaPago(_Form, DocEntryCardCode, _oTxtFechaCreacion.Value.ToString(), _oTxtCodigoAsisten.Value.ToString());

                                    if (string.IsNullOrEmpty(msgError))
                                    {
                                        _Application.MessageBox("Se crea contrato No. : " + DocEntryCardCode + " correctamente");
                                    }
                                    else
                                    {
                                        try
                                        {
                                            _oRec = null;
                                            _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            _oRec.DoQuery(@"DELETE  dbo.[@CALCULO_COMISIONES]
                                                                    WHERE   U_Contrato = '" + DocEntryCardCode + "'");
                                            _Application.MessageBox("Error: " + msgError + "  - Favor de actualizar nuevamente");
                                        }
                                        catch (Exception)
                                        {
                                        }
                                        finally
                                        {
                                            if (_oRec != null)
                                            {
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(_oRec);
                                            }
                                        }
                                    }

                                    if (Extensor.ValidarSiEstaCobrador(DocEntryCardCode, _Company))
                                    {
                                        if (!string.IsNullOrEmpty(_oTxtCodCobrador.Value.ToString()))
                                        {
                                            ActualizarCobrador(DocEntryCardCode, _oTxtSolicitudInterna.Value, _oTxtCodCobrador.Value.ToString(), _oTxtNomCobrador.Value.ToString());
                                        }
                                    }
                                }
                            }
                        }
                        else if (_oCmbClaseSN.Selected.Value == "L")
                        {
                            System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
                            oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey);
                            string DocEntryCardCode = oXmlDoc.SelectSingleNode("/BusinessPartnerParams/CardCode").InnerText;

                            _oRec = null;
                            _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            _oRec.DoQuery(@"SELECT U_CodigoActivacion FROM dbo.OCRD WHERE CardCode = '" + DocEntryCardCode + "'");


                            _Application.MessageBox("El código de activación es: " + _oRec.Fields.Item("U_CodigoActivacion").Value);
                        }
                    }
                    catch (Exception ex)
                    {
                        _Application.StatusBar.SetText("Error al insertar calculo comisiones DATA_ADD: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                if (BusinessObjectInfo.BeforeAction == true && BusinessObjectInfo.FormTypeEx == FRM_DATOS_MAESTROS_SOCIO.ToString() &&
                    BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.Type == OBJETO_DMSN &&
                    BusinessObjectInfo.ActionSuccess == false)
                {
                    string msgError = null;
                    SAPbouiCOM.Form _UDFForm = null;
                    SAPbouiCOM.Form _Form = null;
                    string colonia = null;
                    string municipio = null;

                    try
                    {
                        _UDFForm = _Application.Forms.GetForm("-" + BusinessObjectInfo.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);
                        _Form = _Application.Forms.GetForm(BusinessObjectInfo.FormTypeEx, _Application.Forms.ActiveForm.TypeCount);
                        _oCmbClaseSN = _Form.Items.Item(CMB_CLASE_SN).Specific;

                        Item item = _Form.Items.Item(TXT_COLONIA);
                        Matrix matrix = item.Specific;

                        _oNombreSocioNegocio = _Form.Items.Item(TXT_NAME_PROVEEDOR).Specific;
                        colonia = matrix.Columns.Item("3").Cells.Item(1).Specific.Value;
                        municipio = matrix.Columns.Item("4").Cells.Item(1).Specific.value;
                        _oCmbGrupo = _Form.Items.Item(CMB_GRUPO).Specific;

                        if (_oCmbGrupo.Selected.Description.Contains("PABS"))
                        {
                            msgError = Extensor.ValidarSiExisteSocioDeNegocio(_oNombreSocioNegocio.Value.ToString(), _Company);
                        }

                        if (string.IsNullOrEmpty(msgError))
                        {
                            if (_oCmbClaseSN.Selected.Value == "C")
                            {
                                _oTxtCodCobrador = _UDFForm.Items.Item(CDU_CODCOB).Specific;
                                _oTxtNomCobrador = _UDFForm.Items.Item(CDU_NOMCOB).Specific;
                                _oTxtTraspasoRel = _UDFForm.Items.Item(CDU_TRASPASOREL).Specific;
                                _oTxtSolicitudInterna = _UDFForm.Items.Item(CDU_SOLICITUD_INTERNA).Specific;
                                _oTxtTrasp_CodPlan = _UDFForm.Items.Item(CDU_CODIGO_PLAN_PREVISION).Specific;
                                _oTxtAct_Descripcion = _UDFForm.Items.Item(CDU_NOMBRE_PLAN_PREVISION).Specific;
                                _oTxtInvInicial = _UDFForm.Items.Item(CDU_INVINICIAL).Specific;
                                _oTxtPapaleria = _UDFForm.Items.Item(CDU_PAPELERIA).Specific;
                                _oTxtExcedente_InvIni = _UDFForm.Items.Item(CDU_EXCEDENTE_INVINI).Specific;
                                _oTxtBono = _UDFForm.Items.Item(CDU_BONO).Specific;
                                _oTxtBeneficiarioPagoRecibido = _UDFForm.Items.Item(CDU_BENEFICIARIO_PAGO_RECIBIDO).Specific;

                                if (!string.IsNullOrEmpty(_oTxtTraspasoRel.Value.ToString()) &&
                                    !string.IsNullOrEmpty(_oTxtSolicitudInterna.Value.ToString()) && !string.IsNullOrEmpty(_oTxtTrasp_CodPlan.Value.ToString()) &&
                                    !string.IsNullOrEmpty(_oTxtAct_Descripcion.Value.ToString()) && Convert.ToDouble(_oTxtInvInicial.Value) != 0 &&
                                    Convert.ToDouble(_oTxtPapaleria.Value) != 0)
                                {
                                    if (string.IsNullOrEmpty(_oTxtBeneficiarioPagoRecibido.Value.ToString()))
                                    {
                                        _Application.MessageBox("Capture el nombre del beneficiario");
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                            }
                        }
                        else
                        {

                            if (_Application.MessageBox(msgError + Environment.NewLine + "¿Desea crearlo?", 2, "Si", "No") == 1)
                            {
                                if (_oCmbClaseSN.Selected.Value == "C")
                                {
                                    _oTxtCodCobrador = _UDFForm.Items.Item(CDU_CODCOB).Specific;
                                    _oTxtNomCobrador = _UDFForm.Items.Item(CDU_NOMCOB).Specific;
                                    _oTxtTraspasoRel = _UDFForm.Items.Item(CDU_TRASPASOREL).Specific;
                                    _oTxtSolicitudInterna = _UDFForm.Items.Item(CDU_SOLICITUD_INTERNA).Specific;
                                    _oTxtTrasp_CodPlan = _UDFForm.Items.Item(CDU_CODIGO_PLAN_PREVISION).Specific;
                                    _oTxtAct_Descripcion = _UDFForm.Items.Item(CDU_NOMBRE_PLAN_PREVISION).Specific;
                                    _oTxtInvInicial = _UDFForm.Items.Item(CDU_INVINICIAL).Specific;
                                    _oTxtPapaleria = _UDFForm.Items.Item(CDU_PAPELERIA).Specific;
                                    _oTxtExcedente_InvIni = _UDFForm.Items.Item(CDU_EXCEDENTE_INVINI).Specific;
                                    _oTxtBono = _UDFForm.Items.Item(CDU_BONO).Specific;
                                    _oTxtBeneficiarioPagoRecibido = _UDFForm.Items.Item(CDU_BENEFICIARIO_PAGO_RECIBIDO).Specific;

                                    if (!string.IsNullOrEmpty(_oTxtTraspasoRel.Value.ToString()) &&
                                        !string.IsNullOrEmpty(_oTxtSolicitudInterna.Value.ToString()) && !string.IsNullOrEmpty(_oTxtTrasp_CodPlan.Value.ToString()) &&
                                        !string.IsNullOrEmpty(_oTxtAct_Descripcion.Value.ToString()) && Convert.ToDouble(_oTxtInvInicial.Value) != 0 &&
                                        Convert.ToDouble(_oTxtPapaleria.Value) != 0)
                                    {
                                        if (string.IsNullOrEmpty(_oTxtBeneficiarioPagoRecibido.Value.ToString()))
                                        {
                                            _Application.MessageBox("Capture el nombre del beneficiario");
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                BubbleEvent = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _Application.StatusBar.SetText("Error al insertar calculo comisiones DATA_ADD: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                }
            }
            catch (Exception ex)
            {
                







_Application.MessageBox("Error en FormEvent *clsDatosMaestrosSocioNegocio* : " + ex.Message);
            }
        }


        /// <summary>
        /// Obtiene los eventos del menú click derecho
        /// </summary>
        /// <param name="pVal">Propiedades de la forma</param>
        /// <param name="BubbleEvent">Evento</param>
        private void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Form formanew = _Application.Forms.ActiveForm;
                if (pVal.MenuUID == Convert.ToString((int)ClickDerecho.Cortar) ||
               pVal.MenuUID == Convert.ToString((int)ClickDerecho.Borrar) ||
               pVal.MenuUID == Convert.ToString((int)ClickDerecho.Pegar)
               && pVal.BeforeAction == true)
                {

                    if (formanew.Mode != BoFormMode.fm_FIND_MODE)
                    {
                        if (userFields.Contains(userFielsType) && formTypes.Contains(FormType))
                        {
                            BubbleEvent = false;
                        }
                    }
                }


                if (pVal.BeforeAction == false && formanew.TypeEx == FRM_DATOS_MAESTROS_SOCIO.ToString())
                {
                    if (pVal.MenuUID == BTN_CREAR_SAP)
                    {
                        formanew.Freeze(true);
                        OcultarControlesVentana(formanew);

                        string grupoUsuario = Extensor.ObtenerGrupoSocioNegocio(formanew, _Company);
                        if (!string.IsNullOrEmpty(grupoUsuario))
                        {
                            _oCmbGrupo = formanew.Items.Item(CMB_GRUPO).Specific;
                            if (_oCmbGrupo.Item.Enabled == true)
                            {
                                _oCmbGrupo.Select(grupoUsuario, BoSearchKey.psk_ByValue);
                            }
                        }

                        string claseSN = Extensor.ObtenerClaseSocioNegocio(formanew, _Company);
                        if (!string.IsNullOrEmpty(claseSN))
                        {
                            _oCmbClaseSN = formanew.Items.Item(CMB_CLASE_SN).Specific;
                            _oCmbClaseSN.Select(claseSN, BoSearchKey.psk_ByDescription);
                        }

                        OcultarControlesVentana(formanew);
                    }
                    if (pVal.MenuUID == BTN_BUSCAR_SAP)
                    {
                        string grupoUsuario = Extensor.ObtenerGrupoSocioNegocio(formanew, _Company);
                        if (!string.IsNullOrEmpty(grupoUsuario))
                        {
                            _oCmbGrupo = formanew.Items.Item(CMB_GRUPO).Specific;
                            _oCmbGrupo.Select(grupoUsuario, BoSearchKey.psk_ByValue);
                        }
                        OcultarControlesVentana(formanew);
                    }
                    formanew.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                _Application.MessageBox("Error en MenuEvent *clsDatosMaestrosSocioNegocio* : " + ex.Message);
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
                if (Text.Contains("Se ha producido un error interno (-7780)  [Mensaje 131-183]") && messageType == BoStatusBarMessageType.smt_Error)
                {
                    _Application.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
                }
            }
            catch (Exception ex)
            {
                _Application.MessageBox("Error en StatusBarEvent *clsDatosMaestrosSocioNegocio* : " + ex.Message);
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
            _Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            _Application.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(SBO_Application_StatusBarEvent);
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
                    string s = _oForma.Items.Item(LBL_MOTIVO_SOCIO_INACTIVO).UniqueID;
                }
                catch (Exception)
                {
                    _oForma.Freeze(true);

                    #region SOCIO NEGOCIO INACTIVO

                    //LABEL 'MOTIVO'
                    SAPbouiCOM.Item chkInactivo = null;
                    chkInactivo = _oForma.Items.Item(CHK_INACTIVO);
                    newItem = _oForma.Items.Add(LBL_MOTIVO_SOCIO_INACTIVO, BoFormItemTypes.it_STATIC);
                    newItem.Left = chkInactivo.Left + 90;
                    newItem.Top = chkInactivo.Top;
                    newItem.ToPane = 100;
                    newItem.FromPane = 100;
                    newItem.Width = 50;
                    _oLblMotivo_Inactivo = newItem.Specific;
                    _oLblMotivo_Inactivo.Caption = "Motivo";
                    //newItem.TextStyle = 1;
                    newItem.Visible = false;

                    //EDITTEXT 'MOTIVO'
                    SAPbouiCOM.Item lblMotiv = null;
                    lblMotiv = _oForma.Items.Item(LBL_MOTIVO_SOCIO_INACTIVO);
                    newItem = _oForma.Items.Add(TXT_MOTIVO_SOCIO_INACTIVO, BoFormItemTypes.it_EDIT);
                    newItem.Left = lblMotiv.Left + 50;
                    newItem.Top = lblMotiv.Top;
                    newItem.ToPane = 100;
                    newItem.FromPane = 100;
                    newItem.Width = 150;
                    _oForma.DataSources.UserDataSources.Add(TXT_MOTIVO_SOCIO_INACTIVO, BoDataType.dt_LONG_TEXT, 150);
                    _oTxtMotivo_Inactivo = newItem.Specific;
                    _oTxtMotivo_Inactivo.DataBind.SetBound(true, "OCRD", "U_MotivoInactivo");
                    lblMotiv.LinkTo = newItem.UniqueID;
                    newItem.Visible = false;

                    //LABEL 'COMENTARIOS'
                    SAPbouiCOM.Item txtMotInac = null;
                    txtMotInac = _oForma.Items.Item(TXT_MOTIVO_SOCIO_INACTIVO);
                    newItem = _oForma.Items.Add(LBL_COMENTARIOS_MOTIVO, BoFormItemTypes.it_STATIC);
                    newItem.Left = txtMotInac.Left + 160;
                    newItem.Top = txtMotInac.Top;
                    newItem.ToPane = 100;
                    newItem.FromPane = 100;
                    newItem.Width = 80;
                    _oLblComentarios_Inactivo = newItem.Specific;
                    _oLblComentarios_Inactivo.Caption = "Comentarios";
                    //newItem.TextStyle = 1;
                    newItem.Visible = false;

                    //EDITTEXT 'COMENTARIOS'
                    SAPbouiCOM.Item lblComenInac = null;
                    lblComenInac = _oForma.Items.Item(LBL_COMENTARIOS_MOTIVO);
                    newItem = _oForma.Items.Add(TXT_COMENTARIOS_MOTIVO, BoFormItemTypes.it_EXTEDIT);
                    newItem.Left = lblComenInac.Left + 80;
                    newItem.Top = lblComenInac.Top;
                    newItem.ToPane = 100;
                    newItem.FromPane = 100;
                    newItem.Width = 170;
                    newItem.Height = 40;
                    _oForma.DataSources.UserDataSources.Add(TXT_COMENTARIOS_MOTIVO, BoDataType.dt_LONG_TEXT, 254);
                    _oTxtComentarios_Inactivo = newItem.Specific;
                    _oTxtComentarios_Inactivo.DataBind.SetBound(true, "OCRD", "U_ComentariosMotivo");
                    lblComenInac.LinkTo = newItem.UniqueID;
                    newItem.Visible = false;

                    //LABEL 'FECHA MOVIMIENTO'
                    SAPbouiCOM.Item lblMotivoInac = null;
                    lblMotivoInac = _oForm.Items.Item(LBL_MOTIVO_SOCIO_INACTIVO);
                    newItem = _oForm.Items.Add(LBL_FECHA_INACTIVO, BoFormItemTypes.it_STATIC);
                    newItem.Left = lblMotivoInac.Left;
                    newItem.Top = lblMotivoInac.Top + 15;
                    newItem.ToPane = 100;
                    newItem.FromPane = 100;
                    newItem.Width = 100;
                    _oLblFecha_Inactivo = newItem.Specific;
                    _oLblFecha_Inactivo.Caption = "Fecha movimiento";
                    newItem.Visible = false;

                    //EDITTEXT 'FECHA MOVIMIENTO'
                    SAPbouiCOM.Item lblFechaMo = null;
                    lblFechaMo = _oForma.Items.Item(LBL_FECHA_INACTIVO);
                    newItem = _oForma.Items.Add(TXT_FECHA_INACTIVO, BoFormItemTypes.it_EDIT);
                    newItem.Left = lblFechaMo.Left + 100;
                    newItem.Top = lblFechaMo.Top;
                    newItem.ToPane = 100;
                    newItem.FromPane = 100;
                    newItem.Width = 100;
                    _oForma.DataSources.UserDataSources.Add(TXT_FECHA_INACTIVO, BoDataType.dt_DATE, 150);
                    _oTxtFecha_Inactivo = newItem.Specific;
                    _oTxtFecha_Inactivo.DataBind.SetBound(true, "OCRD", "U_FechaMovimiento");
                    lblFechaMo.LinkTo = newItem.UniqueID;
                    newItem.Visible = false;

                    #endregion

                    #region BOTON VALIDAR SOLICITUD

                    SAPbouiCOM.Item _oBtnCancelarSAP = null;
                    SAPbouiCOM.Button _oBoton = null;

                    _oBtnCancelarSAP = _oForma.Items.Item(BTN_CANCELAR_DATOS_MAESTROS);
                    newItem = _oForma.Items.Add(BTN_VALIDAR_SOLICITUD, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    newItem.Left = _oBtnCancelarSAP.Left + 150;
                    newItem.Top = _oBtnCancelarSAP.Top;
                    newItem.Height = _oBtnCancelarSAP.Height;
                    newItem.Width = 130;
                    newItem.ToPane = _oBtnCancelarSAP.ToPane;
                    newItem.FromPane = _oBtnCancelarSAP.FromPane;
                    newItem.DisplayDesc = true;
                    newItem.TextStyle = 3;
                    _oBoton = newItem.Specific;
                    _oBoton.Caption = "Validar solicitud";

                    #endregion

                    _oForma.Items.Item(TXT_NAME_PROVEEDOR).Specific.TabOrder = 1;
                    _oForma.Items.Item(TXT_DIA).Specific.TabOrder = 2;
                    _oForma.Items.Item(TXT_MES).Specific.TabOrder = 3;
                    _oForma.Items.Item(TXT_YEAR).Specific.TabOrder = 4;
                    _oForma.Items.Item("U_FechaPrimerAbono").Specific.TabOrder = 5;
                    _oForma.Items.Item(TXT_FECHA_CREACION).Specific.TabOrder = 6;
                    _oForma.Items.Item("U_FormaPagoBit").Specific.TabOrder = 7;
                    _oForma.Items.Item("U_MontoPago").Specific.TabOrder = 8;
                    _oForma.Items.Item(TXT_STATUS_SOLICITUD).Specific.TabOrder = 9;
                    _oForma.Items.Item(TXT_ESTADO_CIVIL).Specific.TabOrder = 10;

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
        /// Obtiene el usuario que esta firmado
        /// </summary>
        /// <param name="usuario">Usuario firmado</param>
        /// <returns>Nombre completo mas el código del usuario</returns>
        private string ObtenerUsuarioFirmado(string usuario)
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  T1.middleName + ' ' + T1.lastName + ' - ' + T0.USER_CODE AS UsuarioFirmado
                                FROM    dbo.OUSR T0
                                        INNER JOIN dbo.OHEM T1 ON T0.USERID = T1.userId
                                WHERE   USER_CODE = '" + usuario + "'");

                string resultado = _oRec.Fields.Item("UsuarioFirmado").Value.ToString();

                if (resultado.Length > 40)
                {
                    resultado = resultado.Substring(0, 40);
                }

                if (string.IsNullOrEmpty(resultado))
                {
                    resultado = "No asignado";
                }

                return resultado;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener el usuario firmado *ObtenerUsuarioFirmado* : " + ex.Message);
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
        /// Oculta los controles de la ventana activa
        /// </summary>
        /// <param name="_oForm">Forma activa</param>
        private void OcultarControlesVentana(Form _oForm)
        {
            try
            {
                SAPbouiCOM.Item oItem = null;
                oItem = _oForm.Items.Item(LBL_COMENTARIOS_INACTIVO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(TXT_COMENTARIOS_INACTIVO);
                oItem.Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al ocultar controles *OcultarControlesVentana* : " + ex.Message);
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Asigna los datos de los planes
        /// </summary>
        /// <param name="_oFormaCDU">Forma de campos UDF</param>
        private void AsignarDatosUDF(Form _oFormaCDU)
        {
            try
            {
                _oFormaCDU.Freeze(true);
                _oCmbCategoria = _oFormaCDU.Items.Item(CDU_CMB_CATEGORIA).Specific;
                if (_oCmbCategoria.Selected.Description.ToString().Equals("Contratos"))
                {
                    if (_oFormaCDU.Items.Item(CDU_TRASPASOREL).Visible)
                    {
                        _oTxtTraspasoRel = _oFormaCDU.Items.Item(CDU_TRASPASOREL).Specific;
                        _oTxtSolicitudInterna = _oFormaCDU.Items.Item(CDU_SOLICITUD_INTERNA).Specific;
                        _oTxtTrasp_CodPlan = _oFormaCDU.Items.Item(CDU_CODIGO_PLAN_PREVISION).Specific;
                        _oTxtAct_Descripcion = _oFormaCDU.Items.Item(CDU_NOMBRE_PLAN_PREVISION).Specific;
                        _oTxtInvInicial = _oFormaCDU.Items.Item(CDU_INVINICIAL).Specific;
                        _oTxtComision = _oFormaCDU.Items.Item(CDU_COMISION).Specific;
                        _oTxtPapaleria = _oFormaCDU.Items.Item(CDU_PAPELERIA).Specific;
                        _oTxtImporte_Recibido = _oFormaCDU.Items.Item(CDU_IMPORTE_RECIBIDO).Specific;
                        _oTxtExcedente_InvIni = _oFormaCDU.Items.Item(CDU_EXCEDENTE_INVINI).Specific;
                        _oTxtBono = _oFormaCDU.Items.Item(CDU_BONO).Specific;
                        _oTxtCodCobrador = _oFormaCDU.Items.Item(CDU_CODCOB).Specific;

                        //Documento
                        _oTxtTraspasoRel.Value = (from p in lDatosContratos
                                                  select p.TrasNoDocumento).ElementAt(0).ToString();

                        //SolicitudIntena
                        _oTxtSolicitudInterna.Value = (from p in lDatosContratos
                                                       select p.TrasSerieInterna).ElementAt(0).ToString();

                        //Código plan
                        _oTxtTrasp_CodPlan.Value = (from p in lDatosContratos
                                                    select p.TrasCodigoPlan).ElementAt(0).ToString();

                        //Nombre plan
                        _oTxtAct_Descripcion.Value = (from p in lDatosContratos
                                                      select p.TrasNombrePlan).ElementAt(0).ToString();

                        //Importe pago inicial
                        _oTxtInvInicial.Value = (from p in lDatosContratos
                                                 select p.TrasImportePagoInicial).ElementAt(0).ToString();

                        //Importe comisión
                        _oTxtComision.Value = (from p in lDatosContratos
                                               select p.TrasImporteComision).ElementAt(0).ToString();

                        //Importe papelería
                        _oTxtPapaleria.Value = (from p in lDatosContratos
                                                select p.TrasImportePapeleria).ElementAt(0).ToString();

                        //Importe recibido
                        _oTxtImporte_Recibido.Value = (from p in lDatosContratos
                                                       select p.TrasImporteRecibido).ElementAt(0).ToString();

                        //Excedente inversión inicial
                        _oTxtExcedente_InvIni.Value = (Convert.ToDouble(_oTxtInvInicial.Value) - Convert.ToDouble(_oTxtPapaleria.Value)).ToString();

                        //Bono
                        _oTxtBono.Value = (from p in lDatosContratos
                                           select p.TrasImporteBono).ElementAt(0).ToString();
                        _oTxtBono.ClickPicker();

                        lDatosContratos.Clear();
                    }
                }
                lDatosContratos.Clear();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al asignar los datos del contrato *AsignarDatosUDF* : " + ex.Message);
            }
            finally
            {
                _oFormaCDU.Freeze(false);
            }
        }

        /// <summary>
        /// Crear la factura que corresponde al traspaso relacionado
        /// </summary>
        /// <param name="_oForm">Forma activa</param>
        private string CrearFacturaPago(Form _oForm, string CardCode, string fechaCreacionContrato, string codigoAsistente)
        {
            SAPbouiCOM.Form _UDFForm = null;
            SAPbobsCOM.Documents _oFactura = null;
            SAPbobsCOM.Payments _oPagoFactura = null;
            SAPbobsCOM.Payments _oPagoExcedente = null;
            SAPbobsCOM.Documents _oNotaDeCredito = null;
            SAPbobsCOM.Documents _oNotaCreditoBonoCambioPrecio = null;
            string msgError = null;
            try
            {
                string itemCode = null;
                string solicitudInterna = null;
                string docEntryFactura = null;
                string beneficiarioPago = null;
                double papeleria = 0;
                double excedenteInvInicial = 0;
                double bono = 0;
                double costoTotalPaquete = 0;
                int serialNumber = 0;

                string fechaCreacionConvert = fechaCreacionContrato.Substring(0, 4) + "-" + fechaCreacionContrato.Substring(4, 2) + "-" +
                          fechaCreacionContrato.Substring(6, 2);

                DateTime fechaCreacion = Convert.ToDateTime(fechaCreacionConvert);

                try
                {
                    _UDFForm = _Application.Forms.GetForm("-" + _oForm.TypeEx, _Application.Forms.ActiveForm.TypeCount);
                }
                catch (Exception)
                {
                    msgError = "No están visibles los 'Campos definidos por el usuario'";
                }

                _Company.StartTransaction();

                _oTxtTrasp_CodPlan = _UDFForm.Items.Item(CDU_CODIGO_PLAN_PREVISION).Specific;
                _oTxtInvInicial = _UDFForm.Items.Item(CDU_INVINICIAL).Specific;
                _oTxtSolicitudInterna = _UDFForm.Items.Item(CDU_SOLICITUD_INTERNA).Specific;
                _oTxtBeneficiarioPagoRecibido = _UDFForm.Items.Item(CDU_BENEFICIARIO_PAGO_RECIBIDO).Specific;
                _oTxtBono = _UDFForm.Items.Item(CDU_BONO).Specific;
                _oTxtPapaleria = _UDFForm.Items.Item(CDU_PAPELERIA).Specific;
                _oTxtExcedente_InvIni = _UDFForm.Items.Item(CDU_EXCEDENTE_INVINI).Specific;

                itemCode = _oTxtTrasp_CodPlan.Value.ToString();
                papeleria = Convert.ToDouble(_oTxtPapaleria.Value);
                solicitudInterna = _oTxtSolicitudInterna.Value.ToString();
                beneficiarioPago = _oTxtBeneficiarioPagoRecibido.Value.ToString();
                bono = Convert.ToDouble(_oTxtBono.Value);
                excedenteInvInicial = Convert.ToDouble(_oTxtExcedente_InvIni.Value);
                costoTotalPaquete = Extensor.ObtenerCostoPaquete(solicitudInterna.Substring(0, 6), itemCode, _Company);


                //CREAR LA FACTURA
                _oFactura = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                _oFactura.CardCode = CardCode;
                _oFactura.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                _oFactura.DocDate = Convert.ToDateTime(fechaCreacion);
                _oFactura.Lines.ItemCode = itemCode;
                _oFactura.Lines.Quantity = 1;
                _oFactura.Lines.UnitPrice = costoTotalPaquete; // Extensor.ObtenerCostoPaquete(solicitudInterna.Substring(0, 6), itemCode, _Company);
                _oFactura.Lines.UserFields.Fields.Item("U_Serie").Value = solicitudInterna;
                _oFactura.Lines.TaxCode = Extensor.ObtenerCodigoImpuesto(itemCode, _Company);
                _oFactura.Lines.WarehouseCode = Extensor.ObtenerAlmacen(_Company);

                //Agregar serie a factura
                _oFactura.Lines.SerialNumbers.SetCurrentLine(0);
                _oFactura.Lines.SerialNumbers.InternalSerialNumber = solicitudInterna;
                serialNumber = Extensor.ObtenerNumeroSistema(solicitudInterna, _Company);
                _oFactura.Lines.SerialNumbers.SystemSerialNumber = serialNumber;
                _oFactura.Lines.SerialNumbers.Add();


                if (_oFactura.Add() != 0)
                {
                    msgError = _Company.GetLastErrorDescription();
                }
                else
                {
                    _oRec = null;
                    _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    _oRec.DoQuery(@"UPDATE dbo.OCRD SET U_CostoTotalPaquete = '" + costoTotalPaquete + "'  WHERE CardCode = '" + CardCode + "'");

                    //Se obtiene la referencia de la factura para crear el pago
                    _Company.GetNewObjectCode(out docEntryFactura);
                    _oFactura = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                    _oFactura.GetByKey(Convert.ToInt32(docEntryFactura));

                    // Obtenemos la configuracion del maximo permitido de la Inv Inicial
                    _oRec = null;
                    _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    _oRec.DoQuery("SELECT *  FROM [dbo].[@CONFIG_ADDON]  T0 WHERE T0.[Code] = 3");
                    double MontoMaximoInvIni = 0;
                    if (!(_oRec.EoF))
                        MontoMaximoInvIni = _oRec.Fields.Item("U_Monto_Recomendado").Value;

                    //Se crea el pago referente a la factura (Papelería)
                    if (papeleria > 0 && papeleria <= MontoMaximoInvIni)
                    {
                        //Se crea el pago referente a la factura (Papelería)
                        _oPagoFactura = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);

                        _oPagoFactura.DocDate = Convert.ToDateTime(fechaCreacion);
                        _oPagoFactura.Remarks = ObtenerUsuarioFirmado(_Company.UserName.ToString());
                        _oPagoFactura.JournalRemarks = "Inversión inicial";
                        _oPagoFactura.DocType = SAPbobsCOM.BoRcptTypes.rCustomer;
                        _oPagoFactura.CardCode = _oFactura.CardCode;
                        _oPagoFactura.UserFields.Fields.Item("U_BeneficiarioRecibo").Value = beneficiarioPago;
                        _oPagoFactura.Invoices.DocEntry = Convert.ToInt32(docEntryFactura);
                        _oPagoFactura.CashAccount = Addon.listaCtasSAP.First(x => x.Documento == "INVERSION INICIAL").cuenta; //Extensor.Configuracion.INVERSIONINICIAL.CuentaInversionInicial; 
                        _oPagoFactura.CashSum = papeleria;
                        _oPagoFactura.Invoices.SumApplied = papeleria;
                        _oPagoFactura.UserFields.Fields.Item("U_Es_PagoDirecto").Value = "NO";

                        if (_oPagoFactura.Add() != 0)
                        {
                            msgError = _Company.GetLastErrorDescription();
                        }
                        else
                        {
                            Extensor.InsertarDocentryFacturaCalculoComisiones(_oFactura.CardCode.ToString(), docEntryFactura, _Company);

                            if (excedenteInvInicial > 0)
                            {
                                _oPagoExcedente = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);

                                _oPagoExcedente.DocDate = Convert.ToDateTime(fechaCreacion);
                                _oPagoExcedente.Remarks = ObtenerUsuarioFirmado(_Company.UserName.ToString());
                                _oPagoExcedente.JournalRemarks = "Excedente inversión inicial";
                                _oPagoExcedente.DocType = SAPbobsCOM.BoRcptTypes.rCustomer;
                                _oPagoExcedente.CardCode = _oFactura.CardCode;
                                _oPagoExcedente.UserFields.Fields.Item("U_BeneficiarioRecibo").Value = beneficiarioPago;
                                _oPagoExcedente.Invoices.DocEntry = Convert.ToInt32(docEntryFactura);
                                _oPagoExcedente.CashAccount = Addon.listaCtasSAP.First(x => x.Documento == "EXCEDENTE INV").cuenta; //Extensor.Configuracion.INVERSIONINICIAL.CuentaInversionInicial; 
                                _oPagoExcedente.CashSum = excedenteInvInicial;
                                _oPagoExcedente.Invoices.SumApplied = excedenteInvInicial;
                                _oPagoExcedente.UserFields.Fields.Item("U_Es_PagoDirecto").Value = "NO";

                                if (_oPagoExcedente.Add() != 0)
                                {
                                    msgError = _Company.GetLastErrorDescription();
                                }
                            }

                            _oRec = null;
                            _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            _oRec.DoQuery(@"SELECT U_Reporte_Nombre AS NotaCredito  FROM dbo.[@REPORTES]");

                            if (_oRec.Fields.Item("NotaCredito").Value == "Y")
                            {
                                //NOTA DE CREDITO LIGADA A FACTURA PARA BONO DE 2,000
                                _oNotaCreditoBonoCambioPrecio = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                                _oNotaCreditoBonoCambioPrecio.DocDate = Convert.ToDateTime(fechaCreacion);
                                _oNotaCreditoBonoCambioPrecio.Comments = "Bono por cambio de precio";
                                _oNotaCreditoBonoCambioPrecio.JournalMemo = "Bono por cambio de precio";
                                _oNotaCreditoBonoCambioPrecio.CardCode = CardCode;
                                _oNotaCreditoBonoCambioPrecio.Lines.ItemCode = "BON_PABS-00002";
                                _oNotaCreditoBonoCambioPrecio.Lines.TaxCode = Extensor.ObtenerCodigoImpuesto("BON_PABS-00002", _Company);
                                _oNotaCreditoBonoCambioPrecio.Lines.Quantity = 1;
                                _oNotaCreditoBonoCambioPrecio.Lines.UnitPrice = 2000;
                                _oNotaCreditoBonoCambioPrecio.Lines.AccountCode = Addon.listaCtasSAP.First(x => x.Documento == "NOTA DE CREDITO").cuenta; //Extensor.Configuracion.INVERSIONINICIAL.CuentaInversionInicial;
                                _oNotaCreditoBonoCambioPrecio.Lines.Add();

                                if (_oNotaCreditoBonoCambioPrecio.Add() != 0)
                                {
                                    msgError = _Company.GetLastErrorDescription();
                                    int coderErrror = _Company.GetLastErrorCode();

                                    //_Application.StatusBar.SetText("Ocurrió un error al crear la nota de crédito del Bono por costo de paquete: " + msgError, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                            }

                            if (bono > 0)
                            {

                                //NOTA DE CREDITO LIGADA A FACTURA
                                _oNotaDeCredito = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                                _oNotaDeCredito.DocDate = Convert.ToDateTime(fechaCreacion);
                                _oNotaDeCredito.Comments = "Bono por inversión inicial";
                                _oNotaDeCredito.JournalMemo = "Bono por inversión inicial";
                                _oNotaDeCredito.CardCode = CardCode;
                                _oNotaDeCredito.Lines.ItemCode = "BON_PABS-00001";
                                _oNotaDeCredito.Lines.TaxCode = Extensor.ObtenerCodigoImpuesto("BON_PABS-00001", _Company);
                                _oNotaDeCredito.Lines.Quantity = 1;
                                _oNotaDeCredito.Lines.UnitPrice = bono;
                                _oNotaDeCredito.Lines.AccountCode = Addon.listaCtasSAP.First(x => x.Documento == "NOTA DE CREDITO").cuenta; //Extensor.Configuracion.INVERSIONINICIAL.CuentaInversionInicial;
                                _oNotaDeCredito.Lines.Add();
                                _oNotaDeCredito.Lines.Add();

                                if (_oNotaDeCredito.Add() != 0)
                                {
                                    msgError = _Company.GetLastErrorDescription();
                                    int coderErrror = _Company.GetLastErrorCode();
                                }
                            }
                        }
                    }
                    else
                    {
                        msgError = "El monto de papeleria ingresado exede el maximo configurado. Papeleria : $" + papeleria + ", Maximo permitido: $" + MontoMaximoInvIni;
                        Extensor.InsertarDocentryFacturaCalculoComisiones(_oFactura.CardCode.ToString(), docEntryFactura, _Company);
                    }
                }
            }
            catch (Exception ex)
            {              
                msgError = ex.Message;
            }
            finally
            {
                try
                {
                    if (!string.IsNullOrEmpty(msgError))
                    {
                        _oRec = null;
                        _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        _oRec.DoQuery(@"DELETE  dbo.[@CALCULO_COMISIONES]
                                                    WHERE   U_Contrato = '" + CardCode + "' UPDATE  dbo.[@AYUDAS] " +
                                                        "SET     U_NumeroAyuda = U_NumeroAyuda - 1 " +
                                                        "WHERE   U_CodigoAsistente = '" + codigoAsistente + "'");
                    }

                    if (string.IsNullOrEmpty(msgError))
                        _Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    else
                        _Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                    if (_oRec != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oRec);
                    }
                    if (_oFactura != null)
                    {
                        GC.SuppressFinalize(_oFactura);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oFactura);
                        _oFactura = null;
                    }

                    if (_oPagoFactura != null)
                    {
                        GC.SuppressFinalize(_oPagoFactura);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oPagoFactura);
                        _oPagoFactura = null;
                    }

                    if (_oPagoExcedente != null)
                    {
                        GC.SuppressFinalize(_oPagoExcedente);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oPagoExcedente);
                        _oPagoExcedente = null;
                    }
                    if (_oNotaDeCredito != null)
                    {
                        GC.SuppressFinalize(_oNotaDeCredito);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oNotaDeCredito);
                        _oNotaDeCredito = null;
                    }

                    if (_oNotaCreditoBonoCambioPrecio != null)
                    {
                        GC.SuppressFinalize(_oNotaCreditoBonoCambioPrecio);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oNotaCreditoBonoCambioPrecio);
                        _oNotaCreditoBonoCambioPrecio = null;

                    }
                    GC.Collect();
                    ClearMemory();
                }
                catch (Exception)
                {
                }


            }
            return msgError;
        }


        /// <summary>
        /// Bloquea los controles cuando la forma se encuentra en modo Vista
        /// </summary>
        /// <param name="_oForm">Forma activa</param>
        private void InhabilitarControlesDefinipos(Form _oForm)
        {
            try
            {
                _oForm.Freeze(true);
                _oForm.Items.Item(TXT_MOTIVO_SOCIO_INACTIVO).Enabled = false;
                _oForm.Items.Item(TXT_COMENTARIOS_MOTIVO).Enabled = false;
            }
            catch (Exception ex)
            {
                _Application.StatusBar.SetText("Error al bloquear controles *InhabilitarControlesDefinipos* : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Se insertan los datos en la tabla de @CALCULO_COMISIONES al momento de crear el socio de negocio
        /// </summary>
        /// <param name="traspasoRel">Traspaso relacionado</param>
        /// <param name="solicitud">Solicitud</param>
        /// <param name="solicitudInt">Solicitud Interna</param>
        /// <param name="codigoCobrador">Código del cobrador</param>
        /// <param name="nombreCobrador">Nombre cobrador</param>
        /// <param name="codigoPlan">Código del plan</param>
        /// <param name="inversionInicial">Inversión inicial</param>
        private void InsertarCalculoComisiones(string traspasoRel, string solicitud, string solicitudInt, string codigoCobrador, string nombreCobrador,
                                                string codigoPlan, string inversionInicial, string bono, string comision, string codigoAsistenteContrato,
                                                string nombreAsistenteContrato, string esquemaPago)
        {
            try
            {
                Extensor.DatosDetalleComsiones DetalleComisiones = new Extensor.DatosDetalleComsiones();
                Extensor.DatosTransferencia DatosTransfer = new Extensor.DatosTransferencia();

                string empresa = Extensor.ObtenerEmpresa(solicitudInt.Substring(0, 6), codigoPlan, _Company);
                string validarEsquemaAsistente = "";

                DatosTransfer = Extensor.ObtenerDatosTransferencia(solicitudInt, _Company);
                DetalleComisiones = Extensor.ObtenerDatosDetalleComisiones(empresa, codigoAsistenteContrato, codigoPlan, solicitudInt.Substring(0, 6), _Company);

                double fideicomiso = DetalleComisiones.montoFideicomiso;// -Convert.ToDouble(bono);
                double comisionAsistente = 0;

                if (string.IsNullOrEmpty(esquemaPago))
                {
                    esquemaPago = Extensor.ObtenerEsquemaPago(codigoAsistenteContrato, _Company);
                }
                else
                {
                    validarEsquemaAsistente = Extensor.ObtenerEsquemaPago(codigoAsistenteContrato, _Company);
                }
                if (validarEsquemaAsistente.Contains("SUELDO"))
                {
                    if (esquemaPago.Contains("SUELDO"))
                    {
                        int ObtenerNoAyuda = Extensor.ObtenerAyudasAsistente(codigoAsistenteContrato, _Company);
                        int ObtenerNumeroContratos = Extensor.ObtenerNumeroContratos(codigoAsistenteContrato, _Company);

                        if (ObtenerNoAyuda <= ObtenerNumeroContratos)
                        {
                            comisionAsistente = 0;
                            Extensor.ActualizarEsquemaComision(solicitud, "SUELDO", _Company);
                        }
                        else if (ObtenerNoAyuda > ObtenerNumeroContratos)
                        {
                            Extensor.ActualizarEsquemaComision(solicitud, "COMISION", _Company);
                            comisionAsistente = DetalleComisiones.montoAsistenteSocial;
                        }
                    }
                    else
                    {
                        comisionAsistente = DetalleComisiones.montoAsistenteSocial;
                        Extensor.ActualizarEsquemaComision(solicitud, "COMISION", _Company);
                    }
                }
                else
                {
                    if (esquemaPago.Contains("SUELDO"))
                    {
                        int ObtenerNoAyuda = Extensor.ObtenerAyudasAsistente(codigoAsistenteContrato, _Company);
                        int ObtenerNumeroContratos = Extensor.ObtenerNumeroContratos(codigoAsistenteContrato, _Company);

                        if (ObtenerNoAyuda <= ObtenerNumeroContratos)
                        {
                            comisionAsistente = 0;
                            Extensor.ActualizarEsquemaComision(solicitud, "SUELDO", _Company);
                        }
                        else if (ObtenerNoAyuda > ObtenerNumeroContratos)
                        {
                            Extensor.ActualizarEsquemaComision(solicitud, "COMISION", _Company);
                            comisionAsistente = DetalleComisiones.montoAsistenteSocial;
                        }
                    }
                    else
                    {
                        comisionAsistente = DetalleComisiones.montoAsistenteSocial;
                        Extensor.ActualizarEsquemaComision(solicitud, "COMISION", _Company);
                    }

                }

                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _oRec.DoQuery(@"INSERT  INTO dbo.[@CALCULO_COMISIONES]
                                        (
                                            DocEntry,
                                            U_Nombre_Cobrador,
                                            U_Codigo_Cobrador,
                                            U_Codigo_Asistente,
                                            U_Asistente_Social,
                                            U_Contrato,
                                            U_Asis_Social,
                                            U_Recomendado,
                                            U_Lider,
                                            U_Supervisor,
                                            U_Coordinador,
                                            U_Gerente,
                                            U_Fideicomiso,
                                            U_Empresa,
                                            U_CodigoRecomendado,
                                            U_Nom_Recomendado,
                                            U_CodigoLider,
                                            U_Nom_Lider,
                                            U_CodigoSupervisor,
                                            U_Nom_Supervisor,
                                            U_CodigoCoordinador,
                                            U_Nom_Coordinador,
                                            U_CodigoCoordinador2,
                                            U_Nom_Coordinador2,
                                            U_CodigoGerente,
                                            U_Nom_Gerente,
                                            U_BonoAsistente,
                                            U_Bono,
                                            U_Inv_Inicial,
                                            U_Coordinador2,
                                            U_BonoCoordinador,
                                            U_BonoCoordinador2,
                                            U_His_Inv_Inicial,
                                            U_His_Recomendado,
                                            U_His_Asis_Social,
                                            U_His_BonoAsis,
                                            U_His_Bono,
                                            U_His_Lider,
                                            U_His_Supervisor,
                                            U_His_Coordinador,
                                            U_His_Coordinador2,
                                            U_His_Gerente,
                                            U_His_Fideicomiso,
                                            U_His_BonoCoord,
                                            U_His_BonoCoord2,
                                            U_ContratoInterno,
                                            U_DocEntryTransfer,
                                            U_FechaCreacion
                                        )
                                VALUES  (
                                            ( SELECT  ISNULL(MAX(CAST(DocEntry AS INT)), 0) + 1 AS DocEntry
                                            FROM    dbo.[@CALCULO_COMISIONES] WITH ( UPDLOCK )
                                            ),
	                                /* U_Nombre_Cobrador - nvarchar(20) */
                                            '" + nombreCobrador + "', " +
                    /* U_Codigo_Cobrador - nvarchar(100) */
                               "'" + codigoCobrador + "', " +
                    /* U_Codigo_Asistente - nvarchar(20) */
                               "'" + codigoAsistenteContrato + "', " +
                    /* U_Asistente_Social - nvarchar(100) */
                               "'" + nombreAsistenteContrato + "', " +
                    /* U_Contrato - nvarchar(20) */
                               "'" + solicitud + "', " +
                    /* U_Asis_Social - numeric(19, 6) */
                               "'" + comisionAsistente + "', " +
                    /* U_Recomendado - numeric(19, 6) */
                               "'" + DetalleComisiones.montoRecomendado + "', " +
                    /* U_Lider - numeric(19, 6) */
                               "'" + DetalleComisiones.montoLider + "', " +
                    /* U_Supervisor - numeric(19, 6) */
                               "'" + DetalleComisiones.montoSupervisor + "', " +
                    /* U_Coordinador - numeric(19, 6) */
                               "'" + DetalleComisiones.montoCoordinador + "', " +
                    /* U_Gerente - numeric(19, 6) */
                               "'" + DetalleComisiones.montoGerente + "', " +
                    /* U_Fideicomiso - numeric(19, 6) */
                               "'" + fideicomiso + "', " +
                    /* U_Empresa - nvarchar(50) */
                               "'" + empresa + "', " +
                    /* U_CodigoRecomendado - nvarchar(20) */
                               "'" + DetalleComisiones.codigoRecomendado + "', " +
                    /* U_Nom_Recomendado - nvarchar(100) */
                               "'" + DetalleComisiones.nombreRecomendado + "', " +
                    /* U_CodigoLider - nvarchar(20) */
                               "'" + DetalleComisiones.codigoLider + "', " +
                    /* U_Nom_Lider - nvarchar(100) */
                               "'" + DetalleComisiones.nombreLider + "', " +
                    /* U_CodigoSupervisor - nvarchar(20) */
                               "'" + DetalleComisiones.codigoSupervisor + "', " +
                    /* U_Nom_Supervisor - nvarchar(100) */
                               "'" + DetalleComisiones.nombreSupervisor + "', " +
                    /* U_CodigoCoordinador - nvarchar(20) */
                               "'" + DetalleComisiones.codigoCoordinador + "', " +
                    /* U_Nom_Coordinador - nvarchar(100) */
                               "'" + DetalleComisiones.nombreCoordinador + "', " +
                    /* U_CodigoCoordinador2 - nvarchar(20) */
                               "'" + DetalleComisiones.codigoCoordinador2 + "', " +
                    /* U_Nom_Coordinador2 - nvarchar(100) */
                               "'" + DetalleComisiones.nombreCoordinador2 + "', " +
                    /* U_CodigoGerente - nvarchar(20) */
                               "'" + DetalleComisiones.codigoGerente + "', " +
                    /* U_Nom_Gerente - nvarchar(100) */
                               "'" + DetalleComisiones.nombreGerente + "', " +
                    /* U_BonoAsistente - numeric(19, 6) */
                               "'" + DetalleComisiones.montoBono + "', " +
                    /* U_Bono - numeric(19, 6) */
                               "'" + bono + "', " +
                    /* U_Inv_Inicial - numeric(19, 6) */
                               "'" + DetalleComisiones.montoInvInicial + "', " +
                    /* U_Coordinador2 - numeric(19, 6) */
                               "'" + DetalleComisiones.montoCoordinador2 + "', " +
                    /* U_BonoCoordinador - numeric(19,6) */
                               "'" + DetalleComisiones.montoBonoCoordinador + "', " +
                    /* U_BonoCoordinador2 - numeric(19,6) */
                               "'" + DetalleComisiones.montoBonoCoordinador2 + "', " +
                    /* U_His_Inv_Inicial - numeric(19, 6) */
                               "'" + DetalleComisiones.montoInvInicial + "', " +
                    /* U_His_Recomendado - numeric(19, 6) */
                               "'" + DetalleComisiones.montoRecomendado + "', " +
                    /* U_His_Asis_Social - numeric(19, 6) */
                               "'" + comisionAsistente + "', " +
                    /* U_His_BonoAsis - numeric(19, 6) */
                               "'" + DetalleComisiones.montoBono + "', " +
                    /* U_His_Bono - numeric(19, 6) */
                               "'" + bono + "', " +
                    /* U_His_Lider - numeric(19, 6) */
                               "'" + DetalleComisiones.montoLider + "', " +
                    /* U_His_Supervisor - numeric(19, 6) */
                               "'" + DetalleComisiones.montoSupervisor + "', " +
                    /* U_His_Coordinador - numeric(19, 6) */
                               "'" + DetalleComisiones.montoCoordinador + "', " +
                    /* U_His_Coordinador2 - numeric(19, 6) */
                               "'" + DetalleComisiones.montoCoordinador2 + "', " +
                    /* U_His_Gerente - numeric(19, 6) */
                               "'" + DetalleComisiones.montoGerente + "', " +
                    /* U_His_Fideicomiso - numeric(19, 6) */
                               "'" + fideicomiso + "', " +
                    /* U_His_BonoCoord - numeric(19, 6) */
                               "'" + DetalleComisiones.montoBonoCoordinador + "', " +
                    /* U_His_BonoCoord2 - numeric(19, 6) */
                               "'" + DetalleComisiones.montoBonoCoordinador2 + "', " +
                    /* U_ContratoInterno - nvarchar(20) */
                               "'" + solicitudInt + "', " +
                    /* U_DocEntryTransfer - nvarchar(20) */
                               "'" + DatosTransfer.docEntryTransferencia + "', " +
                    /* U_FechaCreacion - DateTime */
                               "GETDATE())");

                DatosTransfer = null;
                DetalleComisiones = null;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al insertar el calculo de las comisiones *InsertarCalculoComisiones* : " + ex.Message);
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
        /// Se actualizan los datos en la tabla de @CALCULO_COMISIONES al momento de crear el socio de negocio
        /// </summary>
        /// <param name="solicitud">Solicitud</param>
        /// <param name="solicitudInt">Solicitud Interna</param>
        /// <param name="codigoCobrador">Código del cobrador</param>
        /// <param name="nombreCobrador">Nombre cobrador</param>
        private void ActualizarCobrador(string solicitud, string solicitudInt, string codigoCobrador, string nombreCobrador)
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"UPDATE  dbo.[@CALCULO_COMISIONES]
                                SET     U_Nombre_Cobrador = '" + nombreCobrador + "', " +
                                        "U_Codigo_Cobrador = '" + codigoCobrador + "' " +
                                "WHERE   U_Contrato = '" + solicitud + "' " +
                                        "AND U_ContratoInterno = '" + solicitudInt + "'");
            }
            catch (Exception ex)
            {
                throw new Exception("Error al actualizar el cobrador *ActualizarCobrador* : " + ex.Message);
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
        /// Valida que los datos de la fecha de nacimiento estén seleccionados
        /// </summary>
        /// <param name="dia">Día</param>
        /// <param name="mes">Mes</param>
        /// <param name="year">Año</param>
        /// <returns>Mensaje de error</returns>
        private string ValidarFechaNacimiento(ComboBox dia, ComboBox mes, ComboBox year)
        {
            try
            {
                string msgError = null;

                if (dia != null)
                {
                    if (string.IsNullOrEmpty(dia.Value))
                    {
                        return msgError = "Seleccione el día";
                    }
                }
                else
                {
                    return msgError = "Seleccione el día";
                }

                if (mes != null)
                {
                    if (string.IsNullOrEmpty(mes.Value))
                    {
                        return msgError = "Seleccione el mes";
                    }
                }
                else
                {
                    return msgError = "Seleccione el mes";
                }

                if (year != null)
                {
                    if (string.IsNullOrEmpty(year.Value))
                    {
                        return msgError = "Selecciones el año";
                    }
                }
                else
                {
                    return msgError = "Selecciones el año";
                }

                return msgError;
            }
            catch (Exception ex)
            {
                throw new Exception("No se pudo validar la fecha de nacimiento *ValidarFechaNacimiento* : " + ex.Message);
            }
        }

        /// <summary>
        /// Libera la memoria de la aplicación de SAP
        /// </summary>
        /// <param name="procHandle">Proceso asociado</param>
        /// <param name="min">Mínimo</param>
        /// <param name="max">Máximo</param>
        /// <returns>Proceso</returns>
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern bool SetProcessWorkingSetSize(IntPtr procHandle, Int32 min, Int32 max);

        /// <summary>
        /// Libera la memoria
        /// </summary>
        private static void ClearMemory()
        {
            Process mm = null;
            mm = Process.GetCurrentProcess();
            SetProcessWorkingSetSize(mm.Handle, -1, -1);
        }

        #endregion
    }
}
