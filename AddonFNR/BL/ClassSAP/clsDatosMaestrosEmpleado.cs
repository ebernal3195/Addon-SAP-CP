using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace AddonFNR.BL
{
    class clsDatosMaestrosEmpleado : ComportaForm
    {
        #region CONSTANTES

        private const int FRM_DATOS_MAESTROS_EMPLEADO = 60100; 

        //ENCABEZADO
        private const string LBL_TELEFONO_OFICINA = "76";
        private const string LBL_RFC = "lblRfc";

        private const string TXT_RFC = "U_RFC";
        private const string TXT_FECHA_NACIMIENTO = "113";
        private const string TXT_NOMBRE_EMPLEADO = "39";
        private const string TXT_APELLIDO_EMPLEADO = "37";       

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForm = null;
        private static bool _oDatosMestrosEmpleado = false;
        private SAPbouiCOM.StaticText _oLblRfc = null;
        private SAPbouiCOM.EditText _oTxtRfc = null;
        private SAPbouiCOM.EditText _oTxtFechaNacimiento = null;
        private SAPbouiCOM.EditText _oTxtNombreEmpleado = null;
        private SAPbouiCOM.EditText _oTxtApellidoEmpleado = null;
        
        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de datos maestros empleado
        /// </summary>
        /// <param name="_Application">Objeto de la conexión de SAP</param>
        /// <param name="_Company">Objeto de la empresa</param>
        /// <param name="form">Nombre de la forma</param>
        public clsDatosMaestrosEmpleado(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oDatosMestrosEmpleado == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                setEventos();
                _oDatosMestrosEmpleado = true;
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
                _Application.MessageBox("Ocurrió un error en ItemEvent : " + ex.Message);
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
                if (pVal.BeforeAction == false && pVal.FormType == FRM_DATOS_MAESTROS_EMPLEADO)
                {
                    if (pVal.EventType == BoEventTypes.et_FORM_RESIZE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                       // CrearCampoRFC(_oForm);
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
                    {
                        _Application.ItemEvent -= new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                        Dispose();
                        application = null;
                        company = null;
                        _oDatosMestrosEmpleado = false;
                        Addon.typeList.RemoveAll(p => p._forma == formID);
                        return;
                    }

                    if(pVal.EventType == BoEventTypes.et_LOST_FOCUS && pVal.ItemUID == TXT_FECHA_NACIMIENTO)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        if(_oForm.Visible == true)
                        {
                            if (_oForm.Mode == BoFormMode.fm_ADD_MODE || _oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                                _oTxtFechaNacimiento = _oForm.Items.Item(TXT_FECHA_NACIMIENTO).Specific;
                                _oTxtNombreEmpleado = _oForm.Items.Item(TXT_NOMBRE_EMPLEADO).Specific;
                                _oTxtApellidoEmpleado = _oForm.Items.Item(TXT_APELLIDO_EMPLEADO).Specific;
                                _oTxtRfc = _oForm.Items.Item(TXT_RFC).Specific;

                                if (!string.IsNullOrEmpty(_oTxtNombreEmpleado.Value.ToString()))
                                {
                                    if (!string.IsNullOrEmpty(_oTxtApellidoEmpleado.Value.ToString()))
                                    {
                                        if (!String.IsNullOrEmpty(_oTxtFechaNacimiento.Value.ToString()))
                                        {
                                            _oForm.Freeze(true);
                                            string[] separarNombre = null;
                                            string dia = null;
                                            string mes = null;
                                            string year = null;

                                            var nombreSeparado = SepararNombreApllidos_Fomato.FormatoTextoRFC(_oTxtNombreEmpleado.Value.ToString() + " " + _oTxtApellidoEmpleado.Value.ToString());
                                            _oTxtNombreEmpleado.Value = SepararNombreApllidos_Fomato.FormatoMayusculas(_oTxtNombreEmpleado.Value.ToString());
                                            _oTxtApellidoEmpleado.Value = SepararNombreApllidos_Fomato.FormatoMayusculas(_oTxtApellidoEmpleado.Value.ToString());

                                            year = _oTxtFechaNacimiento.Value.Substring(0, 4);
                                            mes = _oTxtFechaNacimiento.Value.Substring(4, 2);
                                            dia = _oTxtFechaNacimiento.Value.Substring(6, 2);

                                            separarNombre = nombreSeparado.ToString().Split('@');
                                            _oTxtRfc.Value = SepararNombreApllidos_Fomato.GenerarRFC(separarNombre, dia, mes, year);
                                            _oForm.Freeze(false);
                                        }
                                    }
                                    else
                                    {
                                        _Application.MessageBox("Capture el apellido del empleado");
                                    }
                                }
                                else
                                {
                                    _Application.MessageBox("Capture el nombre del empleado");
                                }
                            }
                        }
                    }
                }                
            }
            catch (Exception ex)
            {
                throw new Exception("Error en método 'eventos' *clsDatosMaestrosEmpleado* : " + ex.Message);
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
        /// Crea el campo de RFC en la ventana de Datos maestros  empleado
        /// </summary>
        /// <param name="_oForm"></param>
        private void CrearCampoRFC(Form _oForm)
        {
            SAPbouiCOM.Item newItem = null;

            try
            {
                try
                {
                    string s = _oForm.Items.Item(TXT_RFC).UniqueID;
                }
                catch (Exception)
                {
                    _oForm.Freeze(true);
 
                    //LABEL RFC
                    SAPbouiCOM.Item _oLblTelOficina = null;
                    _oLblTelOficina = _oForm.Items.Item(LBL_TELEFONO_OFICINA);
                    newItem = _oForm.Items.Add(LBL_RFC, BoFormItemTypes.it_STATIC);
                    newItem.Left = _oLblTelOficina.Left;
                    newItem.Top = _oLblTelOficina.Top - 16;
                    newItem.Width = 50;
                    newItem.ToPane = 0;
                    newItem.FromPane = 0;
                    _oLblRfc = newItem.Specific;
                    _oLblRfc.Caption = "RFC";

                    //EDITTEXT RFC
                    SAPbouiCOM.Item _oLabelRFC = null;
                    _oLabelRFC = _oForm.Items.Item(LBL_RFC);
                    newItem = _oForm.Items.Add(TXT_RFC, BoFormItemTypes.it_EDIT);
                    newItem.Left = _oLabelRFC.Left + 102;
                    newItem.Top = _oLabelRFC.Top;
                    newItem.Width = 90;
                    newItem.ToPane = 0;
                    newItem.FromPane = 0;
                    _oTxtRfc = (SAPbouiCOM.EditText)newItem.Specific;
                    _oTxtRfc.DataBind.SetBound(true, "OHEM", "U_RFC");
                    _oLabelRFC.LinkTo = newItem.UniqueID;
                    
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al crear campos de usuario *CrearCampoRFC* : " + ex.Message);
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }


        #endregion
         
    }
}
