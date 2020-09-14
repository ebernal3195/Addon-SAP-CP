﻿using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace AddonFNR.BL
{
    class clsSolicitudDePlanes : ComportaForm
    {
        #region CONSTANTES

        private const int FRM_SOLICITUD_DE_PLANES = 540000988;

        private const string GRID_ARTICULOS = "38";
        private const string COLUMNA_SERIE_INICIO = "U_SerieIni";
        private const string COLUMNA_SERIE_FIN = "U_SerieFin";
        private const int CHAR_PRESS_ENTER = 13;

        //ENCABEZADO
        private const string LBL_PERSONA_CONTACTO = "83";
        private const string LBL_NUMERO_REFERENCIA = "15";
        private const string LBL_NUMERO_GRUPO = "540002107";
        private const string LBL_FECHA_CONTABILIZACION = "11";
        private const string LBL_FECHA_DOCUMENTO = "86";

        private const string TXT_NUMERO_REFERENCIA = "14";
        private const string TXT_NUMERO_GRUPO = "540002109";
        private const string TXT_FECHA_CONTABILIZACION = "10";
        private const string TXT_FECHA_DOCUMENTO = "46";

        private const string CMB_NUMERO_GRUPO = "540002108";
        private const string CMB_PERSONA_CONTACTO = "85";

        private const string ICO_PERSONA_CONTACTO = "80";
        private const string ICO_NUMERO_GRUPO = "540002110";


        //PESTAÑAS
        private const string FLD_FINANZAS = "138";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForm = null;
        private static bool _oSolicitudPlanes = false;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de la solicitud de traslado
        /// </summary>
        /// <param name="_Application">Objeto de la conexión de SAP</param>
        /// <param name="_Company">Objeto de la empresa</param>
        /// <param name="form">Nombre de la forma</param>
        public clsSolicitudDePlanes(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oSolicitudPlanes == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                setEventos();
                _oSolicitudPlanes = true;
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
                if (pVal.BeforeAction == false && pVal.FormType == FRM_SOLICITUD_DE_PLANES)
                {
                    if (pVal.EventType == BoEventTypes.et_FORM_RESIZE)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        OcultarControlesVentana(_oForm);
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
                    {
                        _Application.ItemEvent -= new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                        Dispose();
                        application = null;
                        company = null;
                        _oSolicitudPlanes = false;
                        Addon.typeList.RemoveAll(p => p._forma == formID);
                        return;

                    }
                }

                if (pVal.BeforeAction == true && pVal.FormType == FRM_SOLICITUD_DE_PLANES)
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
                throw new Exception("Error en método 'eventos' *clsSolicitudDePlanes* : " + ex.Message);
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
                oItem = _oForm.Items.Item(LBL_PERSONA_CONTACTO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_NUMERO_REFERENCIA);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_NUMERO_GRUPO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_FECHA_CONTABILIZACION);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(LBL_FECHA_DOCUMENTO);
                oItem.Visible = false;

                //TEXBOX
                oItem = _oForm.Items.Item(TXT_NUMERO_REFERENCIA);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(TXT_NUMERO_GRUPO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(TXT_FECHA_CONTABILIZACION);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(TXT_FECHA_DOCUMENTO);
                oItem.Visible = false;

                //COMBOS
                oItem = _oForm.Items.Item(CMB_NUMERO_GRUPO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(CMB_PERSONA_CONTACTO);
                oItem.Visible = false;

                //ICONOS CHOOSE
                oItem = _oForm.Items.Item(ICO_NUMERO_GRUPO);
                oItem.Visible = false;
                oItem = _oForm.Items.Item(ICO_PERSONA_CONTACTO);
                oItem.Visible = false;

                //FOLDER
                oItem = _oForm.Items.Item(FLD_FINANZAS);
                oItem.Visible = false;                 
            }
            catch (Exception ex)
            {
                throw new Exception("Error al ocultar controles *OcultarControlesVentana* : " + ex.Message);
            }
        }

        #endregion

    }
}