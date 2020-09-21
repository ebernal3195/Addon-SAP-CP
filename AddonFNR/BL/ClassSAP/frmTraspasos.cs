using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddonFNR.BL
{
    class frmTraspasos : ComportaForm
    {
        #region CONSTANTES

        private const string FRM_TRASPASOS = "frmTraspasos";

        private const string LBL_ORIGEN_NOMBRE_SN = "lblOrigen";
        private const string LBL_DESTINO_NOMBRE_SN = "lblDesti";

        private const string TXT_CONTRATO_ORIGEN = "txtConOri";
        private const string TXT_CONTRATO_DESTINO = "txtConDes";
        private const string TXT_PLAN_ORIGEN = "txtPlanO";
        private const string TXT_PLAN_DESTINO = "txtPlanD";
        private const string TXT_COSTO_ORIGEN = "txtCostoO";
        private const string TXT_COSTO_DESTINO = "txtCostoD";
        private const string TXT_SALDO_ORIGEN = "txtSaldoO";
        private const string TXT_SALDO_DESTINO = "txtSaldoD";
        private const string TXT_MONTO = "txtMonto";
        private const string TXT_EMPRESA_ORIGEN = "txtEmpreO";
        private const string TXT_EMPRESA_DESTINO = "txtEmpreD";
        private const string TXT_FACTURA_ORIGEN = "txtFacOri";
        private const string TXT_FACTURA_DESTINO = "txtFacDes";

        private const string BTN_PAGOS_ORIGEN = "btnPagosO";
        private const string BTN_PAGOS_DESTINO = "btnPagosD";
        private const string BTN_CERRAR = "btnCerrar";
        private const string BTN_TRASPASO = "btnTraspa";
        private const string BTN_SI = "btnSi";
        private const string BTN_NO = "btnNo";

        private const string GRD_TRASPASO_ORIGEN = "grdTraspO";
        private const string GRD_TRASPASO_DESTINO = "grdTraspD";

        private const string DT_TRASPASO_ORIGEN = "dtTraspasoOrigen";
        private const string DT_TRASPASO_DESTINO = "dtTraspasoDestino";

        private const string BTN_SAP_BUSCAR = "1281";
        private const string BTN_SAP_CREAR = "1282";
        private const string LINK_DM_SOCIO_NEGOCIO = "2";
        private const string LINK_PAGO_EFECTUADO = "24";

        private const string BTN_ORIGEN = "ButtonOrig";
        private const string BTN_DESTINO = "BunttonDes";

        private const string COLUMNA_SELECCIONAR = "Sel";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForma = null;
        private static bool _oTraspasos = false;
        private SAPbobsCOM.Recordset _oRec = null;

        private SAPbouiCOM.Grid _oGridTraspasoOrigen = null;
        private SAPbouiCOM.Grid _oGridTraspasoDestino = null;

        private SAPbouiCOM.EditText _oContratoOrigen = null;
        private SAPbouiCOM.EditText _oContratoDestino = null;
        private SAPbouiCOM.EditText _oPlanOrigen = null;
        private SAPbouiCOM.EditText _oPlanDestino = null;
        private SAPbouiCOM.EditText _oCostoOrigen = null;
        private SAPbouiCOM.EditText _oCostoDestino = null;
        private SAPbouiCOM.EditText _oSaldoOrigen = null;
        private SAPbouiCOM.EditText _oSaldoDestino = null;
        private SAPbouiCOM.EditText _oEmpresaOrigen = null;
        private SAPbouiCOM.EditText _oEmpresaDestino = null;
        private SAPbouiCOM.EditText _oMontoSeleccionado = null;
        private SAPbouiCOM.EditText _oFacturaOrigen = null;
        private SAPbouiCOM.EditText _oFacturaDestino = null;

        private SAPbouiCOM.StaticText _oLblNombreOrigen = null;
        private SAPbouiCOM.StaticText _oLblNombreDestino = null;

        private int ContadorSeleccionados = 0;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de traspasos
        /// </summary>
        /// <param name="_Application">Este es el objeto raíz de la API de interfaz de usuario
        ///                             lo que refleja la cual aplicación SAP Business One en el que se realiza 
        ///                             la conexión</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        public frmTraspasos(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oTraspasos == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                showForm(formID);
                inicializarComponentes();
                setEventos();
                _oTraspasos = true;
            }
        }

        #endregion

        #region EVENTOS

        /// <summary>
        /// Ejecuta evento de la forma
        /// </summary>
        /// <param name="FormUID">Nombre o ID de la forma</param>
        /// <param name="pVal">Propiedades de la forma</param>
        /// <param name="bubbleEvent">Evento</param>
        public void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                if (_oTraspasos != false && pVal.FormType != FormTypeMenu && formID == FormUID)
                {
                    eventos(FormUID, ref pVal, out bubbleEvent);
                }
            }
            catch (Exception ex)
            {
                _Application.MessageBox("Error en ItemEvent: " + ex.Message);
            }

        }

        /// <summary>
        /// Liberar recursos
        /// </summary>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Ejecución de eventos de la forma activa
        /// </summary>
        /// <param name="FormUID">Nombre o ID de la forma</param>
        /// <param name="pVal">Propiedades de la forma</param>
        /// <param name="bubbleEvent">Evento</param>
        public override void eventos(string FormUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.FormUID == formID)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                    string sCFL_ID = null;
                    sCFL_ID = oCFLEvento.ChooseFromListUID;
                    SAPbouiCOM.Form oForm = null;
                    oForm = _Application.Forms.Item(FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = null;
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                    if (oCFLEvento.BeforeAction == false)
                    {
                        SAPbouiCOM.DataTable oDataTable = null;
                        oDataTable = oCFLEvento.SelectedObjects;
                        string val = null;
                        try
                        {
                            val = System.Convert.ToString(oDataTable.GetValue(0, 0));
                        }
                        catch (Exception ex)
                        {

                        }
                        if ((pVal.ItemUID == BTN_ORIGEN))
                        {
                            oForm.DataSources.UserDataSources.Item("dsOrigen").ValueEx = val;
                        }
                        if (pVal.ItemUID == BTN_DESTINO)
                        {
                            oForm.DataSources.UserDataSources.Item("dsDestino").ValueEx = val;
                        }
                    }
                }

                if (pVal.FormUID == formID && pVal.BeforeAction == false)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
                    {
                        _Application.ItemEvent -= new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                        _Application.MenuEvent -= new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                        Dispose();
                        application = null;
                        company = null;
                        _oTraspasos = false;
                        Addon.typeList.RemoveAll(p => p._forma == formID);
                        return;
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if (pVal.ItemUID == BTN_SI)
                        {
                            SeleccionarTodos("Y");
                        }
                        if (pVal.ItemUID == BTN_NO)
                        {
                            SeleccionarTodos("N");
                        }
                        if (pVal.ItemUID == BTN_TRASPASO)
                        {
                            GenerarTraspasos();
                        }
                        if (pVal.ItemUID == BTN_CERRAR)
                        {
                            _oForma.Close();
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        if (pVal.ItemUID == BTN_ORIGEN)
                        {
                            _oContratoOrigen = _oForma.Items.Item(TXT_CONTRATO_ORIGEN).Specific;

                            if (!string.IsNullOrEmpty(_oContratoOrigen.Value.ToString()))
                            {
                                ObtenerDatosOrigen();
                            }
                        }

                        if (pVal.ItemUID == BTN_DESTINO)
                        {
                            _oContratoDestino = _oForma.Items.Item(TXT_CONTRATO_DESTINO).Specific;

                            if (!string.IsNullOrEmpty(_oContratoDestino.Value.ToString()))
                            {
                                ObtenerDatosDestino();
                            }
                        }                      
                    }
                }

                if (pVal.FormUID == formID && pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == GRD_TRASPASO_ORIGEN && pVal.ColUID == COLUMNA_SELECCIONAR && pVal.EventType == BoEventTypes.et_CLICK || pVal.EventType == BoEventTypes.et_DOUBLE_CLICK)
                    {
                        string seleccionado = _oGridTraspasoOrigen.DataTable.GetValue(COLUMNA_SELECCIONAR, pVal.Row);

                        if (seleccionado == "Y")
                        {
                            _oMontoSeleccionado.Value = Convert.ToString(Math.Round(Convert.ToDouble(_oMontoSeleccionado.Value.ToString()) + Convert.ToDouble(_oGridTraspasoOrigen.DataTable.GetValue("Monto", pVal.Row)), 2));
                        }
                        else if (seleccionado == "N")
                        {
                            _oMontoSeleccionado.Value = Convert.ToString(Math.Round(Convert.ToDouble(_oMontoSeleccionado.Value.ToString()) - Convert.ToDouble(_oGridTraspasoOrigen.DataTable.GetValue("Monto", pVal.Row)), 2));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en evento *clsTraspasos* : " + ex.Message);
            }
        }

        /// <summary>
        /// Ejecución de eventos del menú
        /// </summary>
        /// <param name="pVal">Propiedad</param>
        /// <param name="BubbleEvent">true/false</param>
        private void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            try
            {
                BubbleEvent = true;

                //Valida cuando se presione algún botón nativo de SAP no realice alguna acción sobre la ventana             
                if (pVal.MenuUID == BTN_SAP_BUSCAR || pVal.MenuUID == BTN_SAP_CREAR && pVal.BeforeAction == true)
                {
                    if (_Application != null)
                    {
                        if (_Application.Forms.ActiveForm.UniqueID == FRM_TRASPASOS)
                            BubbleEvent = false;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en ManuEvent *clsTraspasos* : " + ex.Message);
            }
        }

        #endregion

        #region METODOS

        /// <summary>
        /// Inicializa los eventos de la forma
        /// </summary>
        public void setEventos()
        {
            _Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            _Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
        }

        /// <summary>
        /// Inicializa los componentes de la forma activa
        /// </summary>
        public void inicializarComponentes()
        {
            try
            {
                _oForma = _Application.Forms.Item(formID);
                _oForma.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                _oForma.Freeze(true);

                AddChooseFromListOrigen();
                AddChooseFromListDestino();

                _oForma.DataSources.UserDataSources.Add("dsOrigen", BoDataType.dt_SHORT_TEXT);
                _oContratoOrigen = _oForma.Items.Item(TXT_CONTRATO_ORIGEN).Specific;
                _oContratoOrigen.DataBind.SetBound(true, "", "dsOrigen");
                _oContratoOrigen.ChooseFromListUID = "CFL2";
                _oContratoOrigen.ChooseFromListAlias = "CardCode";

                _oForma.DataSources.UserDataSources.Add("dsDestino", BoDataType.dt_SHORT_TEXT);
                _oContratoDestino = _oForma.Items.Item(TXT_CONTRATO_DESTINO).Specific;
                _oContratoDestino.DataBind.SetBound(true, "", "dsDestino");
                _oContratoDestino.ChooseFromListUID = "CFL3";
                _oContratoDestino.ChooseFromListAlias = "CardCode";

                //Declarar DataTable
                _oForma.DataSources.DataTables.Add(DT_TRASPASO_ORIGEN);
                _oForma.DataSources.DataTables.Add(DT_TRASPASO_DESTINO);

                //Declarar controles
                _oLblNombreOrigen = _oForma.Items.Item(LBL_ORIGEN_NOMBRE_SN).Specific;
                _oLblNombreDestino = _oForma.Items.Item(LBL_DESTINO_NOMBRE_SN).Specific;
                _oPlanOrigen = _oForma.Items.Item(TXT_PLAN_ORIGEN).Specific;
                _oPlanDestino = _oForma.Items.Item(TXT_PLAN_DESTINO).Specific;

                _oForma.DataSources.UserDataSources.Add(TXT_COSTO_ORIGEN, BoDataType.dt_PRICE);
                _oCostoOrigen = _oForma.Items.Item(TXT_COSTO_ORIGEN).Specific;
                _oCostoOrigen.DataBind.SetBound(true, "", TXT_COSTO_ORIGEN);

                _oForma.DataSources.UserDataSources.Add(TXT_COSTO_DESTINO, BoDataType.dt_PRICE);
                _oCostoDestino = _oForma.Items.Item(TXT_COSTO_DESTINO).Specific;
                _oCostoDestino.DataBind.SetBound(true, "", TXT_COSTO_DESTINO);

                _oForma.DataSources.UserDataSources.Add(TXT_SALDO_ORIGEN, BoDataType.dt_PRICE);
                _oSaldoOrigen = _oForma.Items.Item(TXT_SALDO_ORIGEN).Specific;
                _oSaldoOrigen.DataBind.SetBound(true, "", TXT_SALDO_ORIGEN);

                _oForma.DataSources.UserDataSources.Add(TXT_SALDO_DESTINO, BoDataType.dt_PRICE);
                _oSaldoDestino = _oForma.Items.Item(TXT_SALDO_DESTINO).Specific;
                _oSaldoDestino.DataBind.SetBound(true, "", TXT_SALDO_DESTINO);

                _oForma.DataSources.UserDataSources.Add(TXT_MONTO, BoDataType.dt_PRICE);
                _oMontoSeleccionado = _oForma.Items.Item(TXT_MONTO).Specific;
                _oMontoSeleccionado.DataBind.SetBound(true, "", TXT_MONTO);

                _oEmpresaOrigen = _oForma.Items.Item(TXT_EMPRESA_ORIGEN).Specific;
                _oEmpresaDestino = _oForma.Items.Item(TXT_EMPRESA_DESTINO).Specific;

                _oFacturaOrigen = _oForma.Items.Item(TXT_FACTURA_ORIGEN).Specific;
                _oFacturaDestino = _oForma.Items.Item(TXT_FACTURA_DESTINO).Specific;

                SAPbouiCOM.LinkedButton oLinkButton = null;
                SAPbouiCOM.Item oItem = null;

                oItem = _oForma.Items.Add("Origen", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                oItem.LinkTo = TXT_CONTRATO_ORIGEN;
                oItem.Top = _oForma.Items.Item(TXT_CONTRATO_ORIGEN).Top - 1;
                oItem.Left = _oForma.Items.Item(TXT_CONTRATO_ORIGEN).Left - 20;
                oLinkButton = oItem.Specific;
                oLinkButton.LinkedObjectType = LINK_DM_SOCIO_NEGOCIO;

                oItem = _oForma.Items.Add("Destino", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                oItem.LinkTo = TXT_CONTRATO_DESTINO;
                oItem.Top = _oForma.Items.Item(TXT_CONTRATO_DESTINO).Top - 1;
                oItem.Left = _oForma.Items.Item(TXT_CONTRATO_DESTINO).Left - 20;
                oLinkButton = oItem.Specific;
                oLinkButton.LinkedObjectType = LINK_DM_SOCIO_NEGOCIO;

                SAPbouiCOM.Button oButton = null;

                oItem = _oForma.Items.Add(BTN_ORIGEN, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = _oForma.Items.Item(TXT_CONTRATO_ORIGEN).Left + 144;
                oItem.Top = _oForma.Items.Item(TXT_CONTRATO_ORIGEN).Top - 2;
                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Image;
                oItem.Width = 20;
                oItem.Height = 20;
                oButton.Image = Directory.GetCurrentDirectory() + @"\CFL.BMP";
                oButton.ChooseFromListUID = "CFL2";

                oItem = _oForma.Items.Add(BTN_DESTINO, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = _oForma.Items.Item(TXT_CONTRATO_DESTINO).Left + 144;
                oItem.Top = _oForma.Items.Item(TXT_CONTRATO_DESTINO).Top - 2;
                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Image;
                oItem.Width = 20;
                oItem.Height = 20;
                oButton.Image = Directory.GetCurrentDirectory() + @"\CFL.BMP";
                oButton.ChooseFromListUID = "CFL3";
            }
            catch (Exception ex)
            {
                _Application.MessageBox("Error al inicializar: " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        /// <summary>
        /// Se crea el método de ChooseFromList
        /// </summary>
        public void AddChooseFromListOrigen()
        {
            try
            {

                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                SAPbouiCOM.Conditions oCons = null;
                SAPbouiCOM.Condition oCon = null;

                oCFLs = _oForma.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                //  Adding 2 CFL, one for the button and one for the edit text.
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "2";
                oCFLCreationParams.UniqueID = "CFL2";

                oCFL = oCFLs.Add(oCFLCreationParams);   
                oCons = oCFL.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "C";
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCon = oCons.Add();
                oCon.Alias = "GroupCode";
                if (_Application.Company.Name == "TAMPICO PROGRAMA DE APOYO")
                {
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "102";
                }
                else
                {
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "100";
                }
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCon = oCons.Add();
                oCon.Alias = "balance";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCon.CondVal = "0";

                oCFL.SetConditions(oCons);


            }
            catch (Exception ex)
            {
                throw new Exception("Error al crear ChooseFromListOrigen *AddChooseFromListOrigen* : " + ex.Message);
            }
        }

        /// <summary>
        /// Se crea el método de ChooseFromList
        /// </summary>
        public void AddChooseFromListDestino()
        {
            try
            {
                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                SAPbouiCOM.Conditions oCons = null;
                SAPbouiCOM.Condition oCon = null;

                oCFLs = _oForma.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                //  Adding 2 CFL, one for the button and one for the edit text.
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "2";
                oCFLCreationParams.UniqueID = "CFL3";

                oCFL = oCFLs.Add(oCFLCreationParams);
                oCons = oCFL.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "C";
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCon = oCons.Add();
                oCon.Alias = "GroupCode";
                if (_Application.Company.Name == "TAMPICO PROGRAMA DE APOYO")
                {
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "102";
                }
                else
                {
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "100";
                }
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCon = oCons.Add();
                oCon.Alias = "balance";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCon.CondVal = "0";

                oCFL.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al crear ChooseFromListDestino *AddChooseFromListDestino* : " + ex.Message);
            }
        }

        /// <summary>
        /// Obtiene los pagos del contrato
        /// </summary>
        /// <param name="contrato">Numero de contrato a consultar</param>
        private void CargarGridOrigen(string contrato)
        {
            try
            {
                _oForma.Freeze(true);

                _oForma.DataSources.DataTables.Item(DT_TRASPASO_ORIGEN).ExecuteQuery(@"SELECT  '' AS Sel,
                                                                                                T0.DocEntry AS DocEntry,
                                                                                                T0.CashSum AS Monto,
                                                                                                T0.U_NumeroReciboEcobro AS NumRecibo,
                                                                                                T0.JrnlMemo AS Referencia,
                                                                                                T0.DocDate AS Fecha
                                                                                        FROM    dbo.ORCT T0
                                                                                                INNER JOIN dbo.RCT2 T1 ON T1.DocNum = T0.DocEntry
                                                                                                INNER JOIN dbo.OCRD T2 ON T2.CardCode = T0.CardCode
                                                                                                LEFT JOIN dbo.OINV T3 ON T3.CardCode = T0.CardCode
                                                                                                                            AND T1.baseAbs = T3.DocEntry
                                                                                                                            AND T3.DocStatus = 'O'
                                                                                        WHERE   T0.CardCode = '" + contrato + "' AND T0.Canceled = 'N'" +
                                                                                        "ORDER BY T0.DocEntry ASC");

                _oGridTraspasoOrigen = _oForma.Items.Item(GRD_TRASPASO_ORIGEN).Specific;
                _oGridTraspasoOrigen.DataTable = _oForma.DataSources.DataTables.Item(DT_TRASPASO_ORIGEN);
                FormatoGridOrigen(_oGridTraspasoOrigen);


            }
            catch (Exception ex)
            {
                throw new Exception("Error al cargar grid origen *CargarGridOrigen* : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        /// <summary>
        /// Se le da el formato al grid para visualizarlo
        /// </summary>
        /// <param name="_oGridTraspasoOrigen">Objeto del grid</param>
        private void FormatoGridOrigen(Grid grid)
        {
            try
            {
                _oForma.Freeze(true);
                if (!grid.DataTable.IsEmpty)
                {
                    grid.RowHeaders.TitleObject.Caption = "#";
                    for (int noLinea = 0; noLinea < grid.Rows.Count; noLinea++)
                    {
                        grid.RowHeaders.SetText(noLinea, (noLinea + 1).ToString());
                    }

                    if (!grid.DataTable.IsEmpty)
                    {
                        grid.Columns.Item("Sel").TitleObject.Caption = "Sel.";
                        grid.Columns.Item("Sel").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    }

                    grid.Columns.Item("DocEntry").Editable = false;
                    grid.Columns.Item("DocEntry").TitleObject.Caption = "Folio SAP";
                    SAPbouiCOM.EditTextColumn oColTransferStock = grid.Columns.Item("DocEntry") as SAPbouiCOM.EditTextColumn;
                    oColTransferStock.LinkedObjectType = LINK_PAGO_EFECTUADO;

                    grid.Columns.Item("Monto").Editable = false;
                    grid.Columns.Item("Monto").TitleObject.Caption = "Monto";

                    //Agrega linea con la suma de la columna.
                    SAPbouiCOM.EditTextColumn montosPagos = (SAPbouiCOM.EditTextColumn)grid.Columns.Item("Monto");
                    montosPagos.ColumnSetting.SumType = BoColumnSumType.bst_Auto;


                    grid.Columns.Item("NumRecibo").Editable = false;
                    grid.Columns.Item("NumRecibo").TitleObject.Caption = "Recibo E_Cobro";

                    grid.Columns.Item("Referencia").Editable = false;
                    grid.Columns.Item("Referencia").TitleObject.Caption = "Referencia";

                    grid.Columns.Item("Fecha").Editable = false;
                    grid.Columns.Item("Fecha").TitleObject.Caption = "Fecha";

                    grid.AutoResizeColumns();
                }
                else
                {
                    _Application.MessageBox("No se encontraron pagos para este contrato");
                    _oGridTraspasoOrigen.DataTable.Clear();
                    _oForma.Update();
                }


            }
            catch (Exception ex)
            {
                throw new Exception("Error al dar formato al grid origen *FormatoGridOrigen* : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        /// <summary>
        /// Obtiene los pagos del contrato
        /// </summary>
        /// <param name="contrato">Numero de contrato a consultar</param>
        private void CargarGridDestino(string contrato)
        {
            try
            {
                _oForma.Freeze(true);

                _oForma.DataSources.DataTables.Item(DT_TRASPASO_DESTINO).ExecuteQuery(@"SELECT  T0.DocEntry AS DocEntry,
                                                                                                T0.CashSum AS Monto,
                                                                                                T0.U_NumeroReciboEcobro AS NumRecibo,
                                                                                                T0.JrnlMemo AS Referencia,
                                                                                                T0.DocDate AS Fecha
                                                                                        FROM    dbo.ORCT T0
                                                                                                INNER JOIN dbo.RCT2 T1 ON T1.DocNum = T0.DocEntry
                                                                                                INNER JOIN dbo.OCRD T2 ON T2.CardCode = T0.CardCode
                                                                                                INNER JOIN dbo.OINV T3 ON T3.CardCode = T0.CardCode
                                                                                                                            AND T1.baseAbs = T3.DocEntry
                                                                                                                            AND T3.DocStatus = 'O'
                                                                                        WHERE   T0.CardCode = '" + contrato + "' AND T0.Canceled = 'N'" +
                                                                                        "ORDER BY T0.DocEntry ASC");

                _oGridTraspasoDestino = _oForma.Items.Item(GRD_TRASPASO_DESTINO).Specific;
                _oGridTraspasoDestino.DataTable = _oForma.DataSources.DataTables.Item(DT_TRASPASO_DESTINO);
                FormatoGridDestino(_oGridTraspasoDestino);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al cargar grid origen *CargarGridDestino* : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        /// <summary>
        /// Se le da el formato al grid para visualizarlo
        /// </summary>
        /// <param name="_oGridTraspasoOrigen">Objeto del grid</param>
        private void FormatoGridDestino(Grid grid)
        {
            try
            {
                if (!grid.DataTable.IsEmpty)
                {
                    grid.RowHeaders.TitleObject.Caption = "#";
                    for (int noLinea = 0; noLinea < grid.Rows.Count; noLinea++)
                    {
                        grid.RowHeaders.SetText(noLinea, (noLinea + 1).ToString());
                    }

                    grid.Columns.Item("DocEntry").Editable = false;
                    grid.Columns.Item("DocEntry").TitleObject.Caption = "Folio SAP";
                    SAPbouiCOM.EditTextColumn oColTransferStock = grid.Columns.Item("DocEntry") as SAPbouiCOM.EditTextColumn;
                    oColTransferStock.LinkedObjectType = LINK_PAGO_EFECTUADO;

                    grid.Columns.Item("Monto").Editable = false;
                    grid.Columns.Item("Monto").TitleObject.Caption = "Monto";

                    //Agrega linea con la suma de la columna.
                    SAPbouiCOM.EditTextColumn montosPagos = (SAPbouiCOM.EditTextColumn)grid.Columns.Item("Monto");
                    montosPagos.ColumnSetting.SumType = BoColumnSumType.bst_Auto;

                    grid.Columns.Item("NumRecibo").Editable = false;
                    grid.Columns.Item("NumRecibo").TitleObject.Caption = "Recibo E_Cobro";

                    grid.Columns.Item("Referencia").Editable = false;
                    grid.Columns.Item("Referencia").TitleObject.Caption = "Referencia";

                    grid.Columns.Item("Fecha").Editable = false;
                    grid.Columns.Item("Fecha").TitleObject.Caption = "Fecha";
                    grid.AutoResizeColumns();

                }
                else
                {
                    _Application.MessageBox("No se encontraron pagos para este contrato");
                    _oGridTraspasoDestino.DataTable.Clear();
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Error al dar formato al grid origen *FormatoGridDestino* : " + ex.Message);
            }
        }

        /// <summary>
        /// Se seleccionan o no todos lo pagos
        /// </summary>
        /// <param name="Seleccionar"></param>
        private void SeleccionarTodos(string Seleccionar)
        {
            try
            {
                _oForma.Freeze(true);
                if (_oGridTraspasoOrigen != null)
                {
                    if (_oGridTraspasoOrigen.Rows.Count != 0)
                    {
                        _oMontoSeleccionado.Value = "";
                        for (int jj = 0; jj <= _oGridTraspasoOrigen.DataTable.Rows.Count - 1; jj++)
                        {
                            _oGridTraspasoOrigen.DataTable.SetValue("Sel", jj, Seleccionar);
                            string seleccionado = _oGridTraspasoOrigen.DataTable.GetValue(COLUMNA_SELECCIONAR, jj);
                            if (seleccionado == "Y")
                            {
                                _oMontoSeleccionado.Value = Convert.ToString(Math.Round(Convert.ToDouble(_oMontoSeleccionado.Value.ToString()) + Convert.ToDouble(_oGridTraspasoOrigen.DataTable.GetValue("Monto", jj)), 2));
                            }
                            else if (seleccionado == "N")
                            {
                                _oMontoSeleccionado.Value = "";
                            }
                        }
                    }
                    else
                    {
                        _Application.MessageBox("No se encontraron registros");
                    }
                }
                else
                {
                    _Application.MessageBox("No se encontraron registros");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al seleccionar todos SI/NO *SeleccionarTodos* : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        /// <summary>
        /// Genera los traspasos entre saldos
        /// </summary>
        private void GenerarTraspasos()
        {
            try
            {
                if (Convert.ToDouble(_oMontoSeleccionado.Value) != 0)
                {
                    if (!string.IsNullOrEmpty(_oContratoDestino.Value.ToString()) && !string.IsNullOrEmpty(_oPlanDestino.Value.ToString()) && Convert.ToDouble(_oCostoDestino.Value) > 0 && !string.IsNullOrEmpty(_oEmpresaDestino.Value))
                    {
                        if (!string.IsNullOrEmpty(_oFacturaOrigen.Value.ToString()) && !string.IsNullOrEmpty(_oFacturaDestino.Value.ToString()))
                        {
                            if (_oContratoOrigen.Value.ToString() != _oContratoDestino.Value.ToString())
                            {
                                double montoRestante = Convert.ToDouble(_oSaldoDestino.Value) - Convert.ToDouble(_oMontoSeleccionado.Value);

                                if (montoRestante >= 0)
                                {
                                    if (_Application.MessageBox("Realizar traspaso: " + Environment.NewLine +
                                                            _oContratoOrigen.Value.ToString() + " --> " + _oContratoDestino.Value.ToString() + Environment.NewLine +
                                                            "Monto: " + Convert.ToDouble(_oMontoSeleccionado.Value).ToString("0.00") + Environment.NewLine +
                                                             "¿Desea continuar?", 2, "Si", "No") == 1)
                                    {
                                        _Application.StatusBar.SetText("Realizando traspasos por favor espere...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);

                                        if (_oEmpresaOrigen.Value.ToString() == "COOPERATIVA" && _oEmpresaDestino.Value.ToString() == "COOPERATIVA")
                                        {
                                            string msgUnoAUno = GenerarPagosUnoAUno();
                                            if (string.IsNullOrEmpty(msgUnoAUno))
                                            {
                                                _Application.MessageBox("Traspaso realizado con éxito");
                                                _Application.StatusBar.SetText("Traspaso realizado con éxito...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                            }
                                            else
                                            {
                                                _Application.MessageBox("No se completo el traspaso: " + msgUnoAUno);
                                            }

                                        }
                                        else
                                        {
                                           string msgMushosUno =  GenerarPagosMuchosAUno();
                                           if (string.IsNullOrEmpty(msgMushosUno))
                                           {
                                               _Application.MessageBox("Traspaso realizado con éxito");
                                               _Application.StatusBar.SetText("Traspaso realizado con éxito...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                           }
                                           else
                                           {
                                               _Application.MessageBox("No se completo el traspaso: " + msgMushosUno);
                                           }
                                        }
                                    }
                                }
                                else
                                {
                                    _Application.MessageBox("El monto sobrepasa el saldo del contrato destino");
                                }
                            }
                            else
                            {
                                _Application.MessageBox("Los contratos deben ser diferentes");
                            }
                        }
                        else
                        {
                            _Application.MessageBox("No existen facturas ligadas a estos contratos.");
                        }
                    }
                    else
                    {
                        _Application.MessageBox("El contrato no contiene información: " + _oContratoDestino.Value.ToString());
                    }
                }
                else
                {
                    _Application.MessageBox("No se ha seleccionado un pago");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al generar Traspasos *GenerarTraspasos* : " + ex.Message);
            }
        }

        /// <summary>
        /// Obtiene los datos del contrato Origen 
        /// </summary>
        private void ObtenerDatosOrigen()
        {
            try
            {
                Extensor.DatosTraspasos datoTraspasos = new Extensor.DatosTraspasos();
                datoTraspasos = Extensor.ObtenerDatosTraspasos(_oContratoOrigen.Value.ToString(), _Company);

                _oLblNombreOrigen.Caption = datoTraspasos.nombreSN;
                _oPlanOrigen.Value = datoTraspasos.plan;
                _oCostoOrigen.Value = datoTraspasos.costoPlan.ToString();
                _oSaldoOrigen.Value = datoTraspasos.saldo.ToString();
                _oEmpresaOrigen.Value = datoTraspasos.empresa.ToString();
                _oFacturaOrigen.Value = datoTraspasos.DocEntryFactura.ToString();

                CargarGridOrigen(_oContratoOrigen.Value.ToString());
                _oMontoSeleccionado.Value = "";
            }
            catch (Exception ex)
            {

                throw new Exception("Error al obtener los datos de origen *ObtenerDatosOrigen* : " + ex.Message);
            }
        }

        /// <summary>
        /// Obtiene los datos del contrato destino
        /// </summary>
        private void ObtenerDatosDestino()
        {
            try
            {
                Extensor.DatosTraspasos datoTraspasos = new Extensor.DatosTraspasos();
                datoTraspasos = Extensor.ObtenerDatosTraspasos(_oContratoDestino.Value.ToString(), _Company);


                _oLblNombreDestino.Caption = datoTraspasos.nombreSN;
                _oPlanDestino.Value = datoTraspasos.plan;
                _oCostoDestino.Value = datoTraspasos.costoPlan.ToString();
                _oSaldoDestino.Value = datoTraspasos.saldo.ToString();
                _oEmpresaDestino.Value = datoTraspasos.empresa.ToString();
                _oFacturaDestino.Value = datoTraspasos.DocEntryFactura.ToString();

                CargarGridDestino(_oContratoDestino.Value.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener los datos destino *ObtenerDatosDestino* : " + ex.Message);
            }
        }

        /// <summary>
        /// Genera los pagos cuando son de apoyo a cooperativa
        /// </summary>
        private string GenerarPagosUnoAUno()
        {
            SAPbobsCOM.Payments _oPagoCancel = null;
            SAPbobsCOM.Payments _oPagoTraspaso = null;
            string msgError = null;

            try
            {               
                _oForma.Freeze(true);

                int i = 0;
                string seleccionado = null;
                int docentryPago = 0;
                double monto = 0;
                
                string usuarioFirmado = ObtenerUsuarioFirmado(_Company.UserName.ToString());
                string folioE_cobro = null;
                string fechaPago = null;

                _Company.StartTransaction();

                while (_oGridTraspasoOrigen.Rows.Count - 1 >= i)
                {
                    seleccionado = _oGridTraspasoOrigen.DataTable.GetValue("Sel", i).ToString();

                    if (seleccionado == "Y")
                    {
                        _oPagoCancel = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                        _oPagoTraspaso = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);

                        docentryPago = _oGridTraspasoOrigen.DataTable.GetValue("DocEntry", i);
                        monto = _oGridTraspasoOrigen.DataTable.GetValue("Monto", i);
                        folioE_cobro = _oGridTraspasoOrigen.DataTable.GetValue("NumRecibo", i);
                        fechaPago = _oGridTraspasoOrigen.DataTable.GetValue("Fecha", i).ToString();

                        _oPagoCancel.GetByKey(docentryPago);

                        if (_oPagoCancel.Cancel() != 0)
                        {
                            msgError = _Company.GetLastErrorDescription();
                        }
                        else
                        {
                            

                            _oPagoTraspaso.DocDate = DateTime.Now;
                            _oPagoTraspaso.Remarks = usuarioFirmado;
                            _oPagoTraspaso.JournalRemarks = "Traspaso";
                            _oPagoTraspaso.DocType = SAPbobsCOM.BoRcptTypes.rCustomer;
                            _oPagoTraspaso.CardCode = _oContratoDestino.Value;
                            if (_oLblNombreDestino.Caption.ToString().Length > 30)
                            {
                                _oPagoTraspaso.UserFields.Fields.Item("U_BeneficiarioRecibo").Value = _oLblNombreDestino.Caption.ToString().Substring(0, 30);
                            }
                            else
                            {
                                _oPagoTraspaso.UserFields.Fields.Item("U_BeneficiarioRecibo").Value = _oLblNombreDestino.Caption.ToString();
                            }
                            _oPagoTraspaso.UserFields.Fields.Item("U_Traspaso").Value = "Documento origen: " + docentryPago + " ( " + folioE_cobro + " ) " + "$ " + _oGridTraspasoOrigen.DataTable.GetValue("Monto", i) + " - " + fechaPago.Substring(0, 10);
                            
                            _oPagoTraspaso.Invoices.DocEntry = Convert.ToInt32(_oFacturaDestino.Value);
                            _oPagoTraspaso.CashAccount = Addon.listaCtasSAP.First(x => x.Documento == "TRASPASO").cuenta; //Extensor.Configuracion.CUENTASTRASPASOS.CuentaApoyoCooperativa;
                            _oPagoTraspaso.CashSum = monto;
                            _oPagoTraspaso.Invoices.SumApplied = monto;
                            _oPagoTraspaso.UserFields.Fields.Item("U_Es_PagoDirecto").Value = "NO";

                            if (_oPagoTraspaso.Add() != 0)
                            {
                                msgError = _Company.GetLastErrorDescription();
                            }
                        }
                    }
                    i += 1;
                }
            }
            catch (Exception ex)
            {
                msgError = "Error al generar los pagos de apoyo a Cooperativa *GenerarPagosUnoAUno* : " + ex.Message;
            }
            finally
            {
                try
                {
                    if (string.IsNullOrEmpty(msgError))
                    {
                        _Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        ActualizarFideicomizo();
                        ObtenerDatosDestino();
                        ObtenerDatosOrigen();
                    }
                    else
                    {
                        _Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }

                    if (_oPagoCancel != null)
                    {
                        GC.SuppressFinalize(_oPagoCancel);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oPagoCancel);
                        _oPagoCancel = null;
                    }

                    if (_oPagoTraspaso != null)
                    {
                        GC.SuppressFinalize(_oPagoTraspaso);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oPagoTraspaso);
                        _oPagoTraspaso = null;
                    }
                    GC.Collect();
                    ClearMemory();
                    _oForma.Freeze(false);
                }
                catch (Exception)
                {
                }             
            }
            return msgError;
        }

        /// <summary>
        /// Genera los pagos de muchos a uno de Apoyo a Apoyo, Cooperativa a Cooperativa y Cooperativa a Apoyo
        /// </summary>
        private string GenerarPagosMuchosAUno()
        {
            SAPbobsCOM.Payments _oPagoCancel = null;
            SAPbobsCOM.Payments _oPagoTraspaso = null;
            string msgError = null;
            try
            {
                _oForma.Freeze(true);

                int i = 0;
                string seleccionado = null;
                int docentryPago = 0;
                double monto = 0;                
                string usuarioFirmado = ObtenerUsuarioFirmado(_Company.UserName.ToString());
                string folioE_cobro = null;
                string fechaPago = null;

                _Company.StartTransaction();

                List<DatosTraspaso> lDatosTraspaso = new List<DatosTraspaso>();
                DatosTraspaso itemDatos = new DatosTraspaso();
                               
                lDatosTraspaso.Clear();

                while (_oGridTraspasoOrigen.Rows.Count - 1 >= i)
                {
                    seleccionado = _oGridTraspasoOrigen.DataTable.GetValue("Sel", i).ToString();

                    if (seleccionado == "Y")
                    {
                        itemDatos = new DatosTraspaso();
                        _oPagoCancel = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);

                        docentryPago = _oGridTraspasoOrigen.DataTable.GetValue("DocEntry", i);                        
                        folioE_cobro = _oGridTraspasoOrigen.DataTable.GetValue("NumRecibo", i);
                        fechaPago = _oGridTraspasoOrigen.DataTable.GetValue("Fecha", i).ToString();

                        _oPagoCancel.GetByKey(docentryPago);

                        if (_oPagoCancel.Cancel() != 0)
                        {
                            msgError = _Company.GetLastErrorDescription();
                           // _Application.StatusBar.SetText("Error al cancelar el pago: " + docentryPago + " : " + msgError, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                        else
                        {
                            monto += _oGridTraspasoOrigen.DataTable.GetValue("Monto", i);
                            //_Application.StatusBar.SetText("Pago cancelado correctamente: " + docentryPago, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            itemDatos.DocEntryPago_FolioE_Cobro = "Documento origen: " + docentryPago + " ( " + folioE_cobro + " ) " + "$ " + _oGridTraspasoOrigen.DataTable.GetValue("Monto", i) + " - " + fechaPago.Substring(0, 10);
                            lDatosTraspaso.Add(itemDatos);
                        }
                    }
                    i += 1;
                }

                _oPagoTraspaso = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);             

                _oPagoTraspaso.DocDate = DateTime.Now;
                _oPagoTraspaso.Remarks = usuarioFirmado;
                _oPagoTraspaso.JournalRemarks = "Traspaso";
                _oPagoTraspaso.DocType = SAPbobsCOM.BoRcptTypes.rCustomer;
                _oPagoTraspaso.CardCode = _oContratoDestino.Value;
                if (_oLblNombreDestino.Caption.ToString().Length > 30)
                {
                    _oPagoTraspaso.UserFields.Fields.Item("U_BeneficiarioRecibo").Value = _oLblNombreDestino.Caption.ToString().Substring(0, 30);
                }
                else
                {
                    _oPagoTraspaso.UserFields.Fields.Item("U_BeneficiarioRecibo").Value = _oLblNombreDestino.Caption.ToString();
                }
                _oPagoTraspaso.UserFields.Fields.Item("U_Traspaso").Value = string.Join(Environment.NewLine, lDatosTraspaso.Select(c => c.DocEntryPago_FolioE_Cobro));
                _oPagoTraspaso.Invoices.DocEntry = Convert.ToInt32(_oFacturaDestino.Value);
                _oPagoTraspaso.CashAccount = Addon.listaCtasSAP.First(x => x.Documento == "TRASPASO").cuenta; //Extensor.Configuracion.CUENTASTRASPASOS.CuentaApoyoCooperativa; 
                _oPagoTraspaso.CashSum = monto;
                _oPagoTraspaso.Invoices.SumApplied = monto;
                _oPagoTraspaso.UserFields.Fields.Item("U_Es_PagoDirecto").Value = "NO";

                if (_oPagoTraspaso.Add() != 0)
                {
                    msgError = _Company.GetLastErrorDescription();
                    //_Application.StatusBar.SetText("Ocurrió un error al crear el pago del traspaso: " + docentryPago, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }            
          
            
            }
            catch (Exception ex)
            {
                msgError = "Error al generar los pagos de muchos a uno *GenerarPagosMuchosAUno* : " + ex.Message;
            }
            finally
            {
                try
                {
                    if(string.IsNullOrEmpty(msgError))
                    {
                        _Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        ActualizarFideicomizo();
                        ObtenerDatosDestino();
                        ObtenerDatosOrigen();
                    }
                    else
                    {
                        _Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }

                    if(_oPagoCancel != null)
                    {
                        GC.SuppressFinalize(_oPagoCancel);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oPagoCancel);
                        _oPagoCancel = null;
                    }

                    if(_oPagoTraspaso != null )
                    {
                        GC.SuppressFinalize(_oPagoTraspaso);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oPagoTraspaso);
                        _oPagoTraspaso = null;
                    }
                    GC.Collect();
                    ClearMemory();
                    _oForma.Freeze(false);
                }
                catch (Exception)
                {
                }
            }
            return msgError;
        }

        /// <summary>
        /// Actualiza los datos de la tabla de calculo de comisiones con lo actual del fideicomiso
        /// </summary>
        private void ActualizarFideicomizo()
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  ( U_Inv_Inicial + U_Recomendado + U_Asis_Social + U_BonoAsistente + U_Lider + U_Supervisor + U_Coordinador + U_Coordinador2 + U_Gerente ) AS MontoOrigen
                                FROM    dbo.[@CALCULO_COMISIONES]
                                WHERE   U_Contrato = '" + _oContratoOrigen.Value.ToString() + "' ");
                double montoOrigen = _oRec.Fields.Item("MontoOrigen").Value;

                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  ( U_Inv_Inicial + U_Recomendado + U_Asis_Social + U_BonoAsistente + U_Lider + U_Supervisor + U_Coordinador + U_Coordinador2 + U_Gerente ) AS MontoDestino
                                FROM    dbo.[@CALCULO_COMISIONES]
                                WHERE   U_Contrato = '" + _oContratoDestino.Value.ToString() + "' ");
                double montoDestino = _oRec.Fields.Item("MontoDestino").Value;

                double montoFideicomisoOrigen = Extensor.ObtenerSaldoDeCuenta(_oContratoOrigen.Value.ToString(), _Company);
                double montoFideicomisoDestino = Extensor.ObtenerSaldoDeCuenta(_oContratoDestino.Value.ToString(), _Company);

                double totalOrigen = montoFideicomisoOrigen - montoOrigen;
                double totalDestino = montoFideicomisoDestino - montoDestino;

                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery("UPDATE dbo.[@CALCULO_COMISIONES] SET U_Fideicomiso = '" + totalOrigen + "' WHERE U_Contrato = '" + _oContratoOrigen.Value.ToString() + "'");

                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery("UPDATE dbo.[@CALCULO_COMISIONES] SET U_Fideicomiso = '" + totalDestino + "' WHERE U_Contrato = '" + _oContratoDestino.Value.ToString() + "'");
            }
            catch (Exception ex)
            {
                throw new Exception("Error al actualizar el Fideicomiso *ActualizarFideicomizo* : " + ex.Message);
            }
            finally
            {
                if (_oRec != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_oRec);
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
                if (_oRec != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_oRec);
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

        #region CLASE TRASPASOS

        public class DatosTraspaso
        {
            public string DocEntryPago_FolioE_Cobro { get; set; }
        }

        #endregion
    }
}
