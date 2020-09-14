using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddonFNR.BL
{
    class frmPreContratos : ComportaForm
    {
        #region CONSTANTES

        private const string FRM_PRECONTRATOS = "frmPreContratos";
        private const string FRM_DATOS_MAESTROS_SOCIOS = "-134";
        private const string TXTFECHA_INICIAL = "txtFI";
        private const string TXTFECHA_FINAL = "txtFF";
        private const string TXT_OFICINA = "txtOfici";
        private const string GRDPRECONTRATO = "grdPreCon";
        private const string BTNBUSCAR = "btnBuscar";
        private const string BTNCOLLAP = "btnCollap";
        private const string BTNEXPAN = "btnExpan";
        private const string BTNCREAR = "btnCrear";
        private const string BTN_SELECCION_AUTOMATICA = "btnSelAut";
        private const string BTN_SAP_BUSCAR = "1281";
        private const string BTN_SAP_CREAR = "1282";
        private const string DT_PRECONTRATOS = "dtPreContratos";
        private const string LINK_TRANSFERENCIA_STOCK = "67";
        private const string LINK_ARTICULO = "4";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForma = null;
        private static bool _oPrecontratos = false;
        private SAPbobsCOM.Recordset _oRec = null;

        private SAPbouiCOM.Grid _oGridPrecontratos = null;
        private SAPbouiCOM.EditText _oTxtFechaInicial = null;
        private SAPbouiCOM.EditText _oTxtFechaFinal = null;
        private SAPbouiCOM.EditText _oTxtOficina = null;
        private static List<Datos> lDatos = new List<Datos>();
        private static Datos itemDatos = new Datos();
        private string EstatusCrear = null;
        private int ContadorSeleccionados = 0;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de autorización de ofertas de compra
        /// </summary>
        /// <param name="_Application">Este es el objeto raíz de la API de interfaz de usuario
        ///                             lo que refleja la cual aplicación SAP Business One en el que se realiza 
        ///                             la conexión</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        public frmPreContratos(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oPrecontratos == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                showForm(formID);
                inicializarComponentes();
                setEventos();
                _oPrecontratos = true;
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
                if (_oPrecontratos != false && pVal.FormType != FormTypeMenu && formID == FormUID)
                {
                    eventos(FormUID, ref pVal, out bubbleEvent);
                }

                if (pVal.FormType.ToString() == FRM_DATOS_MAESTROS_SOCIOS && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
                {
                    BuscarDatos();
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
                if (pVal.FormUID == formID && pVal.BeforeAction == false)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
                    {
                        _Application.ItemEvent -= new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                        _Application.MenuEvent -= new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                        Dispose();
                        application = null;
                        company = null;
                        _oPrecontratos = false;
                        Addon.typeList.RemoveAll(p => p._forma == formID);
                        return;
                    }                
                }

                if (pVal.FormUID == formID && pVal.BeforeAction == true)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if (pVal.ItemUID == BTNBUSCAR)
                        {
                            BuscarDatos();
                            ContadorSeleccionados = 0;
                        }

                        if (pVal.ItemUID == BTNCOLLAP)
                        {
                            if (_oGridPrecontratos != null)
                            {
                                if (!_oGridPrecontratos.DataTable.IsEmpty)
                                {
                                    _oGridPrecontratos.Rows.CollapseAll();
                                }
                            }
                        }

                        if (pVal.ItemUID == BTNEXPAN)
                        {
                            if (_oGridPrecontratos != null)
                            {
                                if (!_oGridPrecontratos.DataTable.IsEmpty)
                                {
                                    _oGridPrecontratos.Rows.ExpandAll();
                                }
                            }
                        }

                        if (pVal.ItemUID == BTNCREAR)
                        {
                            if (_oGridPrecontratos != null)
                            {
                                if (!_oGridPrecontratos.DataTable.IsEmpty)
                                {
                                    if (ContadorSeleccionados != 0)
                                    {
                                        CrearPreContratos();
                                        _oForma.Close();
                                        bubbleEvent = false;
                                    }
                                    else
                                    {
                                        _Application.MessageBox("No se ha seleccionado ningún Pre-contrato");
                                    }
                                }
                            }
                        }

                        if (pVal.ItemUID == BTN_SELECCION_AUTOMATICA)
                        {
                            if (_oGridPrecontratos != null)
                            {
                                if (!_oGridPrecontratos.DataTable.IsEmpty)
                                {
                                    SelecionAutomatica();
                                }
                            }
                        }
                    }

                    if (pVal.ItemUID == GRDPRECONTRATO && pVal.ColUID == "RowsHeader" && pVal.EventType == BoEventTypes.et_DOUBLE_CLICK)
                    {
                        bubbleEvent = false;
                        return;
                    }              

                    if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == GRDPRECONTRATO && pVal.ColUID == "Crear")
                    {
                        int index = 0;
                        index = _oGridPrecontratos.GetDataTableRowIndex(pVal.Row);
                        _oGridPrecontratos = _oForma.Items.Item(GRDPRECONTRATO).Specific;
                        EstatusCrear = _oGridPrecontratos.DataTable.GetValue("Crear", index).ToString();

                        if (EstatusCrear == "Y")
                        {
                            ContadorSeleccionados += 1;
                        }
                        else
                        {
                            ContadorSeleccionados -= 1;
                        }

                        if (ContadorSeleccionados > 10)
                        {
                            _oGridPrecontratos.DataTable.SetValue("Crear", index, "N");
                            _Application.MessageBox("Solo se pueden seleccionar hasta 10 Pre-contratos");
                            ContadorSeleccionados -= 1;
                        }
                    }                   
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en evento *clsPreContratos* : " + ex.Message);
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
                        if (_Application.Forms.ActiveForm.UniqueID == FRM_PRECONTRATOS)
                            BubbleEvent = false;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en ManuEvent *clsPreContratos* : " + ex.Message);
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

                //Tipo de dato fecha inicial
                _oForma.DataSources.UserDataSources.Add(TXTFECHA_INICIAL, BoDataType.dt_DATE);
                _oTxtFechaInicial = (SAPbouiCOM.EditText)_oForma.Items.Item(TXTFECHA_INICIAL).Specific;
                _oTxtFechaInicial.DataBind.SetBound(true, "", TXTFECHA_INICIAL);

                //Tipo de dato fecha final
                _oForma.DataSources.UserDataSources.Add(TXTFECHA_FINAL, BoDataType.dt_DATE);
                _oTxtFechaFinal = (SAPbouiCOM.EditText)_oForma.Items.Item(TXTFECHA_FINAL).Specific;
                _oTxtFechaFinal.DataBind.SetBound(true, "", TXTFECHA_FINAL);

                //Tipo de dato Oficina
                _oForma.DataSources.UserDataSources.Add(TXT_OFICINA, BoDataType.dt_LONG_TEXT);
                _oTxtOficina = (SAPbouiCOM.EditText)_oForma.Items.Item(TXT_OFICINA).Specific;
                _oTxtOficina.DataBind.SetBound(true, "", TXT_OFICINA);

                //_oTxtOficina.Value = Extensor.ObtenerAlmacenOficina(_Company);
                _oTxtFechaInicial.Active = true;

                //Declarar DataTable
                _oForma.DataSources.DataTables.Add(DT_PRECONTRATOS);
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
        /// Busca la información de acuerdo a los filtros capturados
        /// </summary>
        /// <param name="efiscal"></param>
        /// <param name="codigo"></param>
        /// <param name="descripcion"></param>
        /// <param name="combos"></param>
        public void BuscarDatos()
        {
            try
            {
                _oForma.Freeze(true);

                string qryFechas = null;
                string qryAlmacen = null;
               

                if (!string.IsNullOrEmpty(_oTxtFechaInicial.Value) || !string.IsNullOrEmpty(_oTxtFechaFinal.Value) ||
                    !string.IsNullOrEmpty(_oTxtOficina.Value))
                {
                    if (!_oTxtFechaInicial.Value.Equals("") && !_oTxtFechaFinal.Value.Equals(""))
                    {
                        string fechaDesde = _oTxtFechaInicial.Value.Substring(0, 4) + "-" + _oTxtFechaInicial.Value.Substring(4, 2) + "-" +
                           _oTxtFechaInicial.Value.Substring(6, 2);

                        string fechaHasta = _oTxtFechaFinal.Value.Substring(0, 4) + "-" + _oTxtFechaFinal.Value.Substring(4, 2) + "-" +
                            _oTxtFechaFinal.Value.Substring(6, 2);

                        DateTime fInicial = Convert.ToDateTime(fechaDesde);
                        DateTime fFinal = Convert.ToDateTime(fechaHasta);

                        if (fInicial <= fFinal)
                        {
                            qryFechas = " AND T1.DocDate BETWEEN '" + _oTxtFechaInicial.Value + "' AND '" + _oTxtFechaFinal.Value + "'";
                        }
                        else
                        {
                            _Application.MessageBox("La fecha inicial es mayor a la final");
                            return;
                        }
                    }
                    else if (!_oTxtFechaInicial.Value.Equals(""))
                    {
                        qryFechas = " AND T1.DocDate = '" + _oTxtFechaInicial.Value + "'";
                    }
                    else if (!_oTxtFechaFinal.Value.Equals(""))
                    {
                        qryFechas = " AND T1.DocDate = '" + _oTxtFechaFinal.Value + "'";
                    }

                    if (!_oTxtOficina.Value.Equals(""))
                    {
                        if (_oTxtOficina.Value == "GDLV97")
                        {
                            qryAlmacen = " AND T1.Filler = 'GDLCONS'";
                        }
                        else
                        {
                            qryAlmacen = " AND T1.Filler = '" + _oTxtOficina.Value + "'";
                        }
                    }                             

                    _oForma.DataSources.DataTables.Item(DT_PRECONTRATOS)
                      .ExecuteQuery(@"SELECT  T1.DocEntry AS DocEntry,
                                                'Crear' AS Crear,
                                                T4.USER_CODE AS CodigoUsuario,
                                                T1.DocNum AS DocNum,
                                                T0.U_Serie AS Serie,
                                                T0.U_CodPromotor AS CodigoPromotor,
                                                T0.U_NombrePromotor AS NombrePromotor,
                                                T0.ItemCode AS CodigoPlan,
                                                T0.Dscription AS NombrePlan,
                                                T0.U_FormaPago AS FormaDePago,
                                                T2.NAME AS OrigenSolicitud,
                                                T0.U_InvInicial AS InversionInicial,
                                                T0.U_Comision AS Comision,
                                                T0.U_PapeleriaSol AS ImportePapeleria,
                                                T0.U_Importe AS ImporteRecibido,
                                                T0.U_ExcInvIni AS ExcedenteInvIni,
                                                T0.U_Bono AS Bono
                                        FROM    dbo.WTR1 T0
                                                LEFT JOIN dbo.OWTR T1 ON T1.DocEntry = T0.DocEntry
                                                INNER JOIN dbo.[@ORIGSOLICITUD] T2 ON T2.Code = T0.U_OrigenSolicitud
                                                LEFT JOIN dbo.OUSR T4 ON T1.UserSign = T4.USERID
                                        WHERE   NOT EXISTS ( SELECT T3.U_TraspasoRel AS DocNum,
                                                                    T3.U_SolicitudInt AS Serie
                                                             FROM   dbo.OCRD T3
                                                             WHERE  T3.U_TraspasoRel = T1.DocNum
                                                                    AND T3.U_SolicitudInt = T0.U_Serie AND T3.U_TraspasoRel IS NOT NULL)
                                                AND T1.U_TipoMov IN( 'OFICINAS - CONTRATOS','PROMOTOR - ADMON CONTRATOS')
                                                AND T0.U_FormaPago IS NOT NULL
                                                AND T0.U_OrigenSolicitud IS NOT NULL
                                                AND T0.U_Importe IS NOT NULL
                                                AND T0.U_InvInicial IS NOT NULL
                                                AND T0.U_StatusSolicitud <> 'N' -- ERRONEO
                                                AND T0.U_StatusSolicitud <> 'C' -- CANCELADO
                                                AND T0.U_StatusSolicitud <> 'A' -- ATRACO
                                                AND T0.U_StatusSolicitud <> 'E' -- EXTRAVIO
                                                AND T0.U_Serie IS NOT NULL
                                                AND T1.Comments IS NULL
                                                AND T1.DataSource <> 'N'
                                                AND T1.DocStatus <> 'C' " +
                                                qryFechas + qryAlmacen +
                                        " ORDER BY T0.DocEntry,T0.LineNum  ASC ");

                }
                else
                {
                    _oForma.DataSources.DataTables.Item(DT_PRECONTRATOS)
                      .ExecuteQuery(@"SELECT  T1.DocEntry AS DocEntry,
                                                'Crear' AS Crear,
                                                T4.USER_CODE AS CodigoUsuario,
                                                T1.DocNum AS DocNum,
                                                T0.U_Serie AS Serie,
                                                T0.U_CodPromotor AS CodigoPromotor,
                                                T0.U_NombrePromotor AS NombrePromotor,
                                                T0.ItemCode AS CodigoPlan,
                                                T0.Dscription AS NombrePlan,
                                                T0.U_FormaPago AS FormaDePago,
                                                T2.NAME AS OrigenSolicitud,
                                                T0.U_InvInicial AS InversionInicial,
                                                T0.U_Comision AS Comision,
                                                T0.U_PapeleriaSol AS ImportePapeleria,
                                                T0.U_Importe AS ImporteRecibido,
                                                T0.U_ExcInvIni AS ExcedenteInvIni,
                                                T0.U_Bono AS Bono                                                
                                        FROM    dbo.WTR1 T0
                                                LEFT JOIN dbo.OWTR T1 ON T1.DocEntry = T0.DocEntry
                                                INNER JOIN dbo.[@ORIGSOLICITUD] T2 ON T2.Code = T0.U_OrigenSolicitud
                                                LEFT JOIN dbo.OUSR T4 ON T1.UserSign = T4.USERID
                                        WHERE   NOT EXISTS ( SELECT T3.U_TraspasoRel AS DocNum,
                                                                    T3.U_SolicitudInt AS Serie
                                                             FROM   dbo.OCRD T3
                                                             WHERE  T3.U_TraspasoRel = T1.DocNum
                                                                    AND T3.U_SolicitudInt = T0.U_Serie AND T3.U_TraspasoRel IS NOT NULL)
                                                AND T1.U_TipoMov IN( 'OFICINAS - CONTRATOS','PROMOTOR - ADMON CONTRATOS')
                                                AND T0.U_FormaPago IS NOT NULL
                                                AND T0.U_OrigenSolicitud IS NOT NULL
                                                AND T0.U_Importe IS NOT NULL
                                                AND T0.U_InvInicial IS NOT NULL
                                                AND T0.U_Serie IS NOT NULL
                                                AND T0.U_StatusSolicitud <> 'N' -- ERRONEO
                                                AND T0.U_StatusSolicitud <> 'C' -- CANCELADO
                                                AND T0.U_StatusSolicitud <> 'A' -- ATRACO
                                                AND T0.U_StatusSolicitud <> 'E' -- EXTRAVIO
                                                AND T0.U_Serie IS NOT NULL
                                                AND T1.DocStatus <> 'C'
                                                AND T1.DataSource <> 'N'
                                                AND T1.Comments IS NULL
                                        ORDER BY T0.DocEntry,T0.LineNum  ASC      
                                        ");

                }

                _oGridPrecontratos = (SAPbouiCOM.Grid)_oForma.Items.Item(GRDPRECONTRATO).Specific;
                _oGridPrecontratos.DataTable = _oForma.DataSources.DataTables.Item(DT_PRECONTRATOS);
                FormatoGrid(_oGridPrecontratos);

            }
            catch (Exception ex)
            {
                throw new Exception("Error al buscar datos *BuscarDatos* : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        /// <summary>
        /// Se le da el formato al grid para visualizarlo
        /// </summary>
        /// <param name="grid">Nombre del grid</param>
        public void FormatoGrid(Grid grid)
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

                    grid.Columns.Item("DocEntry").Editable = false;
                    grid.Columns.Item("DocEntry").TitleObject.Caption = "DocEntry";
                    SAPbouiCOM.EditTextColumn oColTransferStock = grid.Columns.Item("DocEntry") as SAPbouiCOM.EditTextColumn;
                    oColTransferStock.LinkedObjectType = LINK_TRANSFERENCIA_STOCK;
                    grid.Columns.Item("DocEntry").Width = 60;

                    grid.Columns.Item("Crear").Editable = true;
                    grid.Columns.Item("Crear").TitleObject.Caption = "Crear";
                    grid.Columns.Item("Crear").Width = 90;


                    if (!grid.DataTable.IsEmpty)
                    {
                        grid.Columns.Item("Crear").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;                       
                    }

                    grid.Columns.Item("DocNum").Editable = false;
                    grid.Columns.Item("DocNum").TitleObject.Caption = "Número traspaso";
                    grid.Columns.Item("DocNum").Width = 90;

                    grid.Columns.Item("CodigoUsuario").Editable = false;
                    grid.Columns.Item("CodigoUsuario").TitleObject.Caption = "Usuario";

                    grid.Columns.Item("Serie").Editable = false;
                    grid.Columns.Item("Serie").TitleObject.Caption = "Solicitud";

                    grid.Columns.Item("CodigoPromotor").Editable = false;
                    grid.Columns.Item("CodigoPromotor").TitleObject.Caption = "Código promotor";

                    grid.Columns.Item("NombrePromotor").Editable = false;
                    grid.Columns.Item("NombrePromotor").TitleObject.Caption = "Nombre promotor";

                    grid.Columns.Item("CodigoPlan").Editable = false;
                    grid.Columns.Item("CodigoPlan").TitleObject.Caption = "Código plan";
                    SAPbouiCOM.EditTextColumn oColPlan = grid.Columns.Item("CodigoPlan") as SAPbouiCOM.EditTextColumn;
                    oColPlan.LinkedObjectType = LINK_ARTICULO;

                    grid.Columns.Item("NombrePlan").Editable = false;
                    grid.Columns.Item("NombrePlan").TitleObject.Caption = "Nombre plan";
                    grid.Columns.Item("NombrePlan").Width = 120;

                    grid.Columns.Item("FormaDePago").Editable = false;
                    grid.Columns.Item("FormaDePago").TitleObject.Caption = "Forma de pago";
                    grid.Columns.Item("FormaDePago").Width = 90;

                    grid.Columns.Item("OrigenSolicitud").Editable = false;
                    grid.Columns.Item("OrigenSolicitud").TitleObject.Caption = "Origen de solicitud";
                    grid.Columns.Item("OrigenSolicitud").Width = 90;

                    grid.Columns.Item("InversionInicial").Editable = false;
                    grid.Columns.Item("InversionInicial").TitleObject.Caption = "Inversión inicial";
                    grid.Columns.Item("InversionInicial").Width = 80;

                    grid.Columns.Item("Comision").Editable = false;
                    grid.Columns.Item("Comision").TitleObject.Caption = "Comisión";
                    grid.Columns.Item("Comision").Width = 80;

                    grid.Columns.Item("ImportePapeleria").Editable = false;
                    grid.Columns.Item("ImportePapeleria").TitleObject.Caption = "Importe papelería";
                    grid.Columns.Item("ImportePapeleria").Width = 80;

                    grid.Columns.Item("ImporteRecibido").Editable = false;
                    grid.Columns.Item("ImporteRecibido").TitleObject.Caption = "Importe recibido";
                    grid.Columns.Item("ImporteRecibido").Width = 80;

                    grid.Columns.Item("ExcedenteInvIni").Visible = false;
                    grid.Columns.Item("ExcedenteInvIni").TitleObject.Caption = "Excedente individual. inicial";
                    grid.Columns.Item("ExcedenteInvIni").Width = 90;    

                    grid.Columns.Item("Bono").Visible = false;
                    grid.Columns.Item("Bono").TitleObject.Caption = "Bono";
                    grid.Columns.Item("Bono").Width = 80;

                    grid.CollapseLevel = 1;
                    _oGridPrecontratos.AutoResizeColumns();
                }
                else
                {
                    _Application.MessageBox("No se encontraron registros");
                    grid.DataTable.Clear();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al dar formato al grid *FormatoGrid* : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
                _oForma.Update();
            } 
        }

        /// <summary>
        /// Crea los precontratos
        /// </summary>
        private void CrearPreContratos()
        {
            try
            {
                int i = 0;
                string Seleccionado = null;

                _oGridPrecontratos.CollapseLevel = 0;

               for(int y = _oGridPrecontratos.Rows.Count - 1; y >= 0; y-- )
               {
                    try
                    {
                        Seleccionado = _oGridPrecontratos.DataTable.GetValue("Crear", y).ToString();
                    }
                    catch (Exception ex)
                    {
                        _Application.StatusBar.SetText("Pre-contratos abiertos correctamente...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                        return;
                    }

                    if (Seleccionado == "Y")
                    {
                        lDatos.Clear();
                        itemDatos = new Datos();

                        itemDatos.TrasNoDocumento = _oGridPrecontratos.DataTable.GetValue("DocNum", y).ToString();
                        itemDatos.TrasSerie = Extensor.ObtenerPrefijoSerie(_oGridPrecontratos.DataTable.GetValue("Serie", y).ToString().Substring(0, 6),
                                                                            _oGridPrecontratos.DataTable.GetValue("CodigoPlan", y).ToString(), _Company) +
                                                                            _oGridPrecontratos.DataTable.GetValue("Serie", y).ToString().Substring(6);
                        itemDatos.TrasSerieInterna = _oGridPrecontratos.DataTable.GetValue("Serie", y).ToString();
                        itemDatos.TrasCodigoPlan = _oGridPrecontratos.DataTable.GetValue("CodigoPlan", y).ToString();
                        itemDatos.TrasNombrePlan = _oGridPrecontratos.DataTable.GetValue("NombrePlan", y).ToString();
                        itemDatos.TrasImportePagoInicial = _oGridPrecontratos.DataTable.GetValue("InversionInicial", y);
                        itemDatos.TrasImporteComision = _oGridPrecontratos.DataTable.GetValue("Comision", y);
                        itemDatos.TrasImportePapeleria = _oGridPrecontratos.DataTable.GetValue("ImportePapeleria", y);
                        itemDatos.TrasImporteRecibido = _oGridPrecontratos.DataTable.GetValue("ImporteRecibido", y);
                        itemDatos.TrasExcedenteInvInicial = _oGridPrecontratos.DataTable.GetValue("ExcedenteInvIni", y);
                        itemDatos.TrasImporteBono = _oGridPrecontratos.DataTable.GetValue("Bono", y);
                        lDatos.Add(itemDatos);
                        _Application.StatusBar.SetText("Creando Pre-contrato: " + itemDatos.TrasSerie.ToString() + " por favor espere", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

                        if (lDatos.Count != 0)
                        {
                            Addon.Instance.Ejecutaclase("2561", lDatos);
                        }
                    }

                }                
            }
            catch (Exception ex)
            {
                throw new Exception("Error al ejecutar los Pre-contratos *CrearPreContratos* : " + ex.Message);
            }
        }

        /// <summary>
        /// Selecciona los 10 primeros pre-contratos
        /// </summary>
        private void SelecionAutomatica()
        {
            try
            {
                _oForma.Freeze(true);

                if (_oGridPrecontratos != null)
                {
                    if (_oGridPrecontratos.Rows.Count != 0)
                    {
                        for (int preContrato = 0; preContrato <= _oGridPrecontratos.DataTable.Rows.Count - 1; preContrato++)
                        {
                            _oGridPrecontratos.DataTable.SetValue("Crear", preContrato, "N");
                        }
                        ContadorSeleccionados = 0;

                        if (_oGridPrecontratos.DataTable.Rows.Count - 1 >= 9)
                        {
                            for (int preContrato = 0; preContrato <= 9; preContrato++)
                            {
                                _oGridPrecontratos.DataTable.SetValue("Crear", preContrato, "Y");
                                ContadorSeleccionados = preContrato + 1;
                            }
                        }
                        else
                        {
                            for (int preContrato = 0; preContrato <= _oGridPrecontratos.DataTable.Rows.Count - 1; preContrato++)
                            {
                                _oGridPrecontratos.DataTable.SetValue("Crear", preContrato, "Y");
                                ContadorSeleccionados = preContrato + 1;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al ejecutar selección automática *SelecionAutomatica* : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        #endregion

    }
}
