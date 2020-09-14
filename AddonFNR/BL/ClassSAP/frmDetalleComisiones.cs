using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace AddonFNR.BL
{
    class frmDetalleComisiones : ComportaForm
    {
        #region CONSTANTES

        private const string FRM_DETALLE_COMISIONES = "frmDetalleComisiones";
        private const string TXTCODCOB = "txtCodCob";
        private const string TXTNOMCOB = "txtNomCob";
        private const string BTN_BUSCAR = "btnBuscar";
        private const string BTN_UPDATE = "btnUpdate";
        private const string BTN_CANCELAR = "btnCancel";
        private const string BTN_ACTUALIZAR_PLANES = "btnUpdAsi";
        private const string BTN_SAP_BUSCAR = "1281";
        private const string BTN_SAP_CREAR = "1282";
        private const string DT_DETALLECOMISIONES = "dtDetalleComisiones";
        private const string LINK_DM_COBRADORES = "171";
        private const string LINK_DM_ARTICULO = "4";
        private const string GRD_DETALLE_COMISIONES = "grdDetCom";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForma = null;
        private static bool _oDetalleComisiones = false;
        private SAPbobsCOM.Recordset _oRec = null;
        private SAPbouiCOM.Grid _oGridDetalleComisiones = null;
        private SAPbouiCOM.EditText _oTxtCodeCobrador = null;
        private SAPbouiCOM.EditText _oTxtNomCobrador = null;
        private SAPbouiCOM.Button _oBotonUpdate = null;

        private List<string> columnasGridCalculo = new List<string> { "MontoRecomendado", "MontoAsistente", "MontoBono", "MontoLider", "MontoSupervisor", 
                                                                        "MontoCoordinador", "MontoBonoCoordi", "MontoCoordinador2", "MontoBonoCoordi2", "MontoGerente" };

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
        public frmDetalleComisiones(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oDetalleComisiones == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                showForm(formID);
                inicializarComponentes();
                setEventos();
                _oDetalleComisiones = true;
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
                if (_oDetalleComisiones != false && pVal.FormType != FormTypeMenu && formID == FormUID)
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
                if (pVal.FormUID == formID && pVal.BeforeAction == false)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.ItemUID == GRD_DETALLE_COMISIONES)
                    {
                        _oBotonUpdate.Caption = "Actualizar";
                    }

                    if (pVal.EventType == BoEventTypes.et_LOST_FOCUS && pVal.ItemUID == GRD_DETALLE_COMISIONES && columnasGridCalculo.Contains(pVal.ColUID))
                    {
                        _oGridDetalleComisiones = _oForma.Items.Item(GRD_DETALLE_COMISIONES).Specific;
                        RealizarCalculoFideicomiso(_oGridDetalleComisiones, pVal.Row.ToString());
                    }
                }

                if (pVal.FormUID == formID && pVal.BeforeAction == true)
                {
                    if (pVal.EventType == BoEventTypes.et_PICKER_CLICKED)
                    {
                        _oBotonUpdate.Caption = "Actualizar";
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.Action_Success == false)
                    {
                        if (_oBotonUpdate.Caption == "Actualizar")
                        {
                            int opcion = _Application.MessageBox("¿Desea grabar las modificaciones?", 1, "Si", "No", "Cancelar");

                            if (opcion == 1)
                            {
                                ActualizarInformacion();
                                _Application.ItemEvent -= new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                                _Application.MenuEvent -= new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                                Dispose();
                                application = null;
                                company = null;
                                _oDetalleComisiones = false;
                                Addon.typeList.RemoveAll(p => p._forma == formID);
                                return;
                            }
                            else if (opcion == 2)
                            {
                                _Application.ItemEvent -= new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                                _Application.MenuEvent -= new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                                Dispose();
                                application = null;
                                company = null;
                                _oDetalleComisiones = false;
                                Addon.typeList.RemoveAll(p => p._forma == formID);
                                return;
                            }
                            else if (opcion == 3)
                            {
                                bubbleEvent = false;
                                return;
                            }
                        }
                        else
                        {
                            _Application.ItemEvent -= new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                            _Application.MenuEvent -= new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                            Dispose();
                            application = null;
                            company = null;
                            _oDetalleComisiones = false;
                            Addon.typeList.RemoveAll(p => p._forma == formID);
                            return;
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        _oTxtCodeCobrador = _oForma.Items.Item(TXTCODCOB).Specific;
                        _oTxtNomCobrador = _oForma.Items.Item(TXTNOMCOB).Specific;

                        if (pVal.ItemUID == BTN_ACTUALIZAR_PLANES)
                        {
                            if (!string.IsNullOrEmpty(_oTxtCodeCobrador.Value.ToString()))
                            {
                                if (_Application.MessageBox("¿Desea actualizar los planes?", 2, "Si", "No") == 1)
                                {
                                    ActualizarPlanesAsistente(_oTxtCodeCobrador.Value);
                                }
                            }
                            else
                            {
                                _Application.MessageBox("Capture el código del cobrador");
                            }
                        }

                        if (pVal.ItemUID == BTN_BUSCAR)
                        {
                            if (!string.IsNullOrEmpty(_oTxtCodeCobrador.Value.ToString()) || !string.IsNullOrEmpty(_oTxtNomCobrador.Value.ToString()))
                            {
                                BuscarDatos(_oTxtCodeCobrador.Value, _oTxtNomCobrador.Value);
                            }
                            else
                            {
                                _Application.MessageBox("Capture el código o el nombre del cobrador");
                            }
                        }

                        if (pVal.ItemUID == BTN_CANCELAR)
                        {
                            if (_oBotonUpdate.Caption == "Actualizar")
                            {
                                int opcion = _Application.MessageBox("¿Desea grabar las modificaciones?", 1, "Si", "No", "Cancelar");
                                if (opcion == 1)
                                {
                                    ActualizarInformacion();
                                    _oBotonUpdate.Caption = "Ok";
                                    _oForma.Close();
                                }
                                else if (opcion == 2)
                                {
                                    _oBotonUpdate.Caption = "Ok";
                                    _oForma.Close();
                                }
                                else if (opcion == 3)
                                {
                                    _oForma.Select();
                                }
                            }
                            else
                            {
                                _oForma.Close();
                            }
                        }

                        if (pVal.ItemUID == BTN_UPDATE)
                        {
                            if (_oBotonUpdate.Caption == "Ok")
                            {
                                _oForma.Close();
                            }
                            else
                            {
                                _oTxtNomCobrador.Active = true;
                                ActualizarInformacion();
                                _oBotonUpdate.Caption = "Ok";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en evento *clsDetalleComisiones* : " + ex.Message);
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
                        if (_Application.Forms.ActiveForm.UniqueID == FRM_DETALLE_COMISIONES)
                            BubbleEvent = false;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en ManuEvent *clsDetalleComisiones* : " + ex.Message);
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
                //Declarar DataTable
                _oForma.DataSources.DataTables.Add(DT_DETALLECOMISIONES);
                _oTxtCodeCobrador = _oForma.Items.Item(TXTCODCOB).Specific;
                _oTxtNomCobrador = _oForma.Items.Item(TXTNOMCOB).Specific;
                _oBotonUpdate = _oForma.Items.Item(BTN_UPDATE).Specific;
                _oTxtCodeCobrador.Active = true;
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
        /// Actualiza los planes que no se encuentran en detalle de las comisiones y se encuentran en la tabla @COMISIONES
        /// </summary>
        /// <param name="codigoAsistente">Código del asistente</param>
        /// <param name="nombreAsistente">Nombre del asistente</param>
        private void ActualizarPlanesAsistente(string codigoAsistente)
        {
            try
            {
                _oForma.Freeze(true);
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _Application.StatusBar.SetText("Actualizando planes por favor espere...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);

                _oRec.DoQuery(@"   DECLARE @DocEntryMax INT;
  
                                   SET @DocEntryMax = ( SELECT  ISNULL(MAX(CAST(DocEntry AS INT)), 0)
                                                        FROM    dbo.[@DETALLE_COMISION] WITH ( UPDLOCK )
                                                      ); 

                                   INSERT   INTO dbo.[@DETALLE_COMISION]
                                            ( DocEntry ,
                                              U_Empresa ,
                                              U_Codigo_Plan ,
                                              U_Descripcion_Plan ,
                                              U_Fideicomiso ,
                                              U_Codigo_Cobrador ,
                                              U_Cobrador
                                            )
                                            SELECT DISTINCT
                                                    @DocEntryMax + ROW_NUMBER() OVER ( ORDER BY T0.DocEntry ) AS DocEntry ,
                                                    T0.U_Empresa ,
                                                    T0.U_Codigo_Plan ,
                                                    T0.U_Descripcion_Plan ,
                                                    ( T0.U_Costo - T0.U_Inv_Inicial ) ,
                                                    '" + codigoAsistente + "'," +
                                                    "( SELECT TOP 1 " +
                                                               "U_Cobrador " +
                                                      "FROM      dbo.[@DETALLE_COMISION] " +
                                                      "WHERE     U_Codigo_Cobrador = '" + codigoAsistente + "' " +
                                                    ") " +
                                            "FROM    dbo.[@COMISIONES] T0 " +
                                            "WHERE   NOT EXISTS ( SELECT 1 " +
                                                                 "FROM   dbo.[@DETALLE_COMISION] T1 " +
                                                                 "WHERE  T0.U_Empresa = T1.U_Empresa " +
                                                                        "AND T0.U_Codigo_Plan = T1.U_Codigo_Plan " +
                                                                        "AND T0.U_Descripcion_Plan = T1.U_Descripcion_Plan " +
                                                                        "AND T1.U_Codigo_Cobrador = '" + codigoAsistente + "')");
                _Application.StatusBar.SetText("Planes actualizados...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                BuscarDatos(_oTxtCodeCobrador.Value.ToString(), _oTxtNomCobrador.Value.ToString());

            }
            catch (Exception ex)
            {
                throw new Exception("Error al actualizar los planes *ActualizarPlanesAsistente* : " + ex.Message);
            }
            finally
            {
                try
                {
                    if (_oRec != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oRec);
                    _oForma.Freeze(false);
                }
                catch (Exception)
                {
                }
            }

        }

        /// <summary>
        /// Busca la información de acuerdo a los filtros capturados
        /// </summary>
        /// <param name="efiscal"></param>
        /// <param name="codigo"></param>
        /// <param name="descripcion"></param>
        /// <param name="combos"></param>
        public void BuscarDatos(string codeCobrador, string nameCobrador)
        {
            try
            {
                _oForma.Freeze(true);
                string qryCobrador = null;

                qryCobrador = " and REPLACE(LOWER( " +
                                              "REPLACE(LOWER( " +
                                              "REPLACE(LOWER( " +
                                              "REPLACE(LOWER( " +
                                              "REPLACE(LOWER( " +
                                              "T0.U_Cobrador),'á','a') ), " +
                                                       "'é','e') ), " +
                                                       "'í','i') ), " +
                                                       "'ó','o') ), " +
                                                       "'ú','u') LIKE  '%" + nameCobrador.ToLower().Replace('á', 'a').
                                                                                                   Replace('é', 'e').
                                                                                                   Replace('í', 'i').
                                                                                                   Replace('ó', 'o').
                                                                                                   Replace('ú', 'u') + "%' ";

                _oGridDetalleComisiones = _oForma.Items.Item(GRD_DETALLE_COMISIONES).Specific;
                _oGridDetalleComisiones.DataTable = _oForma.DataSources.DataTables.Item(DT_DETALLECOMISIONES);


                _oGridDetalleComisiones.DataTable.ExecuteQuery(@"SELECT  T2.empID AS CodigoID ,
                                                                        T0.U_Codigo_Cobrador AS CodigoAsistente ,
                                                                        T0.U_Cobrador AS NomAsistente ,
                                                                        T0.U_Empresa AS Empresa ,
                                                                        T0.U_Codigo_Plan AS CodPlan ,
                                                                        T0.U_Descripcion_Plan AS DescPlan ,
                                                                        T1.U_Costo AS CostoPaquete ,
                                                                        T1.U_Inv_Inicial AS InversionInicial ,
                                                                        T0.U_CodigoRecomendado AS CodigoRecomendado ,
                                                                        T0.U_Nom_Recomendado AS NomRecomendado ,
                                                                        ISNULL(T0.U_Recomendado, 0.00) AS MontoRecomendado ,
                                                                        ISNULL(T0.U_Asis_Social, 0.00) AS MontoAsistente ,
                                                                        ISNULL(T0.U_Bono, 0.00) AS MontoBono ,
                                                                        T0.U_CodigoLider AS CodigoLider ,
                                                                        T0.U_Nom_Lider AS NomLider ,
                                                                        ISNULL(T0.U_Lider, 0.00) AS MontoLider ,
                                                                        T0.U_CodigoSupervisor AS CodigoSupervisor ,
                                                                        T0.U_Nom_Supervisor AS NomSupervisor ,
                                                                        ISNULL(T0.U_Supervisor, 0.00) AS MontoSupervisor ,
                                                                        T0.U_CodigoCoordinador AS CodigoCoordinador ,
                                                                        T0.U_Nom_Coordinador AS NomCoordinador ,
                                                                        ISNULL(T0.U_Coordinador, 0.00) AS MontoCoordinador ,
                                                                        ISNULL(T0.U_BonoCoordinador, 0.00) AS MontoBonoCoordi,
                                                                        T0.U_CodigoCoordinador2 AS CodigoCoordinador2 ,
                                                                        T0.U_Nom_Coordinador2 AS NomCoordinador2 ,
                                                                        ISNULL(T0.U_Coordinador2, 0.00) AS MontoCoordinador2 ,
                                                                        ISNULL(T0.U_BonoCoordinador2, 0.00) AS MontoBonoCoordi2,
                                                                        T0.U_CodigoGerente AS CodigoGerente ,
                                                                        T0.U_Nom_Gerente AS NomGerente ,
                                                                        ISNULL(T0.U_Gerente, 0.00) AS MontoGerente ,
                                                                        ISNULL(T0.U_Fideicomiso, 0.00) AS MontoFideicomiso
                                                                FROM    dbo.[@DETALLE_COMISION] T0
                                                                        INNER JOIN dbo.[@COMISIONES] T1 ON T1.U_Codigo_Plan = T0.U_Codigo_Plan
                                                                                                            AND T1.U_Empresa = T0.U_Empresa 
                                                                                                            AND T1.U_Descripcion_Plan = T0.U_Descripcion_Plan
                                                                        LEFT JOIN dbo.OHEM T2 ON T2.firstName = T0.U_Codigo_Cobrador
                                                                WHERE   T0.U_Codigo_Cobrador LIKE '%" + codeCobrador + "%' " + qryCobrador +
                                                                "GROUP BY T2.empID, " +
                                                                            "T0.U_Codigo_Cobrador , " +
                                                                            "T0.U_Cobrador , " +
                                                                            "T0.U_Empresa , " +
                                                                            "T0.U_Codigo_Plan , " +
                                                                            "T0.U_Descripcion_Plan , " +
                                                                            "T1.U_Costo , " +
                                                                            "T1.U_Inv_Inicial , " +
                                                                            "T0.U_CodigoRecomendado , " +
                                                                            "T0.U_Nom_Recomendado , " +
                                                                            "T0.U_Recomendado , " +
                                                                            "T0.U_Asis_Social , " +
                                                                            "T0.U_Bono , " +
                                                                            "T0.U_CodigoLider , " +
                                                                            "T0.U_Nom_Lider , " +
                                                                            "T0.U_Lider , " +
                                                                            "T0.U_CodigoSupervisor , " +
                                                                            "T0.U_Nom_Supervisor , " +
                                                                            "T0.U_Supervisor , " +
                                                                            "T0.U_CodigoCoordinador , " +
                                                                            "T0.U_Nom_Coordinador , " +
                                                                            "T0.U_Coordinador , " +
                                                                            "T0.U_BonoCoordinador , " +
                                                                            "T0.U_CodigoCoordinador2 , " +
                                                                            "T0.U_Nom_Coordinador2 , " +
                                                                            "T0.U_Coordinador2 , " +
                                                                            "T0.U_BonoCoordinador2 , " +
                                                                            "T0.U_CodigoGerente , " +
                                                                            "T0.U_Nom_Gerente , " +
                                                                            "T0.U_Gerente , " +
                                                                            "T0.U_Fideicomiso " +
                                                                "ORDER BY T0.U_Codigo_Cobrador,T2.empID ASC");
                if (_oGridDetalleComisiones.DataTable.IsEmpty)
                {
                    _Application.MessageBox("No se encontraron registros");
                    _oGridDetalleComisiones.DataTable.Clear();
                    _oForma.Update();
                }
                else
                {
                    // _oGridDetalleComisiones.AutoResizeColumns();
                    FormatoGrid(_oGridDetalleComisiones);
                    _oGridDetalleComisiones.AutoResizeColumns();
                }
                _oBotonUpdate.Caption = "Ok";
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
                grid.RowHeaders.TitleObject.Caption = "#";
                for (int noLinea = 0; noLinea < grid.Rows.Count; noLinea++)
                {
                    grid.RowHeaders.SetText(noLinea, (noLinea + 1).ToString());
                }

                grid.Columns.Item("CodigoID").Editable = false;
                grid.Columns.Item("CodigoID").TitleObject.Caption = " ";
                grid.Columns.Item("CodigoID").Width = 15;
                SAPbouiCOM.EditTextColumn oColCodEmp = grid.Columns.Item("CodigoID") as SAPbouiCOM.EditTextColumn;
                oColCodEmp.LinkedObjectType = LINK_DM_COBRADORES;

                grid.Columns.Item("CodigoAsistente").Editable = false;
                grid.Columns.Item("CodigoAsistente").TitleObject.Caption = "Código";

                grid.Columns.Item("NomAsistente").Editable = false;
                grid.Columns.Item("NomAsistente").TitleObject.Caption = "Nombre asistente";

                grid.Columns.Item("Empresa").Editable = false;
                grid.Columns.Item("Empresa").TitleObject.Caption = "Empresa";

                grid.Columns.Item("CodPlan").Editable = false;
                grid.Columns.Item("CodPlan").TitleObject.Caption = "Código plan";
                grid.Columns.Item("CodPlan").Width = 75;
                SAPbouiCOM.EditTextColumn oColCodPlan = grid.Columns.Item("CodPlan") as SAPbouiCOM.EditTextColumn;
                oColCodPlan.LinkedObjectType = LINK_DM_ARTICULO;

                grid.Columns.Item("DescPlan").Editable = false;
                grid.Columns.Item("DescPlan").TitleObject.Caption = "Nombre del plan";

                grid.Columns.Item("CostoPaquete").Editable = false;
                grid.Columns.Item("CostoPaquete").TitleObject.Caption = "Costo de servicio";

                grid.Columns.Item("InversionInicial").Editable = false;
                grid.Columns.Item("InversionInicial").TitleObject.Caption = "Inversión inicial";

                //###########################

                //Recomendado
                grid.Columns.Item("CodigoRecomendado").Editable = true;
                grid.Columns.Item("CodigoRecomendado").TitleObject.Caption = "Código recomendado";
                grid.Columns.Item("NomRecomendado").Editable = true;
                grid.Columns.Item("NomRecomendado").TitleObject.Caption = "Nombre recomendado";
                grid.Columns.Item("MontoRecomendado").Editable = true;
                grid.Columns.Item("MontoRecomendado").TitleObject.Caption = "Monto recomendado";

                //Asistente
                grid.Columns.Item("MontoAsistente").Editable = true;
                grid.Columns.Item("MontoAsistente").TitleObject.Caption = "Monto Asistente";
                grid.Columns.Item("MontoBono").Editable = true;
                grid.Columns.Item("MontoBono").TitleObject.Caption = "Bono asistente";

                //Líder
                grid.Columns.Item("CodigoLider").Editable = true;
                grid.Columns.Item("CodigoLider").TitleObject.Caption = "Código líder";
                grid.Columns.Item("NomLider").Editable = true;
                grid.Columns.Item("NomLider").TitleObject.Caption = "Nombre líder";
                grid.Columns.Item("MontoLider").Editable = true;
                grid.Columns.Item("MontoLider").TitleObject.Caption = "Monto líder";

                //Supervisor
                grid.Columns.Item("CodigoSupervisor").Editable = true;
                grid.Columns.Item("CodigoSupervisor").TitleObject.Caption = "Código supervisor";
                grid.Columns.Item("NomSupervisor").Editable = true;
                grid.Columns.Item("NomSupervisor").TitleObject.Caption = "Nombre supervisor";
                grid.Columns.Item("MontoSupervisor").Editable = true;
                grid.Columns.Item("MontoSupervisor").TitleObject.Caption = "Monto supervisor";

                //Coordinador
                grid.Columns.Item("CodigoCoordinador").Editable = true;
                grid.Columns.Item("CodigoCoordinador").TitleObject.Caption = "Código coordinador";
                grid.Columns.Item("NomCoordinador").Editable = true;
                grid.Columns.Item("NomCoordinador").TitleObject.Caption = "Nombre Coordinador";
                grid.Columns.Item("MontoCoordinador").Editable = true;
                grid.Columns.Item("MontoCoordinador").TitleObject.Caption = "Monto coordinador";
                grid.Columns.Item("MontoBonoCoordi").Editable = true;
                grid.Columns.Item("MontoBonoCoordi").TitleObject.Caption = "Bono coordinador";

                //Coordinador 2
                grid.Columns.Item("CodigoCoordinador2").Editable = true;
                grid.Columns.Item("CodigoCoordinador2").TitleObject.Caption = "Código coordinador 2";
                grid.Columns.Item("NomCoordinador2").Editable = true;
                grid.Columns.Item("NomCoordinador2").TitleObject.Caption = "Nombre coordinador 2";
                grid.Columns.Item("MontoCoordinador2").Editable = true;
                grid.Columns.Item("MontoCoordinador2").TitleObject.Caption = "Monto coordinador 2";
                grid.Columns.Item("MontoBonoCoordi2").Editable = true;
                grid.Columns.Item("MontoBonoCoordi2").TitleObject.Caption = "Bono coordinador 2";

                //Gerente
                grid.Columns.Item("CodigoGerente").Editable = true;
                grid.Columns.Item("CodigoGerente").TitleObject.Caption = "Código gerente";
                grid.Columns.Item("NomGerente").Editable = true;
                grid.Columns.Item("NomGerente").TitleObject.Caption = "Nombre gerente";
                grid.Columns.Item("MontoGerente").Editable = true;
                grid.Columns.Item("MontoGerente").TitleObject.Caption = "Monto gerente";

             

                grid.Columns.Item("MontoFideicomiso").Editable = false;
                grid.Columns.Item("MontoFideicomiso").TitleObject.Caption = "Fideicomiso";

                RealizarCalculoFideicomiso(_oGridDetalleComisiones, "");
            }
            catch (Exception ex)
            {
                throw new Exception("Error al dar formato al grid *FormatoGrid* : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        /// <summary>
        /// Se actualizan los datos en la base de datos 
        /// </summary>
        private void ActualizarInformacion()
        {
            try
            {
                int i = 0;
                int CodigoID = 0;
                string CodAsistente = null;
                string Empresa = null;
                string CodigoPlan = null;
                string DescripcionPlan = null;
                double CostoPaquete = 0;
                double InversionIni = 0;
                string codRecomendado = null;
                string nomRecomendado = null;
                double MontoRecomendado = 0;
                double MontoAsistente = 0;
                string codLider = null;
                string nomLider = null;
                double MontoLider = 0;
                string CodSupervisor = null;
                string nomSupervisor = null;
                double MontoSupervisor = 0;
                string CodCoordinador = null;
                string nomCoordinador = null;
                double MontoCoordinador = 0;
                string CodCoordinador2 = null;
                string nomCoordinador2 = null;
                double MontoCoordinador2 = 0;
                string CodGerente = null;
                string nomGerente = null;
                double MontoGerente = 0;
                double MontoBono = 0;
                double MontoBonoCoordinador = 0;
                double MontoBonoCoordinador2 = 0;
                double MontoFideicomiso = 0;

                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _Application.StatusBar.SetText("Guardando datos por favor espere...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                while (_oGridDetalleComisiones.Rows.Count - 1 >= i)
                {
                    CodigoID = Convert.ToInt32(_oGridDetalleComisiones.DataTable.GetValue("CodigoID", i));
                    CodAsistente = _oGridDetalleComisiones.DataTable.GetValue("CodigoAsistente", i).ToString();
                    Empresa = _oGridDetalleComisiones.DataTable.GetValue("Empresa", i).ToString();
                    CodigoPlan = _oGridDetalleComisiones.DataTable.GetValue("CodPlan", i).ToString();
                    DescripcionPlan = _oGridDetalleComisiones.DataTable.GetValue("DescPlan", i).ToString();
                    CostoPaquete = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("CostoPaquete", i));
                    InversionIni = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("InversionInicial", i));
                    codRecomendado = _oGridDetalleComisiones.DataTable.GetValue("CodigoRecomendado", i).ToString();
                    nomRecomendado = _oGridDetalleComisiones.DataTable.GetValue("NomRecomendado", i).ToString();
                    MontoRecomendado = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoRecomendado", i));
                    MontoAsistente = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoAsistente", i));
                    MontoBono = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoBono", i));
                    codLider = _oGridDetalleComisiones.DataTable.GetValue("CodigoLider", i).ToString();
                    nomLider = _oGridDetalleComisiones.DataTable.GetValue("NomLider", i).ToString();
                    MontoLider = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoLider", i));
                    CodSupervisor = _oGridDetalleComisiones.DataTable.GetValue("CodigoSupervisor", i).ToString();
                    nomSupervisor = _oGridDetalleComisiones.DataTable.GetValue("NomSupervisor", i).ToString();
                    MontoSupervisor = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoSupervisor", i));
                    CodCoordinador = _oGridDetalleComisiones.DataTable.GetValue("CodigoCoordinador", i).ToString();
                    nomCoordinador = _oGridDetalleComisiones.DataTable.GetValue("NomCoordinador", i).ToString();
                    MontoCoordinador = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoCoordinador", i));
                    MontoBonoCoordinador = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoBonoCoordi", i));
                    CodCoordinador2 = _oGridDetalleComisiones.DataTable.GetValue("CodigoCoordinador2", i).ToString();
                    nomCoordinador2 = _oGridDetalleComisiones.DataTable.GetValue("NomCoordinador2", i).ToString();
                    MontoCoordinador2 = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoCoordinador2", i));
                    MontoBonoCoordinador2 = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoBonoCoordi2", i));
                    CodGerente = _oGridDetalleComisiones.DataTable.GetValue("CodigoGerente", i).ToString();
                    nomGerente = _oGridDetalleComisiones.DataTable.GetValue("NomGerente", i).ToString();
                    MontoGerente = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoGerente", i));
                    
                    MontoFideicomiso = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoFideicomiso", i));


                    _oRec.DoQuery(@"UPDATE  T0
                                    SET     T0.U_CodigoRecomendado = '" + codRecomendado + "', " +
                                            "T0.U_Nom_Recomendado = '" + nomRecomendado + "', " +
                                            "T0.U_Recomendado = '" + MontoRecomendado + "', " +
                                            "T0.U_Asis_Social = '" + MontoAsistente + "', " +
                                            "T0.U_CodigoLider = '" + codLider + "', " +
                                            "T0.U_Nom_Lider = '" + nomLider + "', " +
                                            "T0.U_Lider = '" + MontoLider + "', " +
                                            "T0.U_CodigoSupervisor = '" + CodSupervisor + "', " +
                                            "T0.U_Nom_Supervisor = '" + nomSupervisor + "', " +
                                            "T0.U_Supervisor = '" + MontoSupervisor + "', " +
                                            "T0.U_CodigoCoordinador = '" + CodCoordinador + "', " +
                                            "T0.U_Nom_Coordinador = '" + nomCoordinador + "', " +
                                            "T0.U_Coordinador = '" + MontoCoordinador + "', " +
                                            "T0.U_CodigoCoordinador2 = '" + CodCoordinador2 + "', " +
                                            "T0.U_Nom_Coordinador2 = '" + nomCoordinador2 + "', " +
                                            "T0.U_Coordinador2 = '" + MontoCoordinador2 + "', " +
                                            "T0.U_CodigoGerente = '" + CodGerente + "', " +
                                            "T0.U_Nom_Gerente = '" + nomGerente + "', " +
                                            "T0.U_Gerente = '" + MontoGerente + "', " +
                                            "T0.U_Bono = '" + MontoBono + "', " +
                                            "T0.U_BonoCoordinador = '" + MontoBonoCoordinador + "', " +
                                            "T0.U_BonoCoordinador2 = '" + MontoBonoCoordinador2 + "', " +
                                            "T0.U_Fideicomiso = '" + MontoFideicomiso + "' " +
                                    "FROM    dbo.[@DETALLE_COMISION] T0 " +
                                            "LEFT JOIN dbo.OHEM T1 ON T1.firstName = T0.U_Codigo_Cobrador " +
                                    "WHERE   T1.empID = '" + CodigoID + "' " +
                                            "AND T0.U_Codigo_Cobrador = '" + CodAsistente + "' " +
                                            "AND T0.U_Empresa = '" + Empresa + "' " +
                                            "AND T0.U_Descripcion_Plan = '" + DescripcionPlan + "' " +
                                            "AND T0.U_Codigo_Plan = '" + CodigoPlan + "' ");
                    i += 1;
                }
                _Application.StatusBar.SetText("Datos guardados correctamente...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                BuscarDatos(_oTxtCodeCobrador.Value.ToString(), _oTxtNomCobrador.Value.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception("Error al guardar datos del cobrador *ActualizarInformacion* : " + ex.Message);
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
        /// Realiza el calculo del fideicomiso de acuerdo a lo que se tiene en las columnas
        /// </summary>
        /// <param name="_oGridDetalleComisiones">Grid de las comisiones</param>
        private void RealizarCalculoFideicomiso(Grid _oGridDetalleComisiones, string lineaAfectada)
        {
            try
            {
                _oForma.Freeze(true);
                int i = 0;
                double CostoPaquete = 0;
                double InversionIni = 0;
                double Recomendado = 0;
                double Asistente = 0;
                double Bono = 0;
                double Lider = 0;
                double Supervisor = 0;
                double Coordinador = 0;
                double BonoCoordinador = 0;
                double Coordinador2 = 0;
                double BonoCoordinador2 = 0;
                double Gerente = 0;
                double fideicomiso = 0;

                if (!string.IsNullOrEmpty(lineaAfectada))
                {
                    CostoPaquete = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("CostoPaquete", Convert.ToInt32(lineaAfectada)));
                    InversionIni = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("InversionInicial", Convert.ToInt32(lineaAfectada)));
                    Recomendado = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoRecomendado", Convert.ToInt32(lineaAfectada)));
                    Asistente = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoAsistente", Convert.ToInt32(lineaAfectada)));
                    Bono = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoBono", Convert.ToInt32(lineaAfectada)));
                    Lider = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoLider", Convert.ToInt32(lineaAfectada)));
                    Supervisor = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoSupervisor", Convert.ToInt32(lineaAfectada)));
                    Coordinador = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoCoordinador", Convert.ToInt32(lineaAfectada)));
                    BonoCoordinador = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoBonoCoordi", Convert.ToInt32(lineaAfectada)));
                    Coordinador2 = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoCoordinador2", Convert.ToInt32(lineaAfectada)));
                    BonoCoordinador2 = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoBonoCoordi2", Convert.ToInt32(lineaAfectada)));
                    Gerente = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoGerente", Convert.ToInt32(lineaAfectada)));
                    fideicomiso = CostoPaquete - (InversionIni + Recomendado + Asistente + Bono + Lider + Supervisor + Coordinador + BonoCoordinador + Coordinador2 + BonoCoordinador2 + Gerente);
                    //_oGridDetalleComisiones.DataTable.SetValue("Fideicomiso", Convert.ToInt32(lineaAfectada), fideicomiso);
                    ((SAPbouiCOM.EditTextColumn)_oGridDetalleComisiones.Columns.Item("MontoFideicomiso")).SetText(Convert.ToInt32(lineaAfectada), fideicomiso.ToString());

                    var recomendado = Recomendado > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //11
                    var asistente = Asistente > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //12
                    var bono = Bono > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //13
                    var lider = Lider > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //16
                    var supervisor = Supervisor > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //19
                    var coordinador = Coordinador > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //22
                    var bonoCoordinador = BonoCoordinador > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb();//23
                    var coordinador2 = Coordinador2 > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //26
                    var bonoCoordinador2 = BonoCoordinador2 > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb();//27
                    var gerente = Gerente > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //30


                    _oGridDetalleComisiones.CommonSetting.SetCellBackColor(Convert.ToInt32(lineaAfectada) + 1, 11, recomendado);
                    _oGridDetalleComisiones.CommonSetting.SetCellBackColor(Convert.ToInt32(lineaAfectada) + 1, 12, asistente);
                    _oGridDetalleComisiones.CommonSetting.SetCellBackColor(Convert.ToInt32(lineaAfectada) + 1, 13, bono);
                    _oGridDetalleComisiones.CommonSetting.SetCellBackColor(Convert.ToInt32(lineaAfectada) + 1, 16, lider);
                    _oGridDetalleComisiones.CommonSetting.SetCellBackColor(Convert.ToInt32(lineaAfectada) + 1, 19, supervisor);
                    _oGridDetalleComisiones.CommonSetting.SetCellBackColor(Convert.ToInt32(lineaAfectada) + 1, 22, coordinador);
                    _oGridDetalleComisiones.CommonSetting.SetCellBackColor(Convert.ToInt32(lineaAfectada) + 1, 23, bonoCoordinador);
                    _oGridDetalleComisiones.CommonSetting.SetCellBackColor(Convert.ToInt32(lineaAfectada) + 1, 26, coordinador2);
                    _oGridDetalleComisiones.CommonSetting.SetCellBackColor(Convert.ToInt32(lineaAfectada) + 1, 27, bonoCoordinador2);
                    _oGridDetalleComisiones.CommonSetting.SetCellBackColor(Convert.ToInt32(lineaAfectada) + 1, 30, gerente);


                    if (fideicomiso < 0)
                    {
                        _oGridDetalleComisiones.CommonSetting.SetRowFontColor(Convert.ToInt32(lineaAfectada) + 1, Color.Blue.ToArgb());
                    }
                    else
                    {
                        _oGridDetalleComisiones.CommonSetting.SetRowFontColor(Convert.ToInt32(lineaAfectada) + 1, Color.Black.ToArgb());
                    }

                }
                else
                {
                    while (_oGridDetalleComisiones.Rows.Count - 1 >= i)
                    {
                        CostoPaquete = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("CostoPaquete", i));
                        InversionIni = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("InversionInicial", i));
                        Recomendado = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoRecomendado", i));
                        Asistente = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoAsistente", i));
                        Bono = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoBono", i));
                        Lider = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoLider", i));
                        Supervisor = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoSupervisor", i));
                        Coordinador = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoCoordinador", i));
                        BonoCoordinador = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoBonoCoordi", i));
                        Coordinador2 = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoCoordinador2", i));
                        BonoCoordinador2 = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoBonoCoordi2", i));
                        Gerente = Convert.ToDouble(_oGridDetalleComisiones.DataTable.GetValue("MontoGerente", i));
                        fideicomiso = CostoPaquete - (InversionIni + Recomendado + Asistente + Bono + Lider + Supervisor + Coordinador + BonoCoordinador + Coordinador2 + BonoCoordinador2 + Gerente);
                        _oGridDetalleComisiones.DataTable.SetValue("MontoFideicomiso", i, fideicomiso);

                        var recomendado = Recomendado > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //11
                        var asistente = Asistente > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //12
                        var bono = Bono > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //13
                        var lider = Lider > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //16
                        var supervisor = Supervisor > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //19
                        var coordinador = Coordinador > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //22
                        var bonoCoordinador = BonoCoordinador > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb();//23
                        var coordinador2 = Coordinador2 > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //26
                        var bonoCoordinador2 = BonoCoordinador2 > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb();//27
                        var gerente = Gerente > 0 ? Color.White.ToArgb() : Color.LightGray.ToArgb(); //30


                        _oGridDetalleComisiones.CommonSetting.SetCellBackColor(i + 1, 11, recomendado);
                        _oGridDetalleComisiones.CommonSetting.SetCellBackColor(i + 1, 12, asistente);
                        _oGridDetalleComisiones.CommonSetting.SetCellBackColor(i + 1, 13, bono);
                        _oGridDetalleComisiones.CommonSetting.SetCellBackColor(i + 1, 16, lider);
                        _oGridDetalleComisiones.CommonSetting.SetCellBackColor(i + 1, 19, supervisor);
                        _oGridDetalleComisiones.CommonSetting.SetCellBackColor(i + 1, 22, coordinador);
                        _oGridDetalleComisiones.CommonSetting.SetCellBackColor(i + 1, 23, bonoCoordinador);
                        _oGridDetalleComisiones.CommonSetting.SetCellBackColor(i + 1, 26, coordinador2);
                        _oGridDetalleComisiones.CommonSetting.SetCellBackColor(i + 1, 27, bonoCoordinador2);
                        _oGridDetalleComisiones.CommonSetting.SetCellBackColor(i + 1, 30, gerente);


                        if (fideicomiso < 0)
                        {
                            _oGridDetalleComisiones.CommonSetting.SetRowFontColor(i + 1, Color.Blue.ToArgb());
                        }
                        else
                        {
                            _oGridDetalleComisiones.CommonSetting.SetRowFontColor(i + 1, Color.Black.ToArgb());
                        }

                        i += 1;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al realizar calculo de Fideicomiso *RealizarCalculoFideicomiso* : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        #endregion

        //#region MATRIX

        //private void CrearMatrix(Form _oForma)
        //{
        //    SAPbouiCOM.Item item = null;
        //    SAPbouiCOM.Button button = null;
        //    SAPbouiCOM.Matrix oMatrix = null;
        //    SAPbouiCOM.Column oColumn = null;
        //    SAPbouiCOM.Columns oColumns = null;
        //    SAPbouiCOM.DBDataSource oDBDataSource = null;
        //    SAPbouiCOM.UserDataSource oUserDataSource = null;

        //    item = _oForma.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
        //    item.Left = 5;
        //    item.Width = 65;
        //    item.Top = 350;
        //    item.Height = 19;

        //    button = item.Specific;
        //    button.Caption = "OK";

        //    item = _oForma.Items.Add("Matrix1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
        //    item.Left = 5;
        //    item.Width = 500;
        //    item.Top = 200;
        //    item.Height = 150;

        //    oMatrix = item.Specific;
        //    oColumns = oMatrix.Columns;


        //    oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
        //    oColumn.TitleObject.Caption = "#";
        //    oColumn.Width = 30;
        //    oColumn.Editable = false;

        //    oColumn = oColumns.Add("DSCardCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
        //    oColumn.TitleObject.Caption = "Card Code";
        //    oColumn.Width = 40;
        //    oColumn.Editable = true;

        //    oColumn = oColumns.Add("DSCardName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
        //    oColumn.TitleObject.Caption = "Name";
        //    oColumn.Width = 40;
        //    oColumn.Editable = true;

        //    oColumn = oColumns.Add("DSPhone", SAPbouiCOM.BoFormItemTypes.it_EDIT);
        //    oColumn.TitleObject.Caption = "Phone";
        //    oColumn.Width = 40;
        //    oColumn.Editable = true;

        //    //oColumn = oColumns.Add("DSPhoneInt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
        //    //oColumn.TitleObject.Caption = "Int. Phone";
        //    //oColumn.Width = 40;
        //    //oColumn.Editable = true;
        //    //###
        //    oUserDataSource = _oForma.DataSources.UserDataSources.Add("IntPhone", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
        //    oDBDataSource = _oForma.DataSources.DBDataSources.Add("@DETALLE_COMISION");
        //    //###
        //    oColumn = oColumns.Item("DSCardCode");
        //    oColumn.DataBind.SetBound(true, "@DETALLE_COMISION", "U_codigo_emp");

        //    oColumn = oColumns.Item("DSCardName");
        //    oColumn.DataBind.SetBound(true, "@DETALLE_COMISION", "U_empleado");

        //    oColumn = oColumns.Item("DSPhone");
        //    oColumn.DataBind.SetBound(true, "@DETALLE_COMISION", "U_costo_paquete");


        ////oColumn = oColumns.Item("DSPhoneInt");
        ////oColumn.DataBind.SetBound(true, "", "IntPhone");

        //    //####
        //oMatrix.Clear();
        //oMatrix.AutoResizeColumns();

        //oDBDataSource.Query(null);
        //oUserDataSource.Value = "Phone with prefix";
        //oMatrix.LoadFromDataSource();

        //oMatrix.FlushToDataSource();

        //}

        //#endregion
    }
}
