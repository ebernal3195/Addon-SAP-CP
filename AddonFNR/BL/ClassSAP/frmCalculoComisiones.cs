using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace AddonFNR.BL
{
    class frmCalculoComisiones : ComportaForm
    {
        #region CONSTANTES

        private const string FRM_CALCULO_COMISIONES = "frmCalculoComisiones";
        private const string TXT_CONTRATO = "txtContra";
        private const string BTN_BUSCAR = "btnBuscar";
        private const string BTN_UPDATE = "btnUpdate";
        private const string BTN_CANCELAR = "btnCancel";
        private const string BTN_SAP_BUSCAR = "1281";
        private const string BTN_SAP_CREAR = "1282";
        private const string DT_CALCULO_COMISIONES = "dtCalculoComisiones";
        private const string LINK_DM_SOCIO_NEGOCIO = "2";
        private const string LINK_DM_COBRADORES = "171";
        private const string LINK_DM_ARTICULO = "4";
        private const string LINK_TRANSFERENCIA = "67";
        private const string LINK_FACTURA = "13";
        private const string GRD_CALCULO_COMISIONES = "grdCalCom";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForma = null;
        private static bool _oCalculoComisiones = false;
        private SAPbobsCOM.Recordset _oRec = null;
        private SAPbouiCOM.Grid _oGridCalculoComisiones = null;
        private SAPbouiCOM.EditText _oTxtContrato = null;
        private SAPbouiCOM.Button _oBotonUpdate = null;

        private List<string> columnasGridCalculo = new List<string> { "MontoRecomendado", "MontoAsistente", "MontoBonoAsistente", "MontoLider", "MontoSupervisor",
                                                                       "MontoCoordinador", "MontoBonoCoordi", "MontoCoordinador2", "MontoBonoCoordi2", "MontoGerente" };

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de Calculo de comisiones
        /// </summary>
        /// <param name="_Application">Este es el objeto raíz de la API de interfaz de usuario
        ///                             lo que refleja la cual aplicación SAP Business One en el que se realiza 
        ///                             la conexión</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        public frmCalculoComisiones(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oCalculoComisiones == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                showForm(formID);
                inicializarComponentes();
                setEventos();
                _oCalculoComisiones = true;
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
                if (_oCalculoComisiones != false && pVal.FormType != FormTypeMenu && formID == FormUID)
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
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.ItemUID == GRD_CALCULO_COMISIONES)
                    {
                        _oBotonUpdate.Caption = "Actualizar";
                    }

                    if (pVal.EventType == BoEventTypes.et_LOST_FOCUS && pVal.ItemUID == GRD_CALCULO_COMISIONES && columnasGridCalculo.Contains(pVal.ColUID))
                    {
                        _oGridCalculoComisiones = _oForma.Items.Item(GRD_CALCULO_COMISIONES).Specific;
                        RealizarCalculoFideicomiso(_oGridCalculoComisiones, pVal.Row.ToString());
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
                                _oCalculoComisiones = false;
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
                                _oCalculoComisiones = false;
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
                            _oCalculoComisiones = false;
                            Addon.typeList.RemoveAll(p => p._forma == formID);
                            return;
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if (pVal.ItemUID == BTN_BUSCAR)
                        {
                            _oTxtContrato = _oForma.Items.Item(TXT_CONTRATO).Specific;

                            if (!string.IsNullOrEmpty(_oTxtContrato.Value.ToString()))
                            {
                                BuscarDatos(_oTxtContrato.Value);
                            }
                            else
                            {
                                _Application.MessageBox("Capture el contrato");
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
                                _oTxtContrato.Active = true;
                                ActualizarInformacion();
                                _oBotonUpdate.Caption = "Ok";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en evento *clsCalculoComisiones* : " + ex.Message);
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
                        if (_Application.Forms.ActiveForm.UniqueID == FRM_CALCULO_COMISIONES)
                            BubbleEvent = false;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en ManuEvent *clsCalculoComisiones* : " + ex.Message);
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
                _oForma.DataSources.DataTables.Add(DT_CALCULO_COMISIONES);
                _oTxtContrato = _oForma.Items.Item(TXT_CONTRATO).Specific;
                _oBotonUpdate = _oForma.Items.Item(BTN_UPDATE).Specific;
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
        public void BuscarDatos(string contrato)
        {
            try
            {
                _oForma.Freeze(true);
                string qryContrato = null;

                qryContrato = " REPLACE(LOWER( " +
                                              "REPLACE(LOWER( " +
                                              "REPLACE(LOWER( " +
                                              "REPLACE(LOWER( " +
                                              "REPLACE(LOWER( " +
                                              "T0.U_Contrato),'á','a') ), " +
                                                       "'é','e') ), " +
                                                       "'í','i') ), " +
                                                       "'ó','o') ), " +
                                                       "'ú','u') LIKE  '%" + contrato.ToLower().Replace('á', 'a').
                                                                                                   Replace('é', 'e').
                                                                                                   Replace('í', 'i').
                                                                                                   Replace('ó', 'o').
                                                                                                   Replace('ú', 'u') + "%' ";

                _oGridCalculoComisiones = _oForma.Items.Item(GRD_CALCULO_COMISIONES).Specific;
                _oGridCalculoComisiones.DataTable = _oForma.DataSources.DataTables.Item(DT_CALCULO_COMISIONES);


                _oGridCalculoComisiones.DataTable.ExecuteQuery(@"SELECT  T1.CardCode AS CardCode ,
                                                                        T0.U_Contrato AS NumContrato ,
                                                                        --T0.U_ContratoInterno AS NumContratoInterno,
                                                                        T0.U_DocEntryTransfer AS DocEntryTransferencia ,
                                                                        T0.U_DocEntryFactura AS DocEntryFactura ,
                                                                        T0.U_Empresa AS Empresa ,
                                                                        T1.U_NumArt_ AS CodigoPlan ,
                                                                        T1.U_Dsciption AS NombrePlan ,
                                                                        T1.U_CostoTotalPaquete AS CostoPlan ,
                                                                        T2.empID AS IDCobrador ,
                                                                        T0.U_Codigo_Cobrador AS CodigoCobrador ,
                                                                        T0.U_Nombre_Cobrador AS NombreCobrador , 
		                                                                --#### CALCULO INICIO
                                                                        T0.U_Inv_Inicial AS InversionInicial ,
                                                                        T0.U_CodigoRecomendado AS CodigoRecomendado ,
                                                                        T0.U_Nom_Recomendado AS NombreRecomendado ,
                                                                        T0.U_Recomendado AS MontoRecomendado ,
                                                                        T0.U_Codigo_Asistente AS CodigoAsistente ,
                                                                        T0.U_Asistente_Social AS NombreAsistente ,
                                                                        T0.U_Asis_Social AS MontoAsistente ,
                                                                        T0.U_BonoAsistente AS MontoBonoAsistente ,
                                                                        --T0.U_Bono AS MontoBonoCliente,
                                                                        T0.U_CodigoLider AS CodigoLider ,
                                                                        T0.U_Nom_Lider AS NombreLider ,
                                                                        T0.U_Lider AS MontoLider ,
                                                                        T0.U_CodigoSupervisor AS CodigoSupervisor ,
                                                                        T0.U_Nom_Supervisor AS NombreSupervisor ,
                                                                        T0.U_Supervisor AS MontoSupervisor ,
                                                                        T0.U_CodigoCoordinador AS CodigoCoordinador ,
                                                                        T0.U_Nom_Coordinador AS NombreCoordinador ,
                                                                        T0.U_Coordinador AS MontoCoordinador ,
                                                                        T0.U_BonoCoordinador AS MontoBonoCoordi,
                                                                        T0.U_CodigoCoordinador2 AS CodigoCoordinador2 ,
                                                                        T0.U_Nom_Coordinador2 AS NombreCoordinador2 ,
                                                                        T0.U_Coordinador2 AS MontoCoordinador2 ,
                                                                        T0.U_BonoCoordinador2 AS MontoBonoCoordi2,
                                                                        T0.U_CodigoGerente AS CodigoGerente ,
                                                                        T0.U_Nom_Gerente AS NombreGerente ,
                                                                        T0.U_Gerente AS MontoGerente ,
                                                                        T0.U_Fideicomiso AS MontoFideicomiso ,
                                                                        T0.U_FechaCreacion AS FechaCreacion ,
                                                                        T0.U_His_Recomendado AS HisRecomendado,
																		T0.U_His_Asis_Social AS HisAsistente,
																		T0.U_His_BonoAsis AS HisBonoAsistente,
																		T0.U_His_Lider AS HisLider,
																		T0.U_His_Supervisor AS HisSupervisor,
																		T0.U_His_Coordinador AS HisCoordinador,
																		T0.U_His_BonoCoord AS HisBonoCoordinador,
																		T0.U_His_Coordinador2 AS HisCoordinador2,
																		T0.U_His_BonoCoord2 AS HisBonoCoordinador2,
																		T0.U_His_Gerente AS HisGerente
                                                                        --#### CALCULO FIN
                                                                FROM    dbo.[@CALCULO_COMISIONES] T0
                                                                        INNER JOIN dbo.OCRD T1 ON T1.CardCode = T0.U_Contrato -- T1.U_SolicitudInt = T0.U_ContratoInterno
                                                                        LEFT JOIN dbo.OHEM T2 ON T2.firstName = T0.U_Codigo_Cobrador 
                                                                        --INNER JOIN dbo.[@COMISIONES] T3 ON T3.U_Codigo_Plan = T1.U_NumArt_ AND T3.U_Empresa = T0.U_Empresa
                                                                        WHERE " + qryContrato);
                if (_oGridCalculoComisiones.DataTable.IsEmpty)
                {
                    _Application.MessageBox("No se encontraron registros");
                    _oGridCalculoComisiones.DataTable.Clear();
                    _oForma.Update();
                }
                else
                {
                    // _oGridDetalleComisiones.AutoResizeColumns();
                    FormatoGrid(_oGridCalculoComisiones);
                    _oGridCalculoComisiones.AutoResizeColumns();
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

                grid.Columns.Item("CardCode").Editable = false;
                grid.Columns.Item("CardCode").TitleObject.Caption = "Cliente";
                SAPbouiCOM.EditTextColumn oColCardcode = grid.Columns.Item("CardCode") as SAPbouiCOM.EditTextColumn;
                oColCardcode.LinkedObjectType = LINK_DM_SOCIO_NEGOCIO;

                grid.Columns.Item("NumContrato").Editable = false;
                grid.Columns.Item("NumContrato").TitleObject.Caption = "Contrato";

                grid.Columns.Item("DocEntryTransferencia").Visible = false;
                grid.Columns.Item("DocEntryTransferencia").Editable = false;
                grid.Columns.Item("DocEntryTransferencia").TitleObject.Caption = "Transferencia";
                SAPbouiCOM.EditTextColumn oColTransfer = grid.Columns.Item("DocEntryTransferencia") as SAPbouiCOM.EditTextColumn;
                oColTransfer.LinkedObjectType = LINK_TRANSFERENCIA;

                grid.Columns.Item("DocEntryFactura").Editable = false;
                grid.Columns.Item("DocEntryFactura").TitleObject.Caption = "Factura";
                SAPbouiCOM.EditTextColumn oColFactura = grid.Columns.Item("DocEntryFactura") as SAPbouiCOM.EditTextColumn;
                oColFactura.LinkedObjectType = LINK_FACTURA;

                grid.Columns.Item("Empresa").Editable = false;
                grid.Columns.Item("Empresa").TitleObject.Caption = "Empresa";

                grid.Columns.Item("CodigoPlan").Editable = false;
                grid.Columns.Item("CodigoPlan").TitleObject.Caption = "Código plan";
                SAPbouiCOM.EditTextColumn oColCodPlan = grid.Columns.Item("CodigoPlan") as SAPbouiCOM.EditTextColumn;
                oColCodPlan.LinkedObjectType = LINK_DM_ARTICULO;

                grid.Columns.Item("NombrePlan").Editable = false;
                grid.Columns.Item("NombrePlan").TitleObject.Caption = "Nombre plan";

                grid.Columns.Item("CostoPlan").Editable = false;
                grid.Columns.Item("CostoPlan").TitleObject.Caption = "Costo";

                grid.Columns.Item("IDCobrador").Editable = false;
                grid.Columns.Item("IDCobrador").TitleObject.Caption = " ";
                SAPbouiCOM.EditTextColumn oColCodEmp = grid.Columns.Item("IDCobrador") as SAPbouiCOM.EditTextColumn;
                oColCodEmp.LinkedObjectType = LINK_DM_COBRADORES;

                grid.Columns.Item("CodigoCobrador").Editable = true;
                grid.Columns.Item("CodigoCobrador").TitleObject.Caption = "Código cobrador";

                grid.Columns.Item("NombreCobrador").Editable = true;
                grid.Columns.Item("NombreCobrador").TitleObject.Caption = "Nombre cobrador";

                //###########################
                grid.Columns.Item("InversionInicial").Editable = false;
                grid.Columns.Item("InversionInicial").TitleObject.Caption = "Inversión inicial";

                grid.Columns.Item("CodigoRecomendado").Editable = true;
                grid.Columns.Item("CodigoRecomendado").TitleObject.Caption = "Código recomendado";
                grid.Columns.Item("NombreRecomendado").Editable = true;
                grid.Columns.Item("NombreRecomendado").TitleObject.Caption = "Nombre recomendado";
                grid.Columns.Item("MontoRecomendado").Editable = true;
                grid.Columns.Item("MontoRecomendado").TitleObject.Caption = "Monto recomendado";

                grid.Columns.Item("CodigoAsistente").Editable = true;
                grid.Columns.Item("CodigoAsistente").TitleObject.Caption = "Código asistente social";
                grid.Columns.Item("NombreAsistente").Editable = true;
                grid.Columns.Item("NombreAsistente").TitleObject.Caption = "Nombre asistente social";
                grid.Columns.Item("MontoAsistente").Editable = true;
                grid.Columns.Item("MontoAsistente").TitleObject.Caption = "Monto asistente";
                grid.Columns.Item("MontoBonoAsistente").Editable = true;
                grid.Columns.Item("MontoBonoAsistente").TitleObject.Caption = "Bono asistente";

                //grid.Columns.Item("MontoBonoCliente").Editable = false;
                //grid.Columns.Item("MontoBonoCliente").TitleObject.Caption = "Bono cliente";

                grid.Columns.Item("CodigoLider").Editable = true;
                grid.Columns.Item("CodigoLider").TitleObject.Caption = "Código líder";
                grid.Columns.Item("NombreLider").Editable = true;
                grid.Columns.Item("NombreLider").TitleObject.Caption = "Nombre líder";
                grid.Columns.Item("MontoLider").Editable = true;
                grid.Columns.Item("MontoLider").TitleObject.Caption = "Monto líder";

                grid.Columns.Item("CodigoSupervisor").Editable = true;
                grid.Columns.Item("CodigoSupervisor").TitleObject.Caption = "Código supervisor";
                grid.Columns.Item("NombreSupervisor").Editable = true;
                grid.Columns.Item("NombreSupervisor").TitleObject.Caption = "Nombre supervisor";
                grid.Columns.Item("MontoSupervisor").Editable = true;
                grid.Columns.Item("MontoSupervisor").TitleObject.Caption = "Monto supervisor";

                grid.Columns.Item("CodigoCoordinador").Editable = true;
                grid.Columns.Item("CodigoCoordinador").TitleObject.Caption = "Código coordinador";
                grid.Columns.Item("NombreCoordinador").Editable = true;
                grid.Columns.Item("NombreCoordinador").TitleObject.Caption = "Nombre coordinador";
                grid.Columns.Item("MontoCoordinador").Editable = true;
                grid.Columns.Item("MontoCoordinador").TitleObject.Caption = "Monto coordinador";
                grid.Columns.Item("MontoBonoCoordi").Editable = true;
                grid.Columns.Item("MontoBonoCoordi").TitleObject.Caption = "Bono coordinador";

                grid.Columns.Item("CodigoCoordinador2").Editable = true;
                grid.Columns.Item("CodigoCoordinador2").TitleObject.Caption = "Código coordinador 2";
                grid.Columns.Item("NombreCoordinador2").Editable = true;
                grid.Columns.Item("NombreCoordinador2").TitleObject.Caption = "Nombre coordinador 2";
                grid.Columns.Item("MontoCoordinador2").Editable = true;
                grid.Columns.Item("MontoCoordinador2").TitleObject.Caption = "Monto coordinador 2";
                grid.Columns.Item("MontoBonoCoordi2").Editable = true;
                grid.Columns.Item("MontoBonoCoordi2").TitleObject.Caption = "Bono coordinador 2";

                grid.Columns.Item("CodigoGerente").Editable = true;
                grid.Columns.Item("CodigoGerente").TitleObject.Caption = "Código gerente";
                grid.Columns.Item("NombreGerente").Editable = true;
                grid.Columns.Item("NombreGerente").TitleObject.Caption = "Nombre gerente";
                grid.Columns.Item("MontoGerente").Editable = true;
                grid.Columns.Item("MontoGerente").TitleObject.Caption = "Monto gerente";

                grid.Columns.Item("MontoFideicomiso").Editable = false;
                grid.Columns.Item("MontoFideicomiso").TitleObject.Caption = "Fideicomiso";

                grid.Columns.Item("FechaCreacion").Editable = false;
                grid.Columns.Item("FechaCreacion").TitleObject.Caption = "Fecha creación";

                grid.Columns.Item("HisRecomendado").Editable = true;
                grid.Columns.Item("HisRecomendado").TitleObject.Caption = "Histórico Recomendado";

                grid.Columns.Item("HisAsistente").Editable = true;
                grid.Columns.Item("HisAsistente").TitleObject.Caption = "Histórico Asistente";

                grid.Columns.Item("HisBonoAsistente").Editable = true;
                grid.Columns.Item("HisBonoAsistente").TitleObject.Caption = "Histórico bono Asistente";

                grid.Columns.Item("HisLider").Editable = true;
                grid.Columns.Item("HisLider").TitleObject.Caption = "Histórico Líder";

                grid.Columns.Item("HisSupervisor").Editable = true;
                grid.Columns.Item("HisSupervisor").TitleObject.Caption = "Histórico Supervisor";

                grid.Columns.Item("HisCoordinador").Editable = true;
                grid.Columns.Item("HisCoordinador").TitleObject.Caption = "Histórico Coordinador";

                grid.Columns.Item("HisBonoCoordinador").Editable = true;
                grid.Columns.Item("HisBonoCoordinador").TitleObject.Caption = "Histórico bono Coordinador";

                grid.Columns.Item("HisCoordinador2").Editable = true;
                grid.Columns.Item("HisCoordinador2").TitleObject.Caption = "Histórico Coordinador 2";

                grid.Columns.Item("HisBonoCoordinador2").Editable = true;
                grid.Columns.Item("HisBonoCoordinador2").TitleObject.Caption = "Histórico bono Coordinador 2";

                grid.Columns.Item("HisGerente").Editable = true;
                grid.Columns.Item("HisGerente").TitleObject.Caption = "Histórico Gerente";



                RealizarCalculoFideicomiso(grid, "");
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
                string contrato = null;
                string codigoCobrador = null;
                string nombreCobrador = null;
                string codigoRecomendado = null;
                string nombreRecomendado = null;
                double montoRecomendado = 0;
                string codigoAsistente = null;
                string nombreAsistente = null;
                double montoAsistente = 0;
                double bonoAsistente = 0;
                string codigoLider = null;
                string nombreLider = null;
                double montoLider = 0;
                string codigoSupervisor = null;
                string nombreSupervisor = null;
                double montoSupervisor = 0;
                string codigoCoordinador = null;
                string nombreCoordinador = null;
                double montoCoordinador = 0;
                double montoBonoCoordinador = 0;
                string codigoCoordinador2 = null;
                string nombreCoordinador2 = null;
                double montoCoordinador2 = 0;
                double montoBonoCoordinador2 = 0;
                string codigoGerente = null;
                string nombreGerente = null;
                double montoGerente = 0;
                double montoFideicomiso = 0;
                double historicoRecomendado = 0;
                double historicoAsistente = 0;
                double historicoBonoAsistente = 0;
                double historicoLider = 0;
                double historicoSupervisor = 0;
                double historicoCoordinador = 0;
                double historicoBonoCoordinador = 0;
                double historicoCoordinador2 = 0;
                double historicoBonoCoordinador2 = 0;
                double historicoGerente = 0;

                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _Application.StatusBar.SetText("Guardando datos por favor espere...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                while (_oGridCalculoComisiones.Rows.Count - 1 >= i)
                {
                    contrato = _oGridCalculoComisiones.DataTable.GetValue("NumContrato", i);
                    codigoCobrador = _oGridCalculoComisiones.DataTable.GetValue("CodigoCobrador", i).ToString();
                    nombreCobrador = _oGridCalculoComisiones.DataTable.GetValue("NombreCobrador", i).ToString();
                    codigoRecomendado = _oGridCalculoComisiones.DataTable.GetValue("CodigoRecomendado", i).ToString();
                    nombreRecomendado = _oGridCalculoComisiones.DataTable.GetValue("NombreRecomendado", i).ToString();
                    montoRecomendado = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoRecomendado", i).ToString());
                    codigoAsistente = _oGridCalculoComisiones.DataTable.GetValue("CodigoAsistente", i).ToString();
                    nombreAsistente = _oGridCalculoComisiones.DataTable.GetValue("NombreAsistente", i).ToString();
                    montoAsistente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoAsistente", i));
                    bonoAsistente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoBonoAsistente", i));
                    codigoLider = _oGridCalculoComisiones.DataTable.GetValue("CodigoLider", i).ToString();
                    nombreLider = _oGridCalculoComisiones.DataTable.GetValue("NombreLider", i);
                    montoLider = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoLider", i).ToString());
                    codigoSupervisor = _oGridCalculoComisiones.DataTable.GetValue("CodigoSupervisor", i).ToString();
                    nombreSupervisor = _oGridCalculoComisiones.DataTable.GetValue("NombreSupervisor", i);
                    montoSupervisor = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoSupervisor", i));
                    codigoCoordinador = _oGridCalculoComisiones.DataTable.GetValue("CodigoCoordinador", i).ToString();
                    nombreCoordinador = _oGridCalculoComisiones.DataTable.GetValue("NombreCoordinador", i).ToString();
                    montoCoordinador = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoCoordinador", i));
                    montoBonoCoordinador = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoBonoCoordi", i));
                    codigoCoordinador2 = _oGridCalculoComisiones.DataTable.GetValue("CodigoCoordinador2", i).ToString();
                    nombreCoordinador2 = _oGridCalculoComisiones.DataTable.GetValue("NombreCoordinador2", i).ToString();
                    montoCoordinador2 = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoCoordinador2", i));
                    montoBonoCoordinador2 = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoBonoCoordi2", i));
                    codigoGerente = _oGridCalculoComisiones.DataTable.GetValue("CodigoGerente", i).ToString();
                    nombreGerente = _oGridCalculoComisiones.DataTable.GetValue("NombreGerente", i).ToString();
                    montoGerente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoGerente", i));
                    montoFideicomiso = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoFideicomiso", i));

                    historicoRecomendado = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("HisRecomendado", i));
                    historicoAsistente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("HisAsistente", i));
                    historicoBonoAsistente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("HisBonoAsistente", i));
                    historicoLider = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("HisLider", i));
                    historicoSupervisor = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("HisSupervisor", i));
                    historicoCoordinador = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("HisCoordinador", i));
                    historicoBonoCoordinador = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("HisBonoCoordinador", i));
                    historicoCoordinador2 = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("HisCoordinador2", i));
                    historicoBonoCoordinador2 = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("HisBonoCoordinador2", i));
                    historicoGerente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("HisGerente", i));

                    _oRec.DoQuery(@"UPDATE  dbo.[@CALCULO_COMISIONES]
                                    SET     U_Codigo_Cobrador = '" + codigoCobrador + "', " +
                                            "U_Nombre_Cobrador = '" + nombreCobrador + "', " +
                                            "U_CodigoRecomendado = '" + codigoRecomendado + "', " +
                                            "U_Nom_Recomendado = '" + nombreRecomendado + "', " +
                                            "U_Recomendado = '" + montoRecomendado + "', " +
                                            "U_Codigo_Asistente = '" + codigoAsistente + "', " +
                                            "U_Asistente_Social = '" + nombreAsistente + "', " +
                                            "U_Asis_Social = '" + montoAsistente + "', " +
                                            "U_BonoAsistente = '" + bonoAsistente + "', " +
                                            "U_CodigoLider = '" + codigoLider + "', " +
                                            "U_Nom_Lider = '" + nombreLider + "', " +
                                            "U_Lider = '" + montoLider + "', " +
                                            "U_CodigoSupervisor = '" + codigoSupervisor + "', " +
                                            "U_Nom_Supervisor = '" + nombreSupervisor + "', " +
                                            "U_Supervisor = '" + montoSupervisor + "', " +
                                            "U_CodigoCoordinador = '" + codigoCoordinador + "', " +
                                            "U_Nom_Coordinador = '" + nombreCoordinador + "', " +
                                            "U_Coordinador = '" + montoCoordinador + "', " +
                                            "U_BonoCoordinador = '" + montoBonoCoordinador + "', " +
                                            "U_CodigoCoordinador2 = '" + codigoCoordinador2 + "', " +
                                            "U_Nom_Coordinador2 = '" + nombreCoordinador2 + "', " +
                                            "U_Coordinador2 = '" + montoCoordinador2 + "', " +
                                            "U_BonoCoordinador2 = '" + montoBonoCoordinador2 + "', " +
                                            "U_CodigoGerente = '" + codigoGerente + "', " +
                                            "U_Nom_Gerente = '" + nombreGerente + "', " +
                                            "U_Gerente = '" + montoGerente + "', " +
                                            "U_Fideicomiso = '" + montoFideicomiso + "', " +
                                            "U_His_Recomendado = '" + historicoRecomendado + "', " +
                                            "U_His_Asis_Social = '" + historicoAsistente + "', " +
                                            "U_His_BonoAsis = '" + historicoBonoAsistente + "', " +
                                            "U_His_Lider = '" + historicoLider + "', " +
                                            "U_His_Supervisor = '" + historicoSupervisor + "', " +
                                            "U_His_Coordinador = '" + historicoCoordinador + "', " +
                                            "U_His_BonoCoord = '" + historicoBonoCoordinador + "', " +
                                            "U_His_Coordinador2 = '" + historicoCoordinador2 + "', " +
                                            "U_His_BonoCoord2 = '" + historicoBonoCoordinador2 + "', " +
                                            "U_His_Gerente = '" + historicoGerente + "' " +
                                    "WHERE   U_Contrato = '" + contrato + "'");
                    i += 1;
                }
                _Application.StatusBar.SetText("Datos guardados correctamente...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                BuscarDatos(_oTxtContrato.Value.ToString());
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
        private void RealizarCalculoFideicomiso(Grid _oGridCalculoComisiones, string lineaAfectada)
        {
            try
            {
                _oForma.Freeze(true);
                int i = 0;
                string cardCode = null;
                double InversionIni = 0;
                double Recomendado = 0;
                double Asistente = 0;
                double bonoAsistente = 0;
                double Lider = 0;
                double Supervisor = 0;
                double Coordinador = 0;
                double BonoCoordinador = 0;
                double Coordinador2 = 0;
                double BonoCoordinador2 = 0;
                double Gerente = 0;
                double fideicomiso = 0;
                double saldoDeCuenta = 0;

                if (!string.IsNullOrEmpty(lineaAfectada))
                {
                    cardCode = _oGridCalculoComisiones.DataTable.GetValue("CardCode", Convert.ToInt32(lineaAfectada));
                    saldoDeCuenta = Extensor.ObtenerSaldoDeCuenta(cardCode, _Company);
                    InversionIni = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("InversionInicial", Convert.ToInt32(lineaAfectada)));
                    Recomendado = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoRecomendado", Convert.ToInt32(lineaAfectada)));
                    Asistente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoAsistente", Convert.ToInt32(lineaAfectada)));
                    bonoAsistente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoBonoAsistente", Convert.ToInt32(lineaAfectada)));
                    //Bono = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoBonoCliente", Convert.ToInt32(lineaAfectada)));
                    Lider = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoLider", Convert.ToInt32(lineaAfectada)));
                    Supervisor = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoSupervisor", Convert.ToInt32(lineaAfectada)));
                    Coordinador = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoCoordinador", Convert.ToInt32(lineaAfectada)));
                    BonoCoordinador = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoBonoCoordi", Convert.ToInt32(lineaAfectada)));
                    Coordinador2 = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoCoordinador2", Convert.ToInt32(lineaAfectada)));
                    BonoCoordinador2 = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoBonoCoordi2", Convert.ToInt32(lineaAfectada)));
                    Gerente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoGerente", Convert.ToInt32(lineaAfectada)));
                    fideicomiso = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoFideicomiso", Convert.ToInt32(lineaAfectada)));
                    if (saldoDeCuenta > 0)
                    {
                        fideicomiso = saldoDeCuenta - (InversionIni + Recomendado + Asistente + bonoAsistente + Lider + Supervisor + Coordinador + BonoCoordinador + Coordinador2 + BonoCoordinador2 + Gerente);
                        ((SAPbouiCOM.EditTextColumn)_oGridCalculoComisiones.Columns.Item("MontoFideicomiso")).SetText(Convert.ToInt32(lineaAfectada), fideicomiso.ToString());
                    }


                    if (fideicomiso == 0)
                    {
                        _oGridCalculoComisiones.CommonSetting.SetRowFontColor(Convert.ToInt32(lineaAfectada) + 1, Color.Red.ToArgb());
                    }
                    else if (fideicomiso < 0)
                    {
                        _oGridCalculoComisiones.CommonSetting.SetRowFontColor(Convert.ToInt32(lineaAfectada) + 1, Color.Blue.ToArgb());
                    }
                    else
                    {
                        _oGridCalculoComisiones.CommonSetting.SetRowFontColor(Convert.ToInt32(lineaAfectada) + 1, Color.Black.ToArgb());
                    }

                }
                else
                {
                    while (_oGridCalculoComisiones.Rows.Count - 1 >= i)
                    {
                        cardCode = _oGridCalculoComisiones.DataTable.GetValue("CardCode", i);
                        saldoDeCuenta = Extensor.ObtenerSaldoDeCuenta(cardCode, _Company);
                        InversionIni = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("InversionInicial", i));
                        Recomendado = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoRecomendado", i));
                        Asistente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoAsistente", i));
                        bonoAsistente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoBonoAsistente", i));
                        //Bono = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoBonoCliente", i));
                        Lider = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoLider", i));
                        Supervisor = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoSupervisor", i));
                        Coordinador = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoCoordinador", i));
                        BonoCoordinador = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoBonoCoordi", i));
                        Coordinador2 = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoCoordinador2", i));
                        BonoCoordinador2 = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoBonoCoordi2", i));
                        Gerente = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoGerente", i));
                        fideicomiso = Convert.ToDouble(_oGridCalculoComisiones.DataTable.GetValue("MontoFideicomiso", i));
                        if (saldoDeCuenta > 0)
                        {
                            fideicomiso = saldoDeCuenta - (InversionIni + Recomendado + Asistente + bonoAsistente + Lider + Supervisor + Coordinador + BonoCoordinador + Coordinador2 + BonoCoordinador2 + Gerente);
                            _oGridCalculoComisiones.DataTable.SetValue("MontoFideicomiso", i, fideicomiso);
                        }


                        if (fideicomiso == 0)
                        {
                            _oGridCalculoComisiones.CommonSetting.SetRowFontColor(i + 1, Color.Red.ToArgb());
                        }
                        else if (fideicomiso < 0)
                        {
                            _oGridCalculoComisiones.CommonSetting.SetRowFontColor(i + 1, Color.Blue.ToArgb());
                        }
                        else
                        {
                            _oGridCalculoComisiones.CommonSetting.SetRowFontColor(i + 1, Color.Black.ToArgb());
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
    }
}
