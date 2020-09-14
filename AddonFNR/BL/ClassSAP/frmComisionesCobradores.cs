using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddonFNR.BL
{
    class frmComisionesCobradores : ComportaForm
    {
        #region CONSTANTES

        private const string FRM_COMISIONES_COBRADORES = "frmComisionesCobradores";
        private const string TXTCODCOB = "txtCodCob";
        private const string TXTNOMCOB = "txtNomCob";
        private const string BTN_BUSCAR = "btnBuscar";
        private const string BTN_UPDATE = "btnUpdate";
        private const string BTN_CANCELAR = "btnCancel";
        private const string BTN_SAP_BUSCAR = "1281";
        private const string BTN_SAP_CREAR = "1282";
        private const string DT_COMISIONES_COBRADORES = "dtComisionesCobradores";
        private const string LINK_DM_COBRADORES = "171";
        private const string GRD_COMISIONES_COBRADORES = "grdComCob";
        private const string COL_SUELDO_BASE = "SueldoBase";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForma = null;
        private static bool _oComisionesCobradores = false;
        private SAPbobsCOM.Recordset _oRec = null;
        private SAPbouiCOM.Grid _oGridComisionesCobradores = null;
        private SAPbouiCOM.EditText _oTxtCodeCobrador = null;
        private SAPbouiCOM.EditText _oTxtNomCobrador = null;
        private SAPbouiCOM.Button _oBotonUpdate = null;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de Comisiones cobradores
        /// </summary>
        /// <param name="_Application">Este es el objeto raíz de la API de interfaz de usuario
        ///                             lo que refleja la cual aplicación SAP Business One en el que se realiza 
        ///                             la conexión</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        public frmComisionesCobradores(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oComisionesCobradores == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                showForm(formID);
                inicializarComponentes();
                setEventos();
                _oComisionesCobradores = true;
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
                if (_oComisionesCobradores != false && pVal.FormType != FormTypeMenu && formID == FormUID)
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
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && (pVal.ItemUID == GRD_COMISIONES_COBRADORES))
                    {
                        _oBotonUpdate.Caption = "Actualizar";
                    }
                }

                if (pVal.FormUID == formID && pVal.BeforeAction == true)
                {
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
                                _oComisionesCobradores = false;
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
                                _oComisionesCobradores = false;
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
                            _oComisionesCobradores = false;
                            Addon.typeList.RemoveAll(p => p._forma == formID);
                            return;
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if (pVal.ItemUID == BTN_BUSCAR)
                        {
                            _oTxtCodeCobrador = _oForma.Items.Item(TXTCODCOB).Specific;
                            _oTxtNomCobrador = _oForma.Items.Item(TXTNOMCOB).Specific;
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
                                ActualizarInformacion();
                                _oBotonUpdate.Caption = "Ok";
                            }
                        }

                        if(pVal.ItemUID == GRD_COMISIONES_COBRADORES && pVal.ColUID == COL_SUELDO_BASE)
                        {
                            _oBotonUpdate.Caption = "Actualizar";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en evento *clsComisionCobradores* : " + ex.Message);
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
                        if (_Application.Forms.ActiveForm.UniqueID == FRM_COMISIONES_COBRADORES)
                            BubbleEvent = false;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en ManuEvent *clsComisionCobradores* : " + ex.Message);
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
                _oForma.DataSources.DataTables.Add(DT_COMISIONES_COBRADORES);
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
                                              "U_Cobrador),'á','a') ), " +
                                                       "'é','e') ), " +
                                                       "'í','i') ), " +
                                                       "'ó','o') ), " +
                                                       "'ú','u') LIKE  '%" + nameCobrador.ToLower().Replace('á', 'a').
                                                                                                   Replace('é', 'e').
                                                                                                   Replace('í', 'i').
                                                                                                   Replace('ó', 'o').
                                                                                                   Replace('ú', 'u') + "%' ";

                _oGridComisionesCobradores = _oForma.Items.Item(GRD_COMISIONES_COBRADORES).Specific;
                _oGridComisionesCobradores.DataTable = _oForma.DataSources.DataTables.Item(DT_COMISIONES_COBRADORES);

                _oGridComisionesCobradores.DataTable.ExecuteQuery(@"SELECT  T1.empID AS CodigoID,
                                                                            T0.U_Codigo_Cobrador AS CodCobrador,
                                                                            T0.U_Cobrador AS NomCobrador,
                                                                            T0.U_Serie_Programa AS SeriePrograma,
                                                                            --T0.U_Serie_Malba AS SerieMalba,
                                                                            T0.U_Serie_Cooperativa AS SerieCooperativa,
                                                                            T0.U_Serie_Panteon AS SeriePanteon,
                                                                            ISNULL(T0.U_Comision_Panteon, 0.00) AS ComisionPanteon,
                                                                            ISNULL(T0.U_Sueldo_Cooperativa, 0.00) AS SueldoCooperativa,
                                                                            ISNULL(T0.U_Sueldo_Apoyo, 0.00) AS SueldoApoyo,
                                                                            ISNULL(T0.U_Comision_Coop, 0.00) AS ComisionCooperativa,
                                                                            ISNULL(T0.U_Comision_Apoyo, 0.00) AS ComisionApoyo,
                                                                            ISNULL(T0.U_Efectividad, 0.00) AS Efectividad,
                                                                            T0.U_SueldoBase AS SueldoBase
                                                                    FROM    dbo.[@COMISION_COBRADORES] T0
                                                                            LEFT JOIN dbo.OHEM T1 ON T1.firstName = T0.U_Codigo_Cobrador
                                                                                                     --AND ( T1.middleName + ' ' + T1.lastName ) = T0.U_Cobrador
                                                                WHERE   --( T1.middleName + ' ' + T1.lastName ) = T0.U_Cobrador
                                                                       -- AND 
                                                                        T0.U_Codigo_Cobrador LIKE '%" + codeCobrador + "%' " + qryCobrador +
                                                                "ORDER BY T0.U_Codigo_Cobrador, " +
                                                                        "T1.empID ASC ");
                if (_oGridComisionesCobradores.DataTable.IsEmpty)
                {
                    _Application.MessageBox("No se encontraron registros");
                    _oGridComisionesCobradores.DataTable.Clear();
                    _oForma.Update();
                }
                else
                {
                    //_oGridComisionesCobradores.AutoResizeColumns();                    
                    FormatoGrid(_oGridComisionesCobradores);
                    _oGridComisionesCobradores.AutoResizeColumns();  
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

                grid.Columns.Item("CodCobrador").Editable = false;
                grid.Columns.Item("CodCobrador").TitleObject.Caption = "Código";

                grid.Columns.Item("NomCobrador").Editable = false;
                grid.Columns.Item("NomCobrador").TitleObject.Caption = "Cobrador";

                grid.Columns.Item("SeriePrograma").Editable = true;
                grid.Columns.Item("SeriePrograma").TitleObject.Caption = "Serie programa";

                //grid.Columns.Item("SerieMalba").Editable = true;
                //grid.Columns.Item("SerieMalba").TitleObject.Caption = "Serie malba";

                grid.Columns.Item("SerieCooperativa").Editable = true;
                grid.Columns.Item("SerieCooperativa").TitleObject.Caption = "Serie cooperativa";

                grid.Columns.Item("SeriePanteon").Editable = true;
                grid.Columns.Item("SeriePanteon").TitleObject.Caption = "Serie panteón";

                grid.Columns.Item("ComisionPanteon").Editable = true;
                grid.Columns.Item("ComisionPanteon").TitleObject.Caption = "Comisión panteón %";

                grid.Columns.Item("SueldoCooperativa").Editable = true;
                grid.Columns.Item("SueldoCooperativa").TitleObject.Caption = "Sueldo cooperativa %";

                grid.Columns.Item("SueldoApoyo").Editable = true;
                grid.Columns.Item("SueldoApoyo").TitleObject.Caption = "Sueldo apoyo %";

                grid.Columns.Item("ComisionCooperativa").Editable = true;
                grid.Columns.Item("ComisionCooperativa").TitleObject.Caption = "Comisión cooperativa %";

                grid.Columns.Item("ComisionApoyo").Editable = true;
                grid.Columns.Item("ComisionApoyo").TitleObject.Caption = "Comisión apoyo %";

                grid.Columns.Item("Efectividad").Editable = true;
                grid.Columns.Item("Efectividad").TitleObject.Caption = "Efectividad %";

                grid.Columns.Item("SueldoBase").Editable = true;
                grid.Columns.Item("SueldoBase").TitleObject.Caption = "Sueldo base";

                if (!grid.DataTable.IsEmpty)
                {
                    grid.Columns.Item("SueldoBase").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                }
                                
                //Agregar filtros a columnas del grid                   
                grid.Columns.Item("SeriePrograma").TitleObject.Sortable = true;
                grid.Columns.Item("SeriePrograma").TitleObject.Sort(BoGridSortType.gst_Ascending);
                //grid.Columns.Item("SerieMalba").TitleObject.Sortable = true;
                //grid.Columns.Item("SerieMalba").TitleObject.Sort(BoGridSortType.gst_Ascending);
                grid.Columns.Item("SerieCooperativa").TitleObject.Sortable = true;
                grid.Columns.Item("SerieCooperativa").TitleObject.Sort(BoGridSortType.gst_Ascending);
                grid.Columns.Item("SeriePanteon").TitleObject.Sortable = true;
                grid.Columns.Item("SeriePanteon").TitleObject.Sort(BoGridSortType.gst_Ascending);
                grid.Columns.Item("ComisionPanteon").TitleObject.Sortable = true;
                grid.Columns.Item("ComisionPanteon").TitleObject.Sort(BoGridSortType.gst_Ascending);
                grid.Columns.Item("SueldoCooperativa").TitleObject.Sortable = true;
                grid.Columns.Item("SueldoCooperativa").TitleObject.Sort(BoGridSortType.gst_Ascending);
                grid.Columns.Item("SueldoApoyo").TitleObject.Sortable = true;
                grid.Columns.Item("SueldoApoyo").TitleObject.Sort(BoGridSortType.gst_Ascending);
                grid.Columns.Item("ComisionCooperativa").TitleObject.Sortable = true;
                grid.Columns.Item("ComisionCooperativa").TitleObject.Sort(BoGridSortType.gst_Ascending);
                grid.Columns.Item("ComisionApoyo").TitleObject.Sortable = true;
                grid.Columns.Item("ComisionApoyo").TitleObject.Sort(BoGridSortType.gst_Ascending);
                grid.Columns.Item("Efectividad").TitleObject.Sortable = true;
                grid.Columns.Item("Efectividad").TitleObject.Sort(BoGridSortType.gst_Ascending);
                grid.Columns.Item("CodCobrador").TitleObject.Sortable = true;
                grid.Columns.Item("CodCobrador").TitleObject.Sort(BoGridSortType.gst_Ascending);
                grid.Columns.Item("NomCobrador").TitleObject.Sortable = true;
                grid.Columns.Item("NomCobrador").TitleObject.Sort(BoGridSortType.gst_Ascending);
                grid.Columns.Item("CodigoID").TitleObject.Sortable = true;
                grid.Columns.Item("CodigoID").TitleObject.Sort(BoGridSortType.gst_Ascending);
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
                string CodCobrador = null;
                string NomCobrador = null;
                string SeriePrograma = null;
                //string SerieMalba = null;
                string SerieCooperativa = null;
                string SeriePanteon = null;
                double comisionPanteon = 0;
                double SueldoCooperativa = 0;
                double SueldoApoyo = 0;
                double ComisionCooperativa = 0;
                double ComisionApoyo = 0;
                double Efectividad = 0;
                string StatusSueldoBase = null;

                // Variable para confirmar que se termino correctgamente
                bool confirmacion = true;

                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _Application.StatusBar.SetText("Guardando datos por favor espere...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                while (_oGridComisionesCobradores.Rows.Count - 1 >= i)
                {
                    CodigoID = Convert.ToInt32(_oGridComisionesCobradores.DataTable.GetValue("CodigoID", i));
                    CodCobrador = _oGridComisionesCobradores.DataTable.GetValue("CodCobrador", i).ToString();
                    NomCobrador = _oGridComisionesCobradores.DataTable.GetValue("NomCobrador", i).ToString();
                    SeriePrograma = _oGridComisionesCobradores.DataTable.GetValue("SeriePrograma", i).ToString();
                    //SerieMalba = _oGridComisionesCobradores.DataTable.GetValue("SerieMalba", i).ToString();
                    SerieCooperativa = _oGridComisionesCobradores.DataTable.GetValue("SerieCooperativa", i).ToString();
                    SeriePanteon = _oGridComisionesCobradores.DataTable.GetValue("SeriePanteon", i).ToString();
                    comisionPanteon = Convert.ToDouble(_oGridComisionesCobradores.DataTable.GetValue("ComisionPanteon", i));
                    SueldoCooperativa = Convert.ToDouble(_oGridComisionesCobradores.DataTable.GetValue("SueldoCooperativa", i));
                    SueldoApoyo = Convert.ToDouble(_oGridComisionesCobradores.DataTable.GetValue("SueldoApoyo", i));
                    ComisionCooperativa = Convert.ToDouble(_oGridComisionesCobradores.DataTable.GetValue("ComisionCooperativa", i));
                    ComisionApoyo = Convert.ToDouble(_oGridComisionesCobradores.DataTable.GetValue("ComisionApoyo", i));
                    Efectividad = Convert.ToDouble(_oGridComisionesCobradores.DataTable.GetValue("Efectividad", i));
                    StatusSueldoBase = _oGridComisionesCobradores.DataTable.GetValue("SueldoBase", i).ToString();

                    if (SueldoApoyo <= 20 && SueldoCooperativa <= 20 && ComisionCooperativa <= 20 && ComisionApoyo <= 20 && comisionPanteon <= 20)
                    {
                        _oRec.DoQuery(@"UPDATE  T0
                                    SET     T0.U_Serie_Programa = '" + SeriePrograma + "', " +
                                            "T0.U_Serie_Cooperativa = '" + SerieCooperativa + "', " +
                                            "T0.U_Serie_Panteon = '" + SeriePanteon + "', " +
                                            "T0.U_Comision_Panteon = '" + comisionPanteon + "', " +
                                            "T0.U_Sueldo_Cooperativa = '" + SueldoCooperativa + "', " +
                                            "T0.U_Sueldo_Apoyo = '" + SueldoApoyo + "', " +
                                            "T0.U_Comision_Coop = '" + ComisionCooperativa + "', " +
                                            "T0.U_Comision_Apoyo = '" + ComisionApoyo + "', " +
                                            "T0.U_Efectividad = '" + Efectividad + "', " +
                                            "T0.U_SueldoBase = '" + StatusSueldoBase + "' " +
                                    "FROM    dbo.[@COMISION_COBRADORES] T0 " +
                                            "LEFT JOIN dbo.OHEM T1 ON T1.firstName = T0.U_Codigo_Cobrador " +
                                                                     "AND ( T1.middleName + ' ' + T1.lastName ) = T0.U_Cobrador " +
                                    "WHERE   T1.empID = '" + CodigoID + "' " +
                                            "AND T0.U_Codigo_Cobrador = '" + CodCobrador + "' " +
                                            "AND T0.U_Cobrador = '" + NomCobrador + "' ");
                        i += 1;
                    }
                    else
                    {
                        _Application.StatusBar.SetText("No se pueden guardar porcentajes que superen el 20%", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        confirmacion = false;

                        i += 1;
                    }
                }

                if (confirmacion == true)
                {
                    _Application.StatusBar.SetText("Datos guardados correctamente...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    BuscarDatos(_oTxtCodeCobrador.Value.ToString(), _oTxtNomCobrador.Value.ToString());
                }
                
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

        #endregion
    }
}
