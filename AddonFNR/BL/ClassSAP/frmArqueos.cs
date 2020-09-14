using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddonFNR.BL
{
    class frmArqueos : ComportaForm
    {
        #region CONSTANTES

        private const string FRM_ARQUEOS = "frmArqueos";
        private const string TXT_CODIGO_ASOCIADO = "txtCodAso";
        private const string TXT_NOMBRE_ASOCIADO = "txtNomAso";
        private const string TXT_CODIGO_OFICINA = "txtCodOfi";
        private const string TXT_NOMBRE_OFICINA = "txtNomOfi";
        private const string TXT_CODIGO_SOLICITUD = "txtCodSol";
        private const string GRD_ARQUEOS = "grdArqueo";
        private const string BTN_CERRAR = "btnCerrar";
        private const string BTN_ARQUEO = "btnArqueo";
        private const string BTN_IMPRIMIR_ARQUEO = "btnImpArq";
        private const string BTN_SAP_BUSCAR = "1281";
        private const string BTN_SAP_CREAR = "1282";
        private const string DT_ARQUEOS = "dtArqueos";
        private const int CHAR_PRESS_ENTER = 13;
        private const int CHAR_PRESS_TAB = 9;

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForma = null;
        private static bool _oArqueos = false;
        private SAPbobsCOM.Recordset _oRec = null;
        private SAPbouiCOM.Grid _oGridArqueos = null;
        private SAPbouiCOM.EditText _oTxtCodigoAsociado = null;
        private SAPbouiCOM.EditText _oTxtNombreAsociado = null;
        private SAPbouiCOM.EditText _oTxtCodigoOficina = null;
        private SAPbouiCOM.EditText _oTxtNombreOficina = null;
        private SAPbouiCOM.EditText _oTxtCodigoSolicitud = null;
        private SAPbouiCOM.Button _oBtnImprimirArqueo = null;
        bool verificar = false;
        DateTime TextoInicial;
        DateTime TextoFinal;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de Arqueos
        /// </summary>
        /// <param name="_Application">Este es el objeto raíz de la API de interfaz de usuario
        ///                             lo que refleja la cual aplicación SAP Business One en el que se realiza 
        ///                             la conexión</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        public frmArqueos(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oArqueos == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                showForm(formID);
                inicializarComponentes();
                setEventos();
                _oArqueos = true;
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
                if (_oArqueos != false && pVal.FormType != FormTypeMenu && formID == FormUID)
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
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
                    {
                        _Application.ItemEvent -= new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                        _Application.MenuEvent -= new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                        Dispose();
                        application = null;
                        company = null;
                        _oArqueos = false;
                        Addon.typeList.RemoveAll(p => p._forma == formID);
                        return;
                    }

                    if (pVal.EventType == BoEventTypes.et_KEY_DOWN && pVal.ItemUID == TXT_CODIGO_ASOCIADO && pVal.CharPressed == CHAR_PRESS_TAB)
                    {
                        if (!string.IsNullOrEmpty(_oTxtCodigoAsociado.Value.ToString()))
                        {
                            ObtenerInformacionAsociado();
                            if (ValidarSiYaEscaneo())
                            {
                                LlenarGridArqueo("AND U_FechaArqueo IS NULL");
                            }
                            else if (ValidarSiYaCerroArqueo())
                            {
                                LlenarGridArqueo(" AND CONVERT(VARCHAR(10), U_FechaArqueo, 103) = CONVERT(VARCHAR(10), GETDATE(), 103)");
                                _oBtnImprimirArqueo.Item.Visible = true;
                                _oTxtCodigoAsociado.Active = true;
                                _oTxtCodigoSolicitud.Item.Enabled = false;
                                _oForma.Items.Item(BTN_ARQUEO).Enabled = false;
                            }
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_KEY_DOWN && pVal.ItemUID == TXT_CODIGO_SOLICITUD && pVal.CharPressed != CHAR_PRESS_TAB)
                    {
                        if (!string.IsNullOrEmpty(_oTxtCodigoAsociado.Value.ToString()))
                        {
                            ValidarEscaneoSerie();
                        }
                        else
                        {
                            _oTxtCodigoSolicitud.Value = "";
                            _Application.StatusBar.SetText("Capture el código del asociado", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_KEY_DOWN && pVal.CharPressed == CHAR_PRESS_ENTER && pVal.ItemUID == TXT_CODIGO_SOLICITUD)
                    {
                        if (_oTxtCodigoSolicitud.Value.ToString().Length > 6)
                        {
                            AgregarSolicitud();
                            _oTxtCodigoSolicitud.Value = "";
                        }
                    }
                }

                if (pVal.FormUID == formID && pVal.BeforeAction == true)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if(pVal.ItemUID == BTN_IMPRIMIR_ARQUEO)
                        {
                            CargarReporteArqueo();
                        }

                        if (pVal.ItemUID == BTN_CERRAR)
                        {
                            _oForma.Close();
                        }

                        if (pVal.ItemUID == BTN_ARQUEO)
                        {
                            if (!string.IsNullOrEmpty(_oTxtCodigoAsociado.Value.ToString()))
                            {
                                if (!_oGridArqueos.DataTable.IsEmpty)
                                {
                                    if (_Application.MessageBox("¿Desea cerrar el arqueo?", 2, "Si", "No") == 1)
                                    {
                                        CerrarAqueo();
                                    }
                                }
                                else
                                {
                                    if (_Application.MessageBox("¿Desea cerrar el arqueo sin solicitudes?", 2, "Si", "No") == 1)
                                    {
                                        CerrarAqueoSinSolicitudes();
                                    }
                                }
                            }
                            else
                            {
                                _Application.MessageBox("Capture el código del asociado");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en evento *clsArqueos* : " + ex.Message);
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
                        if (_Application.Forms.ActiveForm.UniqueID == FRM_ARQUEOS)
                            BubbleEvent = false;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en ManuEvent *clsArqueos* : " + ex.Message);
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

                //Código del asociado
                _oForma.DataSources.UserDataSources.Add(TXT_CODIGO_ASOCIADO, BoDataType.dt_SHORT_TEXT, 10);
                _oTxtCodigoAsociado = (SAPbouiCOM.EditText)_oForma.Items.Item(TXT_CODIGO_ASOCIADO).Specific;
                _oTxtCodigoAsociado.DataBind.SetBound(true, "", TXT_CODIGO_ASOCIADO);

                //Nombre del asociado
                _oForma.DataSources.UserDataSources.Add(TXT_NOMBRE_ASOCIADO, BoDataType.dt_SHORT_TEXT, 150);
                _oTxtNombreAsociado = (SAPbouiCOM.EditText)_oForma.Items.Item(TXT_NOMBRE_ASOCIADO).Specific;
                _oTxtNombreAsociado.DataBind.SetBound(true, "", TXT_NOMBRE_ASOCIADO);

                //Código de oficina
                _oForma.DataSources.UserDataSources.Add(TXT_CODIGO_OFICINA, BoDataType.dt_SHORT_TEXT, 10);
                _oTxtCodigoOficina = (SAPbouiCOM.EditText)_oForma.Items.Item(TXT_CODIGO_OFICINA).Specific;
                _oTxtCodigoOficina.DataBind.SetBound(true, "", TXT_CODIGO_OFICINA);

                //Nombre de oficina
                _oForma.DataSources.UserDataSources.Add(TXT_NOMBRE_OFICINA, BoDataType.dt_SHORT_TEXT, 150);
                _oTxtNombreOficina = (SAPbouiCOM.EditText)_oForma.Items.Item(TXT_NOMBRE_OFICINA).Specific;
                _oTxtNombreOficina.DataBind.SetBound(true, "", TXT_NOMBRE_OFICINA);

                //Código de solicitud
                _oForma.DataSources.UserDataSources.Add(TXT_CODIGO_SOLICITUD, BoDataType.dt_SHORT_TEXT, 20);
                _oTxtCodigoSolicitud = (SAPbouiCOM.EditText)_oForma.Items.Item(TXT_CODIGO_SOLICITUD).Specific;
                _oTxtCodigoSolicitud.DataBind.SetBound(true, "", TXT_CODIGO_SOLICITUD);

                _oBtnImprimirArqueo = _oForma.Items.Item(BTN_IMPRIMIR_ARQUEO).Specific;
                _oBtnImprimirArqueo.Item.Visible = false;

                _oTxtCodigoAsociado.Active = true;

                //Declarar DataTable
                _oForma.DataSources.DataTables.Add(DT_ARQUEOS);

            }
            catch (Exception ex)
            {
                _Application.MessageBox("Error al inicializar : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        /// <summary>
        /// Obtiene la información del asociado capturado
        /// </summary>
        private void ObtenerInformacionAsociado()
        {
            try
            {
                _oForma = _Application.Forms.Item(formID);
                _oForma.Freeze(true);
                string oficinaUsuario = null;
                string oficinaAsociado = null;
                string nombreOficinaAsociado = null;
                string nombreAsociado = null;
                _oRec = null;

                oficinaUsuario = Extensor.ObtenerAlmacenOficina(_Company);
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT T0.pager AS OficinaAsociado,
		                                            (T0.middleName + ' ' + T0.lastName) AS NombreAsociado,
		                                            T1.WhsName AS NombreOficina
                                             FROM   dbo.OHEM T0
                                             LEFT JOIN dbo.OWHS T1 ON T0.pager = T1.WhsCode
                                             WHERE  T0.firstName =  '" + _oTxtCodigoAsociado.Value + "'");

                oficinaAsociado = _oRec.Fields.Item("OficinaAsociado").Value.ToString();
                nombreOficinaAsociado = _oRec.Fields.Item("NombreOficina").Value.ToString();
                nombreAsociado = _oRec.Fields.Item("NombreAsociado").Value.ToString();

                if (!string.IsNullOrEmpty(oficinaAsociado))
                {
                    if (oficinaUsuario.Contains(oficinaAsociado))
                    {
                        _oTxtCodigoOficina.Value = oficinaAsociado;
                        _oTxtNombreOficina.Value = nombreOficinaAsociado;
                        _oTxtNombreAsociado.Value = nombreAsociado;
                        _oBtnImprimirArqueo.Item.Visible = false;
                        _oTxtCodigoSolicitud.Item.Enabled = true;
                        _oForma.Items.Item(BTN_ARQUEO).Enabled = true;
                        CargarGridArqueo();
                    }
                    else
                    {
                        _oTxtCodigoOficina.Value = "";
                        _oTxtNombreOficina.Value = "";
                        _oTxtNombreAsociado.Value = "";
                        _oTxtCodigoAsociado.Value = "";
                        _Application.MessageBox("No existe Asociado o no pertenece a tu oficina.");
                        if (_oGridArqueos != null)
                            _oGridArqueos.DataTable.Clear();
                    }
                }
                else
                {
                    _oTxtCodigoOficina.Value = "";
                    _oTxtNombreOficina.Value = "";
                    _oTxtNombreAsociado.Value = "";
                    _oTxtCodigoAsociado.Value = "";
                    _Application.MessageBox("No existe Asociado o no pertenece a tu oficina.");
                    if (_oGridArqueos != null)
                        _oGridArqueos.DataTable.Clear();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener información del asociado *ObtenerInformacionAsociado* : " + ex.Message);
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
        /// Valida que la captura de la serie sea solo por escáner
        /// </summary>
        private void ValidarEscaneoSerie()
        {
            try
            {
                int countText = _oTxtCodigoSolicitud.Value.ToString().Length;

                if (verificar == false && countText >= 1)
                {
                    TextoInicial = DateTime.Now;
                    verificar = true;
                }
                else
                {
                    if (countText >= 1)
                    {
                        TextoFinal = DateTime.Now;
                        double diferenciaSegundos = (TextoFinal - TextoInicial).TotalSeconds;

                        if (diferenciaSegundos >= 1)
                        {
                            _Application.MessageBox("Solo se debe escanear la solicitud");
                            _oTxtCodigoSolicitud.Value = "";
                            verificar = false;
                            return;
                        }
                    }
                }

                if (countText == 0)
                {
                    verificar = false;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar la entrada de escáner *ValidarEscaneoSerie* : " + ex.Message);
            }
        }

        /// <summary>
        /// Carga el grid de las series escaneadas
        /// </summary>
        private void CargarGridArqueo()
        {
            try
            {
                _oForma = _Application.Forms.Item(formID);
                _oForma.Freeze(true);//              

                _oForma.DataSources.DataTables.Item(DT_ARQUEOS)
                       .ExecuteQuery(@"SELECT  CAST('' AS INT) AS #,
		                                        CAST('' AS NVARCHAR(MAX)) AS CodigoAsistente ,
                                                CAST('' AS NVARCHAR(MAX)) AS NombreAsistente ,
                                                CAST('' AS NVARCHAR(MAX)) AS CodigoOficina ,
                                                CAST('' AS NVARCHAR(MAX)) AS Empresa ,
                                                CAST('' AS NVARCHAR(MAX)) AS 'Plan' ,
                                                CAST('' AS NVARCHAR(MAX)) AS Solicitud ,
                                                CAST('' AS DATETIME) AS FechaArqueo
                                                        ");

                _oGridArqueos = (SAPbouiCOM.Grid)_oForma.Items.Item(GRD_ARQUEOS).Specific;
                _oGridArqueos.DataTable = _oForma.DataSources.DataTables.Item(DT_ARQUEOS);
                FormatoGrid(_oGridArqueos,"");
                _oGridArqueos.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al cargar grid de arqueo *CargarGridArqueo* : " + ex.Message);
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
        public void FormatoGrid(Grid grid, string informacion)
        {
            try
            {
                _oForma = _Application.Forms.Item(formID);
                _oForma.Freeze(true);

                grid.Columns.Item("CodigoAsistente").Visible = false;
                grid.Columns.Item("NombreAsistente").Visible = false;
                grid.Columns.Item("CodigoOficina").Visible = false;
                grid.Columns.Item("FechaArqueo").Visible = false;

                grid.Columns.Item("#").Editable = false;
                grid.Columns.Item("#").TitleObject.Caption = "#";
                grid.Columns.Item("#").Type = BoGridColumnType.gct_EditText;

                grid.Columns.Item("Empresa").Editable = false;
                grid.Columns.Item("Empresa").TitleObject.Caption = "Empresa";
                grid.Columns.Item("Empresa").Type = BoGridColumnType.gct_EditText;

                grid.Columns.Item("Plan").Editable = false;
                grid.Columns.Item("Plan").TitleObject.Caption = "Plan";
                grid.Columns.Item("Plan").Type = BoGridColumnType.gct_EditText;

                grid.Columns.Item("Solicitud").Editable = false;
                grid.Columns.Item("Solicitud").TitleObject.Caption = "Solicitud";
                grid.Columns.Item("Solicitud").Type = BoGridColumnType.gct_EditText;

                if (string.IsNullOrEmpty(informacion))
                    grid.DataTable.Rows.Remove(0);
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
        /// Agrega la solicitud escaneada
        /// </summary>
        private void AgregarSolicitud()
        {
            try
            {
                _oForma.Freeze(true);
                string empresa = null;
                string plan = null;
                int code = 0;

                if (ValidarNoExistaSolicitud())
                {
                    _oRec = null;
                    _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    _oRec.DoQuery("SELECT U_Empresa,U_Descripcion_Plan FROM dbo.[@COMISIONES] WHERE U_Prefijo_Sol =  SUBSTRING('" + _oTxtCodigoSolicitud.Value.ToString() + "', 1,6)");
                    empresa = _oRec.Fields.Item("U_Empresa").Value;
                    plan = _oRec.Fields.Item("U_Descripcion_Plan").Value;

                    _oGridArqueos.DataTable.Rows.Add();
                    int lastRow = _oGridArqueos.DataTable.Rows.Count - 1;

                    _oGridArqueos.DataTable.SetValue("#", lastRow, _oGridArqueos.Rows.Count);
                    _oGridArqueos.DataTable.SetValue("CodigoAsistente", lastRow, _oTxtCodigoAsociado.Value.ToString());
                    _oGridArqueos.DataTable.SetValue("NombreAsistente", lastRow, _oTxtNombreAsociado.Value.ToString());
                    _oGridArqueos.DataTable.SetValue("CodigoOficina", lastRow, _oTxtCodigoOficina.Value.ToString());
                    //_oGridArqueos.DataTable.SetValue("FechaArqueo", lastRow, DateTime.Now);
                    _oGridArqueos.DataTable.SetValue("Empresa", lastRow, empresa);
                    _oGridArqueos.DataTable.SetValue("Plan", lastRow, plan);
                    _oGridArqueos.DataTable.SetValue("Solicitud", lastRow, _oTxtCodigoSolicitud.Value.ToString());
                    _oGridArqueos.AutoResizeColumns();
                    verificar = false;

                    _oRec = null;
                    _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    _oRec.DoQuery(@"SELECT  ISNULL(MAX(CONVERT(INT, Code)), 0) + 1 AS Code
                                    FROM    dbo.[@ARQUEOS]");
                    code = Convert.ToInt32(_oRec.Fields.Item("Code").Value);
                    _oRec = null;
                    _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    _oRec.DoQuery(@"INSERT INTO dbo.[@ARQUEOS]
                                                ( Code ,
                                                  Name ,
                                                  U_CodigoAsistente ,
                                                  U_NombreAsistente ,
                                                  U_CodigoOficina ,
                                                  U_Empresa ,
                                                  U_Plan ,
                                                  U_Solicitud
                                                )
                                        VALUES  ( '" + code + "' ," +
                                                  "'" + code + "' ," +
                                                  "'" + _oTxtCodigoAsociado.Value.ToString().ToUpper() + "' ," +
                                                  "'" + _oTxtNombreAsociado.Value.ToString() + "' ," +
                                                  "'" + _oTxtCodigoOficina.Value.ToString() + "' ," +
                                                  "'" + empresa + "' ," +
                                                  "'" + plan + "' ," +
                                                  "'" + _oTxtCodigoSolicitud.Value.ToString() + "'" +
                                                ")");

                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al agregar solicitud escaneada *AgregarSolicitud* : " + ex.Message);
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
        /// Valida si la solicitud ya se agrego al grid
        /// </summary>
        /// <returns>true/false</returns>
        private bool ValidarNoExistaSolicitud()
        {
            try
            {
                string solicitud = null;
                if (!_oGridArqueos.DataTable.IsEmpty)
                {
                    for (int LineaSerie = 0; LineaSerie <= _oGridArqueos.Rows.Count - 1; LineaSerie++)
                    {
                        solicitud = _oGridArqueos.DataTable.GetValue("Solicitud", LineaSerie).ToString();
                        if (solicitud == _oTxtCodigoSolicitud.Value.ToString())
                        {
                            _Application.MessageBox("Esta solicitud ya esta asignada");
                            verificar = false;
                            return false;
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar si existe solicitud *ValidarNoExistaSolicitud* : " + ex.Message);
            }
        }

        /// <summary>
        /// Valida si ya se escanearon las solicitudes sin cerrar el arqueo
        /// </summary>
        /// <returns>true/false</returns>
        private bool ValidarSiYaEscaneo()
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  Code
                                FROM    dbo.[@ARQUEOS]
                                WHERE   U_FechaArqueo IS NULL
                                        AND U_CodigoAsistente = '" + _oTxtCodigoAsociado.Value.ToString() + "'");
                return _oRec.RecordCount > 0 ? true : false;

            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar escaneo sin cierre de arqueo *ValidarSiYaEscaneo* : " + ex.Message);
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
        /// Valida si el arqueo esta cerrado
        /// </summary>
        /// <returns>true/false</returns>
        private bool ValidarSiYaCerroArqueo()
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  Code
                                FROM    dbo.[@ARQUEOS]
                                WHERE   U_FechaArqueo IS NOT NULL
                                        AND CONVERT(VARCHAR(10), U_FechaArqueo, 103) = CONVERT(VARCHAR(10), GETDATE(), 103)
                                        AND U_CodigoAsistente = '" + _oTxtCodigoAsociado.Value.ToString() + "'");
                return _oRec.RecordCount > 0 ? true : false;

            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar escaneo sin cierre de arqueo *ValidarSiYaCerroArqueo* : " + ex.Message);
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
        /// Llena el grid con la información escaneada
        /// </summary>
        /// <param name="query">consulta extra para información</param>
        private void LlenarGridArqueo(string query)
        {
            try
            {
                _oForma = _Application.Forms.Item(formID);
                _oForma.Freeze(true);

                _oForma.DataSources.DataTables.Item(DT_ARQUEOS)
                       .ExecuteQuery(@"SELECT  ROW_NUMBER() OVER ( ORDER BY U_CodigoAsistente ASC ) AS '#' ,
		                                        U_CodigoAsistente AS CodigoAsistente,
		                                        U_NombreAsistente  AS NombreAsistente,
		                                        U_CodigoOficina AS CodigoOficina,
		                                        U_FechaArqueo AS FechaArqueo,
		                                        U_Empresa AS Empresa,
		                                        U_Plan AS 'Plan',
		                                        U_Solicitud AS Solicitud
                                        FROM    dbo.[@ARQUEOS]
                                        WHERE U_CodigoAsistente = '" + _oTxtCodigoAsociado.Value.ToString() + "'" + query);
                _oGridArqueos.DataTable = _oForma.DataSources.DataTables.Item(DT_ARQUEOS);
                FormatoGrid(_oGridArqueos, "Info");
                _oGridArqueos.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                throw new Exception("Error al llenar grid de Arqueo *LlenarGridArqueo* : " + ex.Message);
            }
            finally
            {
                _oForma.Freeze(false);
            }
        }

        /// <summary>
        /// Ingresa la fecha del día del cierre del arqueo
        /// </summary>
        private void CerrarAqueo()
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"UPDATE  dbo.[@ARQUEOS]
                                SET     U_FechaArqueo = GETDATE()
                                WHERE   U_FechaArqueo IS NULL
                                        AND U_CodigoAsistente = '" + _oTxtCodigoAsociado.Value.ToString() + "'");

                _Application.StatusBar.SetText("Arqueo cerrado correctamente...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                _oBtnImprimirArqueo.Item.Visible = true;
                _oTxtCodigoAsociado.Active = true;
                _oTxtCodigoSolicitud.Item.Enabled = false;
                _oForma.Items.Item(BTN_ARQUEO).Enabled = false;

            }
            catch (Exception ex)
            {
                throw new Exception("Error al llenar grid de Arqueo *LlenarGridArqueo* : " + ex.Message);
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
        /// Cierra el arqueo sin solicitudes
        /// </summary>
        private void CerrarAqueoSinSolicitudes()
        {
            try
            {
                string empresa = null;
                string plan = null;
                int code = 0;

                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery("SELECT U_Empresa,U_Descripcion_Plan FROM dbo.[@COMISIONES] WHERE U_Prefijo_Sol =  SUBSTRING('" + _oTxtCodigoSolicitud.Value.ToString() + "', 1,6)");
                empresa = _oRec.Fields.Item("U_Empresa").Value;
                plan = _oRec.Fields.Item("U_Descripcion_Plan").Value;              

                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  ISNULL(MAX(CONVERT(INT, Code)), 0) + 1 AS Code
                                    FROM    dbo.[@ARQUEOS]");
                code = Convert.ToInt32(_oRec.Fields.Item("Code").Value);
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"INSERT INTO dbo.[@ARQUEOS]
                                                ( Code ,
                                                  Name ,
                                                  U_CodigoAsistente ,
                                                  U_NombreAsistente ,
                                                  U_CodigoOficina ,
                                                  U_Empresa ,
                                                  U_Plan ,
                                                  U_Solicitud , 
                                                  U_FechaArqueo
                                                )
                                        VALUES  ( '" + code + "' ," +
                                              "'" + code + "' ," +
                                              "'" + _oTxtCodigoAsociado.Value.ToString().ToUpper() + "' ," +
                                              "'" + _oTxtNombreAsociado.Value.ToString() + "' ," +
                                              "'" + _oTxtCodigoOficina.Value.ToString() + "' ," +
                                              "'" + empresa + "' ," +
                                              "'" + plan + "' ," +
                                              "'' ," +
                                              "GETDATE()" + 
                                            ")");

                _Application.StatusBar.SetText("Arqueo cerrado correctamente...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                _oBtnImprimirArqueo.Item.Visible = true;
                _oTxtCodigoAsociado.Active = true;
                _oTxtCodigoSolicitud.Item.Enabled = false;
                _oForma.Items.Item(BTN_ARQUEO).Enabled = false;

            }
            catch (Exception ex)
            {
                throw new Exception("Error al cerrar Arqueo *CerrarAqueoSinSolicitudes* : " + ex.Message);
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
        /// Obtiene el reporte agregado en SAP
        /// </summary>
        private void CargarReporteArqueo()
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbouiCOM.Form _formaRpt = null;
                SAPbouiCOM.EditText asistente = null;

                _oRec.DoQuery(@"SELECT  MenuUID
                                FROM    dbo.OCMN
                                WHERE   Name = 'Arqueos'
                                        AND Type = 'C'");
                _Application.ActivateMenuItem(_oRec.Fields.Item("MenuUID").Value);
                _formaRpt = _Application.Forms.ActiveForm;
                asistente = _formaRpt.Items.Item("1000003").Specific;
                asistente.Value = _oTxtCodigoAsociado.Value;
                _formaRpt.Items.Item(0).Click();
                _Application.Menus.Item("520").Activate();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al cargar el reporte de Arqueo *CargarReporteArqueo* : " + ex.Message);
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





//Private Sub txtCodigoBar_TextChanged(sender As Object, e As System.EventArgs) Handles txtCodigoBar.TextChanged
//    Dim largo As Long = txtCodigoBar.Text.Length

//    If evalua = False And largo >= 1 Then
//        t_ini = Now
//        evalua = True
//    Else
//        If largo >= 1 Then
//            t_fin = Now
//            Dim seg As Long = DateDiff(DateInterval.Second, t_ini, t_fin)
//            If seg >= 1 Then
//                MessageBox.Show("No se permite captura con teclado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
//                Me.txtCodigoBar.Clear()
//                evalua = False
//            End If
//        End If
//    End If

//    If largo = 0 Then
//        evalua = False
//    End If
//End Sub