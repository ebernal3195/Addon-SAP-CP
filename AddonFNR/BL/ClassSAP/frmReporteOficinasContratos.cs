using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonFNR.BL
{
    class frmReporteOficinasContratos : ComportaForm
    {
        #region CONSTANTES

        private const string FRM_RPT_CORTE_SOLICITUDES = "frmCorteSolicitudes";
        private const string TXT_FECHA_INICIAL = "txtFI";
        private const string TXT_FECHA_FINAL = "txtFF";
        private const string CMB_OFICINA_VENTAS = "cmbOficin";
        //private const string TXT_USUARIO = "txtUser";
        private const string BTN_IMPRIMIR = "btnImpri";
        private const string BTN_CANCELAR = "btnCancel";
        private const string BTN_SAP_BUSCAR = "1281";
        private const string BTN_SAP_CREAR = "1282";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForma = null;
        private static bool _oRptCorteSolicitudes = false;
        private SAPbobsCOM.Recordset _oRec = null;
        private SAPbouiCOM.EditText _oTxtFechaInicial = null;
        private SAPbouiCOM.EditText _oTxtFechaFinal = null;
        private SAPbouiCOM.ComboBox _oCmbOficinaVentas = null;
        //private SAPbouiCOM.EditText _oTxtUsuario = null;

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
        public frmReporteOficinasContratos(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            if (_oRptCorteSolicitudes == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                showForm(formID);
                inicializarComponentes();
                setEventos();
                _oRptCorteSolicitudes = true;
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
                if (_oRptCorteSolicitudes != false && pVal.FormType != FormTypeMenu && formID == FormUID)
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
                        _oRptCorteSolicitudes = false;
                        Addon.typeList.RemoveAll(p => p._forma == formID);
                        return;
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if (pVal.ItemUID == BTN_CANCELAR)
                        {
                            _oForma.Close();
                        }

                        if (pVal.ItemUID == BTN_IMPRIMIR)
                        {
                            _oTxtFechaInicial = _oForma.Items.Item(TXT_FECHA_INICIAL).Specific;
                            _oTxtFechaFinal = _oForma.Items.Item(TXT_FECHA_FINAL).Specific;
                            _oCmbOficinaVentas = _oForma.Items.Item(CMB_OFICINA_VENTAS).Specific;
                            //_oTxtUsuario = _oForma.Items.Item(TXT_USUARIO).Specific;

                            if (ValidarCampos())
                            {
                                if (Extensor.ValidarImpresionCorteSolicitudes(_Company, _oCmbOficinaVentas.Selected.Value.ToString()))
                                {
                                    if (_Application.MessageBox("¿Desea generar el corte de solicitudes?", 2, "Si", "No") == 1)
                                    {
                                        if (ImprimirReporteCorteSolicitudes())
                                        {
                                            _Application.StatusBar.SetText("Generar reporte terminado correctamente...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                        }
                                    }
                                }
                                else
                                {
                                    _Application.MessageBox("El reporte ya fue impreso");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en evento *clsReporteOficinasContratos* : " + ex.Message);
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
                        if (_Application.Forms.ActiveForm.UniqueID == FRM_RPT_CORTE_SOLICITUDES)
                            BubbleEvent = false;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error en ManuEvent *clsReporteOficinasContratos* : " + ex.Message);
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
                _oForma.DataSources.UserDataSources.Add(TXT_FECHA_INICIAL, BoDataType.dt_DATE);
                _oTxtFechaInicial = (SAPbouiCOM.EditText)_oForma.Items.Item(TXT_FECHA_INICIAL).Specific;
                _oTxtFechaInicial.DataBind.SetBound(true, "", TXT_FECHA_INICIAL);

                //Tipo de dato fecha final
                _oForma.DataSources.UserDataSources.Add(TXT_FECHA_FINAL, BoDataType.dt_DATE);
                _oTxtFechaFinal = (SAPbouiCOM.EditText)_oForma.Items.Item(TXT_FECHA_FINAL).Specific;
                _oTxtFechaFinal.DataBind.SetBound(true, "", TXT_FECHA_FINAL);

                _oCmbOficinaVentas = (SAPbouiCOM.ComboBox)_oForma.Items.Item(CMB_OFICINA_VENTAS).Specific;
                _oCmbOficinaVentas.LlenarCargarOficinas(_Company, _oForma);
            

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
        /// Imprime el reporte del corte de la solicitud
        /// </summary>
        /// <returns>true / false</returns>
        private bool ImprimirReporteCorteSolicitudes()
        {
            try
            {
                //_" + _Company.UserName.ToString() + "
                CrystalReportManager crManager = new CrystalReportManager();
                string msgError = string.Empty;
                string reportPath = System.IO.Directory.GetCurrentDirectory() + "\\Reportes\\" + "ReporteCorteSolicitudes.rpt";
                string carpeta = @"C:\\CORTE SOLICITUDES_" + _oCmbOficinaVentas.Selected.Description.Replace(":"," ");
                string outPDF = carpeta + "\\Oficinas a contratos_" + _oCmbOficinaVentas.Selected.Description.Replace(":", " ") + "_" + DateTime.Now.ToString("dd-MM-yyyy H-mm-ss") + ".pdf";

                if (!(System.IO.Directory.Exists(carpeta)))
                {
                    System.IO.Directory.CreateDirectory(carpeta);
                }

                _Application.StatusBar.SetText("Generando reporte...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

                string fechaDesde = _oTxtFechaInicial.Value.Substring(0, 4) + "-" + _oTxtFechaInicial.Value.Substring(4, 2) + "-" +
                          _oTxtFechaInicial.Value.Substring(6, 2);

                string fechaHasta = _oTxtFechaFinal.Value.Substring(0, 4) + "-" + _oTxtFechaFinal.Value.Substring(4, 2) + "-" +
                    _oTxtFechaFinal.Value.Substring(6, 2);

                DateTime fInicial = Convert.ToDateTime(fechaDesde);
                DateTime fFinal = Convert.ToDateTime(fechaHasta);

                if (crManager.RunReport(reportPath, outPDF, fInicial, fFinal, _oCmbOficinaVentas.Selected.Value.ToString(), _Company, ref msgError))
                {
                    try
                    {
                        if (System.IO.File.Exists(outPDF))
                        {
                            System.Diagnostics.Process.Start(outPDF);
                            InsertarFechaImpresion();
                        }
                    }
                    catch (Exception )
                    {
                    }
                    return true;
                }
                else
                {
                    _Application.StatusBar.SetText("No se pudo crear el PDF: " + msgError, SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return false;
                }

            }
            catch (Exception ex)
            {

                throw new Exception("Error al imprimir el reporte del corte de solicitud *ImprimirReporteCorteSolicitudes* : " + ex.Message);
            }
        }     

        /// <summary>
        /// Valida que todos los datos estén correctos
        /// </summary>
        /// <returns>true/false</returns>
        private bool ValidarCampos()
        {
            try
            {
                if (_oTxtFechaInicial.Value.Equals(""))
                {
                    _Application.StatusBar.SetText("El parámetro [Fecha inicial] es obligatorio", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (_oTxtFechaFinal.Value.Equals(""))
                {
                    _Application.StatusBar.SetText("El parámetro [Fecha final] es obligatorio", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (_oCmbOficinaVentas.Selected == null)
                {
                        _Application.StatusBar.SetText("El parámetro [Oficina de ventas] es obligatorio ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        return false;                    
                }
             
                //if (_oTxtUsuario.Value.Equals(""))
                //{
                //    _Application.StatusBar.SetText("El parámetro [Usuarios SAP] es obligatorio ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //    return false;
                //}

                string fechaDesde = _oTxtFechaInicial.Value.Substring(0, 4) + "-" + _oTxtFechaInicial.Value.Substring(4, 2) + "-" +
                            _oTxtFechaInicial.Value.Substring(6, 2);

                string fechaHasta = _oTxtFechaFinal.Value.Substring(0, 4) + "-" + _oTxtFechaFinal.Value.Substring(4, 2) + "-" +
                    _oTxtFechaFinal.Value.Substring(6, 2);

                DateTime fInicial = Convert.ToDateTime(fechaDesde);
                DateTime fFinal = Convert.ToDateTime(fechaHasta);

                if (fInicial > fFinal)
                {
                    _Application.StatusBar.SetText("El parámetro [Fecha inicial] es mayor al parámetro [Fecha final]", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }
                return true;

            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar campos *ValidarCampos* : " + ex.Message);
            }
        }

        /// <summary>
        /// Inserta o actualiza la fecha de impresion;
        /// </summary>
        private void InsertarFechaImpresion()
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                
                _oRec.DoQuery(@"SELECT TOP 1
                                        U_Oficina
                                FROM    dbo.[@SAPCP_CONFIGCORTSOL]
                                WHERE   U_Oficina = '" + _oCmbOficinaVentas.Selected.Value.ToString() + "' ");

                if(_oRec.RecordCount == 0)
                {
                    _oRec = null;
                    _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    _oRec.DoQuery(@"SELECT  ISNULL(MAX(CONVERT(INT, Code)), 0) + 1 AS Code
                                    FROM    dbo.[@SAPCP_CONFIGCORTSOL]");
                    string code = _oRec.Fields.Item("Code").Value.ToString();


                    _oRec = null;
                    _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    _oRec.DoQuery(@"INSERT	INTO	dbo.[@SAPCP_CONFIGCORTSOL] (
	                                    Code,
	                                    [Name],
	                                    U_Oficina,
	                                    U_FechaImpresion
                                    ) VALUES ( 
	                                    '" + code + "', "  +
	                                    "'" + code + "', "  +
                                        "'" + _oCmbOficinaVentas.Selected.Value.ToString() + "', " +
	                                    "'" + DateTime.Now.ToShortDateString() + "') ");
                }
                else
                {
                    if(Extensor.ValidarImpresionCorteSolicitudes(_Company,_oCmbOficinaVentas.Selected.Value.ToString()))
                    {
                        _oRec = null;
                        _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        _oRec.DoQuery(@"UPDATE  dbo.[@SAPCP_CONFIGCORTSOL]
                                        SET     U_FechaImpresion = CONVERT(VARCHAR(10), GETDATE(), 103)
                                        WHERE   U_Oficina = '" + _oCmbOficinaVentas.Selected.Value.ToString() + "'");
                    }
                }

            }
            catch (Exception)
            {                
                throw;
            }
            finally
            {
                if (_oRec != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_oRec);
            }
        }

        #endregion
    }
}
