using AddonFNR.BL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using System.Configuration;
using SAPbouiCOM;


namespace AddonFNR
{
    public sealed class Addon
    {
        #region CONSTANTES

        private const int FRM_DATOS_MAESTROS_SOCIO = 134;
        private const int FRM_SOLICITUD_DE_PLANES = 540000988;
        private const int FRM_ORDEN_DE_COMPRA = 142;
        private const int FRM_ENTRADA_DE_MERCANCIA = 143;
        private const int FRM_FACTURA_DE_PROVEEDOR = 141;
        private const int FRM_SOLICITUD_DE_TRASLADO = 1250000940;
        private const int FRM_TRANSFERENCIA_DE_STOCK = 940;
        private const int FRM_FACTURA_CLIENTES = 133;
        private const int FRM_DATOS_MAESTROS_EMPLEADO = 60100;
        private const int FRM_ALARMAS = 198;

        private const string FRM_TRASPASOS = "UDO_FT_TRASPASOS";
        private const string CAMPO_SERIE = "18_U_E";
        private const int CHAR_PRESS_ENTER = 13; 

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.EventFilters _oFilters;
        private SAPbouiCOM.EventFilter _oFilter;
        public static ListaClase oListClases;
        public static List<DatosConfiguracionCampos> listaDatos = new List<DatosConfiguracionCampos>();
        public static DatosConfiguracionCampos itemDatos = new DatosConfiguracionCampos();
        public static List<ListaCuentasSAP> listaCtasSAP = new List<ListaCuentasSAP>();
        public static ListaCuentasSAP itemCuentas = new ListaCuentasSAP();
        
        public static int LongitudDeSerieConfig = 0;

        static SAPbouiCOM.Form oForm;
        static SAPbobsCOM.Recordset _oRec = null;
        static SAPbobsCOM.Recordset _oRec2 = null;

        public string FormType = "";
        public bool submenu;

        public static List<ListadoDeClases> typeList = new List<ListadoDeClases>();
        static readonly Addon _instancia = new Addon();
        private static ConexionAddon.ConexionAddon _oConnection = new ConexionAddon.ConexionAddon();

        #endregion

        #region INSTANCE

        /// <summary>
        /// Instancia
        /// </summary>
        public static Addon Instance
        {
            get
            {
                return _instancia;

            }
        }

        #endregion

        #region METODOS

        /// <summary>
        /// Inicializar las propiedades de los eventos.
        /// </summary>
        private void SetFilters()
        {
            _oFilters = new SAPbouiCOM.EventFilters();

            _oFilters = new SAPbouiCOM.EventFilters();
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            //_oFilter.AddEx("43520");
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST); 
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORMAT_SEARCH_COMPLETED);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_GRID_SORT);
            _oFilter = _oFilters.Add(SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED);

            _oFilter = _oFilters.Add((SAPbouiCOM.BoEventTypes)SAPbouiCOM.BoAppEventTypes.aet_ShutDown);
            _oFilter = _oFilters.Add((SAPbouiCOM.BoEventTypes)SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged);
            //_oFilter.AddEx("frmClasfTG");
            SBO_Application.SetFilter(_oFilters);
        }

        /// <summary>
        ///  Inicializa los eventos de la forma.
        /// </summary>
        private void SetEvents()
        {
            SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            //SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
        }

        /// <summary>
        /// Termina el proceso del Addon.
        /// </summary>
        public void salirAddon()
        {
            try
            {
                if (oCompany.Connected == true)
                {
                    oCompany.Disconnect();
                }

                System.Windows.Forms.Application.Exit();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Addon término su proceso - " + ex.Message);
                System.Windows.Forms.Application.Exit();
            }
        }

        /// <summary>
        /// Ejecuta la clase de acuerdo a la forma seleccionada.
        /// </summary>
        public void AgregaClase()
        {
            System.Xml.XmlDocument oXMLDoc = new System.Xml.XmlDocument();
            System.Xml.Linq.XDocument oProcesarFormas;
            System.Xml.XmlReader oReaderFormas;

            string FileName = "ConfiguracionXML.xml";

            string OONE_FldMenus = System.IO.Directory.GetCurrentDirectory() + "\\XMLConfig\\";

            try
            {
                if (!System.IO.File.Exists(OONE_FldMenus + FileName))
                {
                    SBO_Application.MessageBox("File Not Found: " + OONE_FldMenus + FileName);
                    salirAddon();
                }
                oXMLDoc.Load(OONE_FldMenus + FileName);

                oReaderFormas = System.Xml.XmlReader.Create(new StringReader(oXMLDoc.InnerXml));

                oProcesarFormas = System.Xml.Linq.XDocument.Load(oReaderFormas);
                XmlTextReader Reader = new XmlTextReader(new StringReader(oXMLDoc.InnerXml));
                XmlSerializer serializer = new XmlSerializer(typeof(ListaClase));
                oListClases = (ListaClase)serializer.Deserialize(Reader);
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// Crea los Menús de acuerdo al archivo XML
        /// </summary>
        public void AddMenus()
        {
            System.Xml.XmlDocument oXMLDoc = new System.Xml.XmlDocument();
            string FileName = "MenuSAP.xml";
            string OONE_FldMenus = System.IO.Directory.GetCurrentDirectory() + "\\XMLConfig\\";

            try
            {
                if (!System.IO.File.Exists(OONE_FldMenus + FileName))
                {
                    SBO_Application.MessageBox("File Not Found: " + OONE_FldMenus + FileName);
                    salirAddon();
                }

                oForm = SBO_Application.Forms.GetFormByTypeAndCount(169, 1);
                oForm.Freeze(true);

                oXMLDoc.Load(OONE_FldMenus + FileName);
                SBO_Application.LoadBatchActions(oXMLDoc.InnerXml);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Addon término de forma inesperada:" + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Update();
                oXMLDoc = null;
            }
        }

        private Addon()
        {
            conectarAddon();
            SBO_Application = ConexionAddon.ConexionAddon.SBO_Application;
            oCompany = ConexionAddon.ConexionAddon._oCompany;
            SetFilters();
            SetEvents();
            AddMenus();
            AgregaClase();
            ObtenerConfiguracionInicial();
            Extensor.CargarConfiguraciones();
            ObtenerCuentasSAP();
        }

        private void ObtenerCuentasSAP()
        {
            try
            {
                listaCtasSAP.Clear();               
                _oRec = null;
                _oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT * FROM dbo.CUENTAS_SAP");

                for (int fila = 0; fila < _oRec.RecordCount; fila++)
                {
                    itemCuentas = new ListaCuentasSAP();
                    itemCuentas.cuenta = _oRec.Fields.Item("Cuenta").Value.ToString();
                    itemCuentas.Documento = _oRec.Fields.Item("Documento").Value.ToString();
                    listaCtasSAP.Add(itemCuentas);
                    _oRec.MoveNext();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener cuentas SAP documentos: " + ex.Message);
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
        /// Obtiene la configuración inicial
        /// </summary>
        private void ObtenerConfiguracionInicial()
        {
            try
            {
                listaDatos.Clear();
                _oRec = null;
                _oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec2 = null;
                _oRec2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  T0.FieldID,
                                        T0.AliasID,
                                        T3.U_Usuario
                                FROM    dbo.CUFD T0
                                        INNER JOIN dbo.UFD1 T1 ON T1.TableID = T0.TableID
                                                                  AND T0.FieldID = T1.FieldID,
                                        dbo.[@SAPCP_CONFIGUSUARIO] T3
                                WHERE   T0.TableID = '@SAPCP_CONFIGUSUARIO'
                                GROUP BY T0.FieldID,
                                        T0.AliasID,
                                        T3.U_Usuario
                                ORDER BY T3.U_Usuario ASC");

                for (int fila = 0; fila < _oRec.RecordCount; fila++)
                {
                    itemDatos = new DatosConfiguracionCampos();
                    itemDatos.usuario = _oRec.Fields.Item("U_Usuario").Value.ToString();
                    itemDatos.campo = "U_" + _oRec.Fields.Item("AliasID").Value.ToString();

                    _oRec2.DoQuery(@"SELECT " + itemDatos.campo.ToString() + " " +
                                    "FROM    dbo.[@SAPCP_CONFIGUSUARIO] " +
                                    "WHERE   U_Usuario = '" + itemDatos.usuario.ToString() + "'");
                    itemDatos.activo = _oRec2.Fields.Item(itemDatos.campo.ToString()).Value == "Y" ? true : false;
                    listaDatos.Add(itemDatos);
                    _oRec.MoveNext();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener configuración inicial: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (_oRec != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oRec);
                    if (_oRec2 != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_oRec2);         
                }
                catch (Exception)
                {
                }                    
            }
        }       

        /// <summary>
        /// Conecta el Addon
        /// </summary>
        private static void conectarAddon()
        {
            try
            {
                _oConnection = new ConexionAddon.ConexionAddon();
                if (!(bool)_oConnection.Conectar())
                {
                    System.Environment.Exit(0);
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                return;
            }
        }

        #endregion

        #region EVENTOS

        /// <summary>
        /// Maneja los eventos para terminar la aplicación y cuando cambia de compañia
        /// </summary>
        /// <param name="eventType"></param>
        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes eventType)
        {
            //Finaliza la Aplicación
            if (eventType == SAPbouiCOM.BoAppEventTypes.aet_ShutDown)
            {
                salirAddon();
            }
            //Cambiar Empresa
            if (eventType == SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged)
            {
                salirAddon();
            }
        }

        /// <summary>
        /// Obtener la forma que se ejecuto.
        /// </summary>
        /// <param name="pVal">Propiedades de la forma</param>
        /// <param name="bubbleEvent">Evento</param>
        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;//Default Valué

            if (pVal.BeforeAction)
            {
                try
                {
                    SBO_Application.Forms.Item(pVal.MenuUID);
                    return;
                }
                catch
                {
                }

                Ejecutaclase(pVal.MenuUID);

            }

        }

        /// <summary>
        /// Ejecuta la clase
        /// </summary>
        /// <param name="pMenuID">
        /// Forma que se va a ejecutar
        /// </param>
        public void Ejecutaclase(string pMenuID, List<Datos> lDatos = null)
        {
            string form = pMenuID;
            ListaClaseDependencia le;
            le = oListClases.Items.ToList().SingleOrDefault(p => p.Forma == form);
            if (le != null)
            {
                ListadoDeClases item = new ListadoDeClases();
                item._forma = form;
                item.Tipo = Type.GetType("AddonFNR.BL." + le.NombreClase);
                typeList.Add(item);

                int max = typeList.Count();
                object o = typeList[max - 1].Tipo.InvokeMember(le.NombreClase, BindingFlags.CreateInstance, null, null,
                    new object[] { SBO_Application, oCompany, form, lDatos });

            }
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.FormTypeEx == "0" && pVal.FormTypeCount == 2 &&
                            pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                {
                    SAPbouiCOM.Form formaEmergente = null;
                    formaEmergente = SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                    if (formaEmergente.Title == "Mensaje sistema")
                    {
                        SAPbouiCOM.StaticText itemMensaje = null;

                        itemMensaje = formaEmergente.Items.Item(7).Specific;
                        var msgDescripcion = itemMensaje.Caption;

                        if (msgDescripcion.Contains(" ya se ha definido en el sistema en un campo único."))
                        {
                            SAPbouiCOM.Button btnCancel = null;
                            btnCancel = formaEmergente.Items.Item(1).Specific;
                            btnCancel.Item.Click();
                        }

                    }
                }

                if (pVal.BeforeAction == true)
                {                  

                    if (pVal.FormTypeEx == FRM_DATOS_MAESTROS_SOCIO.ToString() && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        Instance.Ejecutaclase("2561");
                    }

                    if (pVal.FormTypeEx == FRM_SOLICITUD_DE_PLANES.ToString() && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        Instance.Ejecutaclase("39698");
                    }

                    if (pVal.FormTypeEx == FRM_ORDEN_DE_COMPRA.ToString() && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        Instance.Ejecutaclase("2305");
                    }

                    if (pVal.FormTypeEx == FRM_ENTRADA_DE_MERCANCIA.ToString() && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        Instance.Ejecutaclase("2306");
                    }

                    if (pVal.FormTypeEx == FRM_FACTURA_DE_PROVEEDOR.ToString() && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        Instance.Ejecutaclase("2308");
                    }

                    if (pVal.FormTypeEx == FRM_SOLICITUD_DE_TRASLADO.ToString() && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        Instance.Ejecutaclase("3088");
                    }

                    if (pVal.FormTypeEx == FRM_TRANSFERENCIA_DE_STOCK.ToString() && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        Instance.Ejecutaclase("3080");
                    }

                    if (pVal.FormTypeEx == FRM_FACTURA_CLIENTES.ToString() && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        Instance.Ejecutaclase("2053");
                    }

                    if (pVal.FormTypeEx == FRM_DATOS_MAESTROS_EMPLEADO.ToString() && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        Instance.Ejecutaclase("3590");
                    }

                    if ( pVal.FormTypeEx == FRM_TRASPASOS)                   
                    {
                        if (pVal.EventType == BoEventTypes.et_KEY_DOWN && pVal.ItemUID == CAMPO_SERIE && pVal.CharPressed == CHAR_PRESS_ENTER)
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }            
                }
            }
            catch (Exception ex)
            {

            }
        }

        #endregion

    }
}
