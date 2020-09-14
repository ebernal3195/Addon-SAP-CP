using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonFNR.BL
{
    class clsSeleccionarNumerosDeSerie : ComportaForm
    {
        #region CONSTANTES

        private const int FRM_SELECCIONAR_NUMEROS_DE_SERIE = 25;
        private const string GRID_ARTICULOS_SERIE = "3";
        private const string GRID_SERIE_DEFINIR = "5";
        private const string GRID_SELECCIONADOS = "55";

        private const string CAMPO_FILTRO = "10000059";

        private const string COLUMNA_CLAVE_ARTICULO = "1";
        private const string COLUMNA_NIVEL_LINEA_ARTICULO = "0";
        private const string COLUMNA_NIVEL_LINEA_SERIE = "0";
        private const string COLUMNA_NUMERO_SERIE = "19";
        private const string COLUMNA_TOTAL_SELECCIONADO = "5";
        private const string COLUMNA_CANTIDAD = "4";

        private const string BTN_ACTUALIZAR = "1";
        private const string BTN_ASIGNAR_SERIE = "8";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForm = null;
        private static bool _oSeleccionarNumerosDeSerie = false;
        private static List<Datos> lDatosTransferenciaStock = null;
        private SAPbouiCOM.Matrix _oMatrixSerieArticulos = null;
        private SAPbouiCOM.Matrix _oMatrixSerieDefinir = null;
        private SAPbouiCOM.Matrix _oMatrixSeriesSeleccionados = null;
        private SAPbouiCOM.EditText _oClaveArticulo = null;
        private SAPbouiCOM.EditText _oTotalSeleccionados = null;
        private SAPbouiCOM.EditText _oCantidad = null;
        private SAPbouiCOM.EditText _oTxtFiltro = null;
        private SAPbouiCOM.ProgressBar oProgBar = null;

        private bool ProcesoActivo = false;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de seleccionar números de serie
        /// </summary>
        /// <param name="_Application">Objeto de la conexión de SAP</param>
        /// <param name="_Company">Objeto de la empresa</param>
        /// <param name="form">Nombre de la forma</param>
        public clsSeleccionarNumerosDeSerie(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            lDatosTransferenciaStock = new List<Datos>(lDatos);
            lDatos.Clear();
            if (_oSeleccionarNumerosDeSerie == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                setEventos();
                _oSeleccionarNumerosDeSerie = true;
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
                if (pVal.BeforeAction == false && pVal.FormType == FRM_SELECCIONAR_NUMEROS_DE_SERIE)
                {
                    if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
                    {
                        _Application.ItemEvent -= new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                        Dispose();
                        application = null;
                        company = null;
                        _oSeleccionarNumerosDeSerie = false;
                        Addon.typeList.RemoveAll(p => p._forma == formID);
                        return;
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_ACTIVATE && pVal.ActionSuccess == true)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        _oTxtFiltro = _oForm.Items.Item(CAMPO_FILTRO).Specific;
                        if (ProcesoActivo == false)
                        {
                            if (lDatosTransferenciaStock.Count != 0)
                            {
                                AsignarNumerosDeSeries(_oForm);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Error en método 'eventos' *clsSeleccionarNumerosDeSerie* : " + ex.Message);
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
        /// Asigna los numero de serie correspondientes a cada artículo
        /// </summary>
        /// <param name="_oForm">Forma activa</param>
        private void AsignarNumerosDeSeries(Form _oForm)
        {
            try
            {
                _oMatrixSerieArticulos = _oForm.Items.Item(GRID_ARTICULOS_SERIE).Specific;
                _oMatrixSerieDefinir = _oForm.Items.Item(GRID_SERIE_DEFINIR).Specific;
                List<string> mensajes = new List<string>();
                List<string> solicitudes = new List<string>();
                bool seEncontroSerie = false;

                for (int noLinea = 1; noLinea <= _oMatrixSerieArticulos.RowCount; noLinea++)
                {
                    ProcesoActivo = true;
                    _oForm.Freeze(false);
                    seEncontroSerie = false;

                    _oClaveArticulo = (SAPbouiCOM.EditText)_oMatrixSerieArticulos.Columns.Item(COLUMNA_CLAVE_ARTICULO).Cells.Item(noLinea).Specific;
                    _oTotalSeleccionados = (SAPbouiCOM.EditText)_oMatrixSerieArticulos.Columns.Item(COLUMNA_TOTAL_SELECCIONADO).Cells.Item(noLinea).Specific;

                    var serieI = from p in lDatosTransferenciaStock
                                 where p.itemCode == _oClaveArticulo.Value.ToString() && p.noLinea == noLinea
                                 select p.serieInial;
                    var serieF = from p in lDatosTransferenciaStock
                                 where p.itemCode == _oClaveArticulo.Value.ToString() && p.noLinea == noLinea
                                 select p.serieFinal;

                    int x = 1;

                    string PrefijoSerie = serieI.ElementAt(0).ToString().Substring(0, 6);
                    int serieInicial = Convert.ToInt32(serieI.ElementAt(0).ToString().Substring(6));
                    int serieFinal = Convert.ToInt32(serieF.ElementAt(0).ToString().Substring(6));
                    int totalSeries = (serieFinal - serieInicial) + 1;
                    int longitudSerieInicial = serieI.ElementAt(0).ToString().Substring(6).Length;
                    int longitudSerieFinal = serieF.ElementAt(0).ToString().Substring(6).Length;

                    if (_oTotalSeleccionados.Value == "0.0")
                    {
                        _oForm.Freeze(false);
                        _oTxtFiltro = _oForm.Items.Item(CAMPO_FILTRO).Specific;

                        if (serieInicial == serieFinal)
                        {
                            if (serieInicial.ToString().Length < longitudSerieInicial)
                            {
                                _oTxtFiltro.Value = PrefijoSerie + serieInicial.ToString().PadLeft(longitudSerieInicial, '0'); // +"0" + serieInicial;
                            }
                            else
                            {
                                _oTxtFiltro.Value = PrefijoSerie + serieInicial;
                            }
                        }
                        else
                        {
                            for (int i = 0; i < serieI.ElementAt(0).ToString().Length; i++ )
                            {
                                if(serieI.ElementAt(0).ToString().Substring(i,1) != serieF.ElementAt(0).ToString().Substring(i,1))
                                {
                                    _oTxtFiltro.Value = serieI.ElementAt(0).ToString().Substring(0, i);
                                    break;
                                }
                            }
                                //_oTxtFiltro.Value = PrefijoSerie;
                        }

                        string SerFinal = null;
                        if (serieFinal.ToString().Length < longitudSerieFinal)
                        {
                            SerFinal = PrefijoSerie + serieFinal.ToString().PadLeft(longitudSerieFinal, '0'); // +"0" + serieFinal;
                        }
                        else
                        {
                            SerFinal = PrefijoSerie + serieFinal;
                        }

                        //Presionar de forma automática tabulación para que el sistema realice la búsqueda
                        _Application.SendKeys("{TAB}");
                        _oForm.Select();
                        if (_oMatrixSerieDefinir.RowCount != 0)
                        {
                            _oForm.Freeze(true);
                            _oMatrixSerieDefinir.Columns.Item("19").TitleObject.Click(BoCellClickType.ct_Regular, 4096);                 
                            _oMatrixSerieDefinir.Columns.Item("19").TitleObject.Click(BoCellClickType.ct_Double);
                            _oForm.Select();
                            // oProgBar = _Application.StatusBar.CreateProgressBar("Seleccionando series del articulo: " + _oClaveArticulo.Value.ToString(), totalSeries, false);
                            _Application.StatusBar.SetText("Buscando series por favor espere....", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning);
                            for (int noLineaSerie = 1; noLineaSerie <= _oMatrixSerieDefinir.RowCount; noLineaSerie++)
                            {
                                _oForm.Select();
                                if (seEncontroSerie == false)
                                {
                                    _Application.StatusBar.SetText("Buscando series por favor espere...." + noLineaSerie + " de " + _oMatrixSerieDefinir.RowCount.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                }
                                //Realiza la función de mantener presionado CTRL y seleccionar cada fila para asignar la serie.
                                SAPbouiCOM.EditText SerieSeleccionada = _oMatrixSerieDefinir.Columns.Item(COLUMNA_NUMERO_SERIE).Cells.Item(noLineaSerie).Specific;

                                try
                                {
                                    if (SerFinal != SerieSeleccionada.Value.ToString())
                                    {
                                        if ((serieInicial.ToString().Length < longitudSerieInicial ?
                                            PrefijoSerie + serieInicial.ToString().PadLeft(longitudSerieInicial, '0') : // "0" + serieInicial.ToString() :
                                                PrefijoSerie + serieInicial.ToString()) == SerieSeleccionada.Value.ToString())
                                        {
                                            //  oProgBar.Text = "Seleccionando serie " + _oClaveArticulo.Value.ToString() + " : " + x + " de " + totalSeries;
                                            _Application.StatusBar.SetText("Seleccionando serie " + _oClaveArticulo.Value.ToString() + " : " + x + " de " + totalSeries,
                                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                            seEncontroSerie = true;
                                            _oMatrixSerieDefinir.Columns.Item(0).Cells.Item(noLineaSerie).Click(BoCellClickType.ct_Regular, 4096);
                                            x += 1;
                                            serieInicial += 1;
                                            //oProgBar.Value += 1;
                                        }
                                    }
                                    else
                                    {
                                        //x += 1;
                                        //oProgBar.Text = "Seleccionando serie " + _oClaveArticulo.Value.ToString() + " : " + x + " de " + totalSeries;
                                        _Application.StatusBar.SetText("Seleccionando serie " + _oClaveArticulo.Value.ToString() + " : " + x + " de " + totalSeries,
                                       BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                        _oMatrixSerieDefinir.Columns.Item(0).Cells.Item(noLineaSerie).Click(BoCellClickType.ct_Regular, 4096);
                                        break;
                                    }
                                }
                                catch (Exception)
                                {
                                    solicitudes.Add(SerFinal);

                                }                              
                            }
                            _oForm.Select();
                            // oProgBar.Stop();
                            int selRow = _oMatrixSerieDefinir.GetNextSelectedRow(0, BoOrderType.ot_SelectionOrder);

                            if (selRow != -1)
                            {
                                _oForm.Items.Item(BTN_ASIGNAR_SERIE).Click();
                                _oForm.Items.Item(BTN_ACTUALIZAR).Click();
                            }

                            if (noLinea < _oMatrixSerieArticulos.RowCount)
                            {
                                _oMatrixSerieArticulos.Columns.Item(COLUMNA_NIVEL_LINEA_ARTICULO).Cells.Item(noLinea + 1).Click(BoCellClickType.ct_Regular, 4096);
                            }
                            _oForm.Select();
                            _oTotalSeleccionados = (SAPbouiCOM.EditText)_oMatrixSerieArticulos.Columns.Item(COLUMNA_TOTAL_SELECCIONADO).Cells.Item(noLinea).Specific;
                            _oCantidad = (SAPbouiCOM.EditText)_oMatrixSerieArticulos.Columns.Item(COLUMNA_CANTIDAD).Cells.Item(noLinea).Specific;

                            if (_oTotalSeleccionados.Value != _oCantidad.Value)
                            {
                                solicitudes.Add(SerFinal);
                                mensajes.Add("No se encontraron algunas series, favor de verificar");
                            }
                        }
                        else
                        {
                            solicitudes.Add(SerFinal);
                            mensajes.Add( "No se encontraron algunas series, favor de verificar");
                        }
                    }
                }
                lDatosTransferenciaStock.Clear();
                if(mensajes.Count() != 0)
                {
                     var listaSolicitudes = string.Join(", ", solicitudes.Select(s => s.ToString()));
                    string mostrar = mensajes[0] + Environment.NewLine + listaSolicitudes;
                    _Application.MessageBox(mostrar);
                }
                else
                {
                    _Application.MessageBox("Se agregaron las series correctamente.");
                }              
            }
            catch (Exception ex)
            {
                //oProgBar.Stop();
                throw new Exception("Error al asignar números de serie *AsignarNumerosDeSeries* : " + ex.Message);
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }

        #endregion
    }
}
