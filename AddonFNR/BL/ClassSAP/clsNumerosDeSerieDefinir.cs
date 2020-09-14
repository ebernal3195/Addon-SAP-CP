using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonFNR.BL
{
    class clsNumerosDeSerieDefinir : ComportaForm
    {
        #region CONSTANTES

        private const int FRM_NUMEROS_DE_SERIE = 21;
        private const string GRID_ARTICULOS_SERIE = "43";
        private const string GRID_SERIE_DEFINIR = "3";
        private const string COLUMNA_CLAVE_ARTICULO = "5";
        private const string COLUMNA_NIVEL_LINEA = "0";

        private const string COLUMNA_SERIE_FABRICANTE = "1";
        private const string COLUMNA_NUMERO_SERIE = "54";
        private const string COLUMNA_TOTAL_NECESITADOS = "37";

        private const string BTN_ACTUALIZAR = "1";
        private const string VENTANA_EMERGENTE = "0";

        #endregion

        #region VARIABLES

        private SAPbouiCOM.Form _oForm = null;
        private static bool _oNumerosDeSerie = false;
        private static List<Datos> lDatosEntradaMercancia = null;
        private SAPbouiCOM.Matrix _oMatrixSerieArticulos = null;
        private SAPbouiCOM.Matrix _oMatrixSerieDefinir = null;
        private SAPbouiCOM.EditText _oClaveArticulo = null;
        private SAPbouiCOM.EditText _oTotalNecesarios = null;
        private SAPbouiCOM.ProgressBar oProgBar = null;

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor de numero de serie definir
        /// </summary>
        /// <param name="_Application">Objeto de la conexión de SAP</param>
        /// <param name="_Company">Objeto de la empresa</param>
        /// <param name="form">Nombre de la forma</param>
        public clsNumerosDeSerieDefinir(SAPbouiCOM.Application _Application, SAPbobsCOM.Company _Company, string form, List<Datos> lDatos = null)
        {
            lDatosEntradaMercancia = new List<Datos>(lDatos);
            if (_oNumerosDeSerie == false)
            {
                company = _Company;
                application = _Application;
                formID = form;
                setEventos();
                _oNumerosDeSerie = true;
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
                if (pVal.BeforeAction == false && pVal.FormType == FRM_NUMEROS_DE_SERIE)
                {
                    if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
                    {
                        _Application.ItemEvent -= new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                        Dispose();
                        application = null;
                        company = null;
                        _oNumerosDeSerie = false;
                        Addon.typeList.RemoveAll(p => p._forma == formID);
                        return;
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.ActionSuccess == true)
                    {
                        _oForm = _Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        AsignarNumerosDeSeries(_oForm);
                    }
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Error en método 'eventos' *clsNumerosDeSerieDefinir*  : " + ex.Message);
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
                _oForm.Freeze(true);
                _oMatrixSerieArticulos = _oForm.Items.Item(GRID_ARTICULOS_SERIE).Specific;

                for (int noLinea = 1; noLinea <= _oMatrixSerieArticulos.RowCount; noLinea++)
                {
                    _oMatrixSerieDefinir = _oForm.Items.Item(GRID_SERIE_DEFINIR).Specific;
                    if (_oMatrixSerieDefinir.RowCount == 1)
                    {

                        _oClaveArticulo = (SAPbouiCOM.EditText)_oMatrixSerieArticulos.Columns.Item(COLUMNA_CLAVE_ARTICULO).Cells.Item(noLinea).Specific;
                        _oTotalNecesarios = (SAPbouiCOM.EditText)_oMatrixSerieArticulos.Columns.Item(COLUMNA_TOTAL_NECESITADOS).Cells.Item(noLinea).Specific;

                        var serieI = lDatosEntradaMercancia.ElementAt(noLinea - 1).serieInial;
                        var serieF = lDatosEntradaMercancia.ElementAt(noLinea - 1).serieFinal;

                        int x = 1;

                        string PrefijoSerie = serieI.ToString().Substring(0, 6);
                        int serieInicial = Convert.ToInt32(serieI.ToString().Substring(6));
                        int longitudSerie = serieI.ToString().Substring(6).Length;
                        int serieFinal = Convert.ToInt32(serieF.ToString().Substring(6));
                        int totalSeries = (serieFinal - serieInicial) + 1;
                        double necesarios = Convert.ToDouble(_oTotalNecesarios.Value);

                        if (totalSeries <= necesarios)
                        {
                            while (serieInicial <= serieFinal)
                            {
                                _Application.StatusBar.SetText("Seleccionando serie " + _oClaveArticulo.Value.ToString() + " : " + x + " de " + totalSeries,
                                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

                                if (serieInicial.ToString().Length < longitudSerie)
                                {
                                    _oMatrixSerieDefinir.Columns.Item(COLUMNA_SERIE_FABRICANTE).Cells.Item(x).Specific.Value = PrefijoSerie + serieInicial.ToString().PadLeft(longitudSerie, '0');
                                    _oMatrixSerieDefinir.Columns.Item(COLUMNA_NUMERO_SERIE).Cells.Item(x).Specific.Value = PrefijoSerie + serieInicial.ToString().PadLeft(longitudSerie, '0');
                                }
                                else
                                {
                                    _oMatrixSerieDefinir.Columns.Item(COLUMNA_SERIE_FABRICANTE).Cells.Item(x).Specific.Value = PrefijoSerie + serieInicial.ToString();
                                    _oMatrixSerieDefinir.Columns.Item(COLUMNA_NUMERO_SERIE).Cells.Item(x).Specific.Value = PrefijoSerie + serieInicial.ToString();
                                }

                                x += 1;
                                serieInicial += 1;
                            }
                        }
                        else
                        {
                            _Application.MessageBox("El total de series es mayor al total necesitado: " + _oClaveArticulo.Value.ToString());
                        }
                    }

                    _oMatrixSerieArticulos = _oForm.Items.Item(GRID_ARTICULOS_SERIE).Specific;
                    if (_oMatrixSerieArticulos.RowCount >= 1)
                    {
                        SAPbouiCOM.Button item = _oForm.Items.Item(BTN_ACTUALIZAR).Specific;
                        if (item.Caption == "Actualizar")
                        {
                            _oForm.Items.Item(BTN_ACTUALIZAR).Click();
                        }
                    }
                    _oMatrixSerieArticulos = _oForm.Items.Item(GRID_ARTICULOS_SERIE).Specific;
                    if (noLinea < _oMatrixSerieArticulos.RowCount)
                    {
                        _oMatrixSerieArticulos.Columns.Item(COLUMNA_NIVEL_LINEA).Cells.Item(noLinea + 1).Click();
                    }

                }
                lDatosEntradaMercancia.Clear();
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Form - Bad Value"))
                {
                    _Application.MessageBox("El número de Solicitud ya existe");
                }
                else
                {
                    throw new Exception("Error al asignar números de serie *AsignarNumerosDeSeries* : " + ex.Message);
                }
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }

        #endregion
    }
}
