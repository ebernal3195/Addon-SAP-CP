using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SAPbouiCOM;

namespace AddonFNR.BL
{
    public abstract class ComportaForm
    {
        #region VARIABLES

        protected SAPbobsCOM.Company _Company;
        protected SAPbouiCOM.Application _Application;
        protected const int FormTypeMenu = 169;

        #endregion

        #region PROPIEDADES

        /// <summary>
        /// ID Formulario
        /// </summary>
        protected string formID
        {
            get;
            set;
        }

        /// <summary>
        /// Compañia
        /// </summary>
        protected SAPbobsCOM.Company company
        {
            set
            {
                _Company = value;
            }
            //get;
            //set;
        }

        /// <summary>
        /// 
        /// </summary>
        protected SAPbouiCOM.Application application
        {
            //get;
            //set;
            set
            {
                _Application = value;
            }
        }

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor ComportaForm
        /// </summary>
        public ComportaForm()
        {
        }

        #endregion

        #region EVENTOS

        /// <summary>
        /// Ejecutar evento
        /// </summary>
        /// <param name="FormUID">Nombre o ID de la forma.</param>
        /// <param name="pVal">Propiedades de la forma</param>
        /// <param name="bubbleEvent">Evento</param>
        public static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool bubbleEvent)
        {

            bubbleEvent = true;

        }

        /// <summary>
        /// Metodo abstracto eventos.
        /// </summary>
        /// <param name="FormUID">Nombre o ID de la forma.</param>
        /// <param name="pVal">Propiedades de la forma</param>
        /// <param name="bubbleEvent">Evento</param>
        public abstract void eventos(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool bubbleEvent);
         

        #endregion

        #region METODOS

        /// <summary>
        /// Cargar XML para la creación del Menú.
        /// </summary>
        /// <param name="FileName">Nombre del archivo.</param>
        private void LoadFromXML(string FileName)
        {
            System.Xml.XmlDocument oXmlDoc = default(System.Xml.XmlDocument);

            oXmlDoc = new System.Xml.XmlDocument();

            string sPath = null;


            sPath = System.IO.Directory.GetCurrentDirectory() + "\\Forms";

            oXmlDoc.Load(sPath + "\\" + FileName + ".srf");

            _Application.LoadBatchActions(oXmlDoc.InnerXml);
        }

        /// <summary>
        /// Muestra el formulario seleccionado.
        /// </summary>
        /// <param name="formID">Nombre o ID de la forma.</param>
        /// <returns>Nombre del Formulario</returns>
        public SAPbouiCOM.Form showForm(string formID)
        {
            SAPbouiCOM.Form fForm = null;
            try
            {

                try
                {
                    LoadFromXML(formID);
                }
                catch { }

                fForm = _Application.Forms.Item(formID);
                fForm.Select();
                return fForm;
            }
            catch (Exception ex)
            {
                return fForm;
            }
            finally
            {

            }
        }

        /// <summary>
        /// Cierra la ventana activa.
        /// </summary>
        public void CerrarVenatana()
        {
            SAPbouiCOM.Form fForm = null;
            try
            {
                fForm = _Application.Forms.Item(formID);
                fForm.Close();
            }
            catch (Exception e)
            {

            }
        }

        #endregion
    }
}
