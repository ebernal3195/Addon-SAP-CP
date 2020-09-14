using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace AddonFNR.BL
{
    class CrystalReportManager
    {
        #region VARIABLES

        private ReportDocument rpt;
        private CrystalDecisions.Shared.ConnectionInfo loginfo;

        #endregion

        #region METODOS

        /// <summary>
        /// Genera el reporte
        /// </summary>
        /// <param name="reportPath">Ruta del reporte</param>
        /// <param name="outputPath">Regresa el reporte</param>
        /// <param name="fechaInicial">Valor fecha inicial</param>
        /// <param name="fechaFinal">Valor fecha final</param>
        /// <param name="oficinaVentas">Código de la oficina</param>
        /// <param name="Usuario">Código del usuario</param>
        /// <param name="_Company">Objeto de SAP</param>
        /// <param name="msgError">Mensaje de error</param>
        /// <returns>true/false</returns>
        public bool RunReport(string reportPath, string outputPath, DateTime fechaInicial, DateTime fechaFinal, string oficinaVentas, SAPbobsCOM.Company _Company, ref string msgError)
        {
            msgError = string.Empty;
            if (System.IO.File.Exists(outputPath))
            {
                try
                {
                    System.IO.File.Delete(outputPath);
                }
                catch (Exception ex)
                {
                    throw new Exception("No se pudo reemplazar el archivo " + System.IO.Path.GetFileName(outputPath) + ", revise que no se esté ejecutando.");
                }
            }
            try
            {
                rpt = new ReportDocument();

                loginfo = new CrystalDecisions.Shared.ConnectionInfo();
                loginfo.ServerName = _Company.Server;
                loginfo.DatabaseName = _Company.CompanyDB;
                loginfo.UserID = _Company.DbUserName;
                loginfo.Password = Extensor.Configuracion.PBD.PBD;   //ConfigurationManager.AppSettings["PBD"].ToString();

                rpt.Load(reportPath);
                Actualiza(ref rpt, loginfo);

                rpt.SetParameterValue("@FechaInicial", fechaInicial);
                rpt.SetParameterValue("@FechaFinal", fechaFinal);
                rpt.SetParameterValue("@AlmacenOficina", oficinaVentas);
                //rpt.SetParameterValue("@Usuario", Usuario);


                CrystalDecisions.Shared.DiskFileDestinationOptions filedest = new CrystalDecisions.Shared.DiskFileDestinationOptions();
                CrystalDecisions.Shared.ExportOptions o = default(CrystalDecisions.Shared.ExportOptions);
                o = new CrystalDecisions.Shared.ExportOptions();

                ExportOptions opt = new ExportOptions();

                opt = rpt.ExportOptions;

                o.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat;
                o.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile;
                filedest.DiskFileName = outputPath;


                o.ExportDestinationOptions = (ExportDestinationOptions)filedest.Clone();

                rpt.Export(o);
                o = null;

                filedest = null;
                rpt.Dispose();
                rpt.Close();


                opt = null;
                rpt = null;

                return true;
            }
            catch (Exception ex)
            {
                msgError = "No se creó el archivo: " + ex.Message + Environment.NewLine +
                   "Detalles: " + ex.InnerException.Message;
                return false;
            }

        }

        /// <summary>
        /// Actualiza la conexión a la base de datos
        /// </summary>
        /// <param name="rpt">
        /// Objeto del reporte
        /// </param>
        /// <param name="loginfo">
        /// Datos de conexión
        /// </param>
        private void Actualiza(ref ReportDocument rpt, ConnectionInfo loginfo)
        {
            try
            {
                CrystalDecisions.Shared.TableLogOnInfo TableLogOnInfo = new TableLogOnInfo();
                CrystalDecisions.Shared.TableLogOnInfos tableLogOnInfos = new TableLogOnInfos();

                TableLogOnInfo.ConnectionInfo = loginfo;
                tableLogOnInfos.Add(TableLogOnInfo);
                ActualizarConexionReporte(rpt, TableLogOnInfo);
            }
            catch (Exception ex)
            {

                throw new Exception("Error al actualizar reporte: " + ex.Message);
            }
        }

        /// <summary>
        /// Actualiza la conexión con el reporte para obtener la información
        /// </summary>
        /// <param name="rpt">
        /// Objeto del reporte
        /// </param>
        /// <param name="tableLogOnInfo">
        /// Tabla de conexión a la base de datos
        /// </param>
        private void ActualizarConexionReporte(ReportDocument rpt, TableLogOnInfo tableLogOnInfo)
        {
            try
            {
                foreach (CrystalDecisions.CrystalReports.Engine.Table x in rpt.Database.Tables)
                {
                    x.ApplyLogOnInfo(tableLogOnInfo);
                }
                rpt.Refresh();
            }
            catch (Exception ex)
            {

                throw new Exception("Error al actualizar conexión reporte: " + ex.Message);
            }
        }

        #endregion
    }
}
 