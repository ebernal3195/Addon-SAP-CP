using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using System.Windows.Forms;

namespace AddonFNR.BL
{
    public static class Extensor
    {
        #region VARIABLES

        private static SAPbobsCOM.Recordset _oRec = null;
        public static CONFIGURACION Configuracion;

        #endregion

        #region METODOS

        /// <summary>
        /// Obtiene el permiso del campo
        /// </summary>
        /// <param name="campo">Campos a solicitar</param>
        /// <param name="usuario">Usuario con permiso</param>
        /// <returns>true - Activo / False - Inactivo</returns>
        public static bool ObtenerAutorizacionCampo(string campo, string usuario)
        {
            try
            {
                var asdf = from p in Addon.listaDatos where p.usuario == usuario && p.campo == "U_" + campo select p.activo;
                if (asdf.Any() != false)
                {
                    return asdf.ElementAt(0);
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener autorización del campo *ObtenerAutorizacionCampo*: " + ex.Message);
            }
        }

        /// <summary>
        /// Recibe los campos que se van a bloquear
        /// </summary>
        /// <param name="campos">Campos a bloquear</param>
        /// <param name="usuario">Usuario asignado</param>
        /// <param name="forma">Forma activa</param>
        public static void ActivarInactivarCampo(string[] campos, string usuario, SAPbouiCOM.Form forma)
        {
            try
            {
                for (int i = 0; i < campos.Length; i++)
                {
                    forma.Items.Item(campos[i]).Enabled = Extensor.ObtenerAutorizacionCampo(campos[i], usuario);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al activar/inactivar campos *ActivarInactivarCampo* : " + ex.Message);
            }
        }

        /// <summary>
        ///  Obtiene el grupo que se encuentra definido en la tabla
        /// </summary>
        /// <param name="_oForm">Forma activa</param>
        /// <param name="_Company">Objeto Company</param>
        /// <returns>Grupo</returns>
        public static string ObtenerGrupoSocioNegocio(SAPbouiCOM.Form _oForm, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT TOP 1
                                        U_Grupo
                                FROM    dbo.[@SAPCP_CONFIGGRUPO]
                                WHERE   U_Usuario = '" + _Company.UserName.ToString() + "'");

                return _oRec.Fields.Item("U_Grupo").Value.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener el grupo *ObtenerGrupoSocioNegocio* : " + ex.Message);
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
        ///  Obtiene la clase de socio que se encuentra definido en la tabla
        /// </summary>
        /// <param name="_oForm">Forma activa</param>
        /// <param name="_Company">Objeto Company</param>
        /// <returns>Clase de socio de negocio</returns>
        public static string ObtenerClaseSocioNegocio(SAPbouiCOM.Form _oForm, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT TOP 1
                                        U_ClaseSocio
                                FROM    dbo.[@SAPCP_CONFIGGRUPO]
                                WHERE   U_Usuario = '" + _Company.UserName.ToString() + "'");

                return _oRec.Fields.Item("U_ClaseSocio").Value.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener la clase de socio de negocio *ObtenerClaseSocioNegocio* : " + ex.Message);
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
        /// Se llena el combobox de forma de pago
        /// </summary>
        /// <param name="cmbBox">Combobox a cargar</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <param name="_oForma">Formulario</param>
        /// <returns>Regresa el combobox de forma de pago cargado con datos</returns>
        public static SAPbouiCOM.ComboBox LlenarComboFormaPago(this SAPbouiCOM.ComboBox cmbBox, SAPbobsCOM.Company _Company, SAPbouiCOM.Form _oForma)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _oRec.DoQuery(@"SELECT  T1.FldValue,
                                        T1.Descr
                                FROM    dbo.CUFD T0
                                        LEFT JOIN dbo.UFD1 T1 ON T1.FieldID = T0.FieldID
                                                                    AND T0.TableID = T1.TableID
                                WHERE   T0.AliasID = 'FormaPago'
                                        AND T0.TableID = 'WTR1'");

                cmbBox.ValidValues.Add("0", "Todos");
                for (int i = 0; i < _oRec.RecordCount; i++)
                {
                    cmbBox.ValidValues.Add(_oRec.Fields.Item("FldValue").Value.ToString(), _oRec.Fields.Item("Descr").Value.ToString());
                    _oRec.MoveNext();
                }
                return cmbBox;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al cargar el combo de forma de pago *LlenarComboFormaPago* : " + ex.Message);
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
        /// Se llena el combobox de origen de la solicitud
        /// </summary>
        /// <param name="cmbBox">Combobox a cargar</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <param name="_oForma">Formulario</param>
        /// <returns>Regresa el combobox de origen de la solicitud cargado con datos</returns>
        public static SAPbouiCOM.ComboBox LlenarComboOrigenSolicitud(this SAPbouiCOM.ComboBox cmbBox, SAPbobsCOM.Company _Company, SAPbouiCOM.Form _oForma)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _oRec.DoQuery(@"SELECT  Code,
                                        Name
                                FROM    dbo.[@ORIGSOLICITUD]");

                cmbBox.ValidValues.Add("0", "Todos");
                for (int i = 0; i < _oRec.RecordCount; i++)
                {
                    cmbBox.ValidValues.Add(_oRec.Fields.Item("Code").Value.ToString(), _oRec.Fields.Item("Name").Value.ToString());
                    _oRec.MoveNext();
                }
                return cmbBox;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al cargar el combo de origen de solicitud *LlenarComboOrigenSolicitud* : " + ex.Message);
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
        /// Validar si el reporte ya fue impreso
        /// </summary>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>true / false</returns>
        public static bool ValidarImpresionCorteSolicitudes(SAPbobsCOM.Company _Company, string CodigoOficina)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT TOP 1
                                        U_Oficina
                                FROM    dbo.[@SAPCP_CONFIGCORTSOL]
                                WHERE   U_Oficina = '" + CodigoOficina + "' " +
                                        "AND U_FechaImpresion = CONVERT(VARCHAR(10), GETDATE(), 103)");
                return _oRec.RecordCount == 0 ? true : false;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar si se imprimió el reporte de corte de solicitudes *ValidarImpresionCorteSolicitudes* : " + ex.Message);
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
        /// Obtiene el prefijo de la serie
        /// </summary>
        /// <param name="prefijo">Prefijo de la serie</param>
        /// <param name="plan">Plan del paquete</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>Regresa el prefijo</returns>
        public static string ObtenerPrefijoSerie(string prefijo, string plan, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  U_Prefijo_Contr
                                FROM    dbo.[@COMISIONES]
                                WHERE   U_Codigo_Plan = '" + plan + "' " +
                                        "AND U_Prefijo_Sol = '" + prefijo + "' ");
                return _oRec.Fields.Item("U_Prefijo_Contr").Value.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar si se imprimió el reporte de corte de solicitudes *ObtenerPrefijoSerie* : " + ex.Message);
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
        /// Obtiene el código del impuesto
        /// </summary>
        /// <param name="itemCode">Código del articulo</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>Código del impuesto</returns>
        public static string ObtenerCodigoImpuesto(string itemCode, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"   SELECT  U_ImpuestoVentas
                                   FROM    OITM
                                   WHERE   ItemCode = '" + itemCode + "'");
                return _oRec.Fields.Item("U_ImpuestoVentas").Value.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener el código del impuesto *ObtenerCodigoImpuesto* : " + ex.Message);
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
        /// Obtiene el almacén del articulo
        /// </summary>
        /// <param name="solicitudInterna">Numero de serie</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>Código del almacén</returns>
        public static string ObtenerAlmacen(SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  WhsCode AS OficinaContratos
                                FROM    dbo.OWHS
                                WHERE   WhsName = 'OFICINA ELABORACION DE CONTRATOS'");
                return _oRec.Fields.Item("OficinaContratos").Value.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener el almacén de contratos *ObtenerAlmacen* : " + ex.Message);
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
        /// Obtiene el número del sistema de la serie
        /// </summary>
        /// <param name="solicitudInterna">Numero de serie</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>Numero del sistema</returns>
        public static int ObtenerNumeroSistema(string solicitudInterna, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT TOP 1  SysNumber FROM dbo.OSRN WHERE DistNumber = '" + solicitudInterna + "'");
                return _oRec.Fields.Item("SysNumber").Value;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener el numero del sistema *ObtenerNumeroSistema* : " + ex.Message);
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
        /// Obtiene el nombre de la empresa del plan
        /// </summary>
        /// <param name="prefijo">Prefijo de la serie</param>
        /// <param name="plan">Plan del paquete</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>El nombre de la empresa</returns>
        public static string ObtenerEmpresa(string prefijo, string plan, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  U_Empresa
                                FROM    dbo.[@COMISIONES]
                                WHERE   U_Codigo_Plan = '" + plan + "' " +
                                        "AND U_Prefijo_Sol = '" + prefijo + "' ");
                return _oRec.Fields.Item("U_Empresa").Value.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener empresa *ObtenerEmpresa* : " + ex.Message);
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
        ///  Obtiene el código y nombre de la secretaria
        /// </summary>
        /// <param name="_oForm">Forma activa</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>Grupo</returns>
        public static string ObtenerSecretaria(SAPbobsCOM.Company _Company, string campo)
        {
            try
            {
                //T0.U_codigo_secretaria AS CodigoSecretaria,
                //T0.U_nombre_secretaria AS NombreSecretaria
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _oRec.DoQuery(@"SELECT TOP 1
                                         " + campo + " AS Campo " +
                                "FROM    dbo.[@SECRETARIAS] T0 " +
                                        "INNER JOIN dbo.OHEM T1 ON T1.firstName = T0.U_codigo_secretaria " +
                                        "INNER JOIN dbo.OUSR T2 ON T2.USERID = T1.userId " +
                                "WHERE   T2.USER_CODE = '" + _Company.UserName.ToString() + "'");

                return _oRec.Fields.Item("Campo").Value.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener la secretaria *ObtenerSecretaria* : " + ex.Message);
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
        /// Obtiene el almacén de la persona registrada
        /// </summary>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>Almacén</returns>
        public static string ObtenerAlmacenOficina(SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                _oRec.DoQuery(@"DECLARE @CADENA VARCHAR(MAX);
                                SELECT  @CADENA = COALESCE(@CADENA + ',', '') + U_Oficina
                                FROM    dbo.[@CAMBIO_OFICINA]
                                WHERE   U_Usuario = '" + _Company.UserName.ToString() + "'" +
                                " SELECT  @CADENA [AlmacenOrigen];");

                return _oRec.Fields.Item("AlmacenOrigen").Value.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener el almacén de la oficina *ObtenerAlmacenOficina* : " + ex.Message);
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
        /// Obtiene los datos del contrato para mostrar en traspasos
        /// </summary>
        /// <param name="contrato">Numero de contrato</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns></returns>
        public static DatosTraspasos ObtenerDatosTraspasos(string contrato, SAPbobsCOM.Company _Company)
        {
            try
            {
                 DatosTraspasos datTraspaso = null;
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                datTraspaso = new DatosTraspasos();

                _oRec.DoQuery(@"SELECT  T1.DocEntry AS Docentry,
		                                T0.CardName AS nombreSocio,
                                        T3.U_Descripcion_Plan AS 'Plan',
                                        T1.DocTotal AS CostoPlan,
                                        T0.Balance AS Saldo,
                                        T3.U_Empresa AS Empresa
                                FROM    dbo.OCRD T0
                                        INNER JOIN dbo.OINV T1 ON T1.CardCode = T0.CardCode
                                                                    AND T1.DocStatus = 'O'
                                        INNER JOIN dbo.INV1 T2 ON T2.DocEntry = T1.DocEntry
                                        INNER JOIN dbo.[@COMISIONES] T3 ON T3.U_Codigo_Plan = T2.ItemCode
                                                                            AND T3.U_Prefijo_Contr = SUBSTRING(T0.CardCode, 1, 3)
                                                                            AND T3.U_Descripcion_Plan = T0.U_Dsciption
                                                                            AND T3.U_Prefijo_Sol = SUBSTRING(T0.U_SolicitudInt,1,6)
                                WHERE   T0.CardCode = '" + contrato + "' ORDER BY T1.DocDate ASC");

                datTraspaso.nombreSN = _oRec.Fields.Item("nombreSocio").Value;
                datTraspaso.plan = _oRec.Fields.Item("Plan").Value;
                datTraspaso.costoPlan = _oRec.Fields.Item("CostoPlan").Value;
                datTraspaso.saldo = _oRec.Fields.Item("Saldo").Value;
                datTraspaso.empresa = _oRec.Fields.Item("Empresa").Value;
                datTraspaso.DocEntryFactura = _oRec.Fields.Item("Docentry").Value;
                return datTraspaso;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener datos detalle de comisiones *ObtenerDatosTraspasos* : " + ex.Message);
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
        /// Obtiene los montos que le corresponden al cobrador definidos por puestos
        /// </summary>
        /// <param name="empresa">Nombre de la empresa Apoyo / Cooperativa</param>
        /// <param name="codigoCobrador">Código del cobrador</param>
        /// <param name="codigoPlan">Código del plan</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>Datos de la clase</returns>
        public static DatosDetalleComsiones ObtenerDatosDetalleComisiones(string empresa, string codigoCobrador, string codigoPlan,string solicitudInt, SAPbobsCOM.Company _Company)
        {
        regresar:
            try
            {           
                DatosDetalleComsiones datos = null;
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                datos = new DatosDetalleComsiones();

                _oRec.DoQuery(@"SELECT  --T0.U_Codigo_Cobrador,
                                        --T0.U_Cobrador,
                                        --T0.U_Empresa,
                                        --T0.U_Codigo_Plan,
                                        --T0.U_Descripcion_Plan,
                                        T0.U_CodigoRecomendado AS CodigoRecomendado,  
                                        REPLACE(T0.U_Nom_Recomendado, '''','' ) AS NomRecomendado,                                                                              
                                        ISNULL(T0.U_Recomendado,0) AS Recomendado,
                                        ISNULL(T0.U_Asis_Social,0) AS AsisSocial,
                                        ISNULL(T0.U_Bono,0) AS Bono,
                                        T0.U_CodigoLider AS CodigoLider,
                                        REPLACE(T0.U_Nom_Lider, '''','' ) AS NomLider,
                                        ISNULL(T0.U_Lider,0) AS Lider,
                                        T0.U_CodigoSupervisor AS CodigoSupervisor,
                                        REPLACE(T0.U_Nom_Supervisor, '''','' ) AS NomSupervisor,
                                        ISNULL(T0.U_Supervisor,0) AS Supervisor,
                                        T0.U_CodigoCoordinador AS CodigoCoordinador,
                                        REPLACE(T0.U_Nom_Coordinador, '''','' ) AS NomCoordinador,
                                        ISNULL(T0.U_Coordinador,0) AS Coordinador,
                                        ISNULL(T0.U_BonoCoordinador,0) AS BonoCoordinador,
                                        T0.U_CodigoCoordinador2 AS CodigoCoordinador2,
                                        REPLACE(T0.U_Nom_Coordinador2, '''','' ) AS NomCoordinador2,
                                        ISNULL(T0.U_Coordinador2,0) AS Coordinador2,
                                        ISNULL(T0.U_BonoCoordinador2,0) AS BonoCoordinador2,
                                        T0.U_CodigoGerente AS CodigoGerente,
                                        REPLACE(T0.U_Nom_Gerente, '''','' ) AS NomGerente,
                                        ISNULL(T0.U_Gerente,0) AS Gerente,
                                        ISNULL(T0.U_Fideicomiso,0) AS Fideicomiso,
                                        ISNULL(T1.U_Inv_Inicial,0) AS InvInicial
                                FROM    dbo.[@DETALLE_COMISION] T0
                                        INNER JOIN dbo.[@COMISIONES] T1 ON T1.U_Empresa = T0.U_Empresa
                                                                           AND T1.U_Codigo_Plan = T0.U_Codigo_Plan
                                WHERE   T0.U_Empresa = '" + empresa + "' " +
                                        "AND T0.U_Codigo_Cobrador = '" + codigoCobrador + "' " +
                                        "AND T0.U_Codigo_Plan = '" + codigoPlan + "' " +
                                        "AND T1.U_Prefijo_Sol = '" + solicitudInt + "'");

                datos.codigoRecomendado = _oRec.Fields.Item("CodigoRecomendado").Value;
                datos.nombreRecomendado = _oRec.Fields.Item("NomRecomendado").Value;
                datos.codigoLider = _oRec.Fields.Item("CodigoLider").Value;
                datos.nombreLider = _oRec.Fields.Item("NomLider").Value;
                datos.codigoSupervisor = _oRec.Fields.Item("CodigoSupervisor").Value;
                datos.nombreSupervisor = _oRec.Fields.Item("NomSupervisor").Value;
                datos.codigoCoordinador = _oRec.Fields.Item("CodigoCoordinador").Value;
                datos.nombreCoordinador = _oRec.Fields.Item("NomCoordinador").Value;
                datos.codigoCoordinador2 = _oRec.Fields.Item("CodigoCoordinador2").Value;
                datos.nombreCoordinador2 = _oRec.Fields.Item("NomCoordinador2").Value;
                datos.codigoGerente = _oRec.Fields.Item("CodigoGerente").Value;
                datos.nombreGerente = _oRec.Fields.Item("NomGerente").Value;
                datos.montoAsistenteSocial = _oRec.Fields.Item("AsisSocial").Value;
                datos.montoRecomendado = _oRec.Fields.Item("Recomendado").Value;
                datos.montoLider = _oRec.Fields.Item("Lider").Value;
                datos.montoSupervisor = _oRec.Fields.Item("Supervisor").Value;
                datos.montoCoordinador = _oRec.Fields.Item("Coordinador").Value;
                datos.montoBonoCoordinador = _oRec.Fields.Item("BonoCoordinador").Value;
                datos.montoCoordinador2 = _oRec.Fields.Item("Coordinador2").Value;
                datos.montoBonoCoordinador2 = _oRec.Fields.Item("BonoCoordinador2").Value;
                datos.montoGerente = _oRec.Fields.Item("Gerente").Value;
                datos.montoFideicomiso = _oRec.Fields.Item("Fideicomiso").Value;
                datos.montoBono = _oRec.Fields.Item("Bono").Value;
                datos.montoInvInicial = _oRec.Fields.Item("InvInicial").Value;

                return datos;
            }
            catch (Exception ex)
            {                     
                //throw new Exception("Error al obtener datos detalle de comisiones *ObtenerDatosDetalleComisiones* : " + ex.Message);
                goto regresar;                
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
        /// Obtiene los datos de la transferencia solicitada
        /// </summary>
        /// <param name="traspasoRel">Numero de traspaso</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>Datos de la clase</returns>
        public static DatosTransferencia ObtenerDatosTransferencia(string SolicitudInterna, SAPbobsCOM.Company _Company)
        {
            try
            {
                 DatosTransferencia datos = null;
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                datos = new DatosTransferencia();

                _oRec.DoQuery(@"SELECT TOP 1
                                        T1.DocEntry ,
                                        T0.U_CodPromotor ,
                                        REPLACE(T0.U_NombrePromotor, '''', '') AS U_NombrePromotor
                                FROM    dbo.WTR1 T0
                                        LEFT JOIN dbo.OWTR T1 ON T1.DocEntry = T0.DocEntry
                                WHERE   T1.U_TipoMov IN ( 'OFICINAS - CONTRATOS', 'PROMOTOR - ADMON CONTRATOS' )
                                        AND T0.U_FormaPago IS NOT NULL
                                        AND T0.U_OrigenSolicitud IS NOT NULL
                                        AND T0.U_Importe IS NOT NULL
                                        AND T0.U_InvInicial IS NOT NULL
                                        AND T0.U_Serie = '" + SolicitudInterna + "' " +
                                        "AND T1.CANCELED <> 'Y' " +
                                        "AND T1.DataSource <> 'N' " +
                                "ORDER BY T0.DocEntry DESC");

                datos.docEntryTransferencia = _oRec.Fields.Item("DocEntry").Value;
                datos.codigoAsistente = _oRec.Fields.Item("U_CodPromotor").Value;
                datos.nombreAsistente = _oRec.Fields.Item("U_NombrePromotor").Value;

                return datos;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener datos de la transferencia *ObtenerDatosTransferencia* : " + ex.Message);
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

        public static DatosSolicitud ObtenerDatosSolicitud(string SolicitudInterna, SAPbobsCOM.Company _Company)
        {
            try
            {
                DatosSolicitud datos = null;
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                datos = new DatosSolicitud();

                _oRec.DoQuery(@"SELECT TOP 1
                                        T4.U_Codigo_Plan AS codigoPlan ,
                                        T4.U_Descripcion_Plan AS nombrePlan ,
                                        T3.firstName AS codigoAsistente ,
                                        CONCAT(T3.middleName, ' ', T3.lastName) AS nombreAsistente ,
                                        T4.U_Prefijo_Contr AS prefijoPlan ,
                                        T4.U_Empresa AS empresa
                                FROM    dbo.OWTR T0
                                        INNER JOIN dbo.WTR1 T1 ON T0.DocEntry = T1.DocEntry
                                        INNER JOIN dbo.OSLP T2 ON T0.SlpCode = T2.SlpCode
                                        INNER JOIN dbo.OHEM T3 ON T2.SlpCode = T3.salesPrson
                                        INNER JOIN dbo.[@COMISIONES] T4 ON T4.U_Prefijo_Sol = SUBSTRING(T1.U_Serie,
                                                                                              1, 6)
                                WHERE   T1.[U_Serie] = '" + SolicitudInterna + "' " +
                                        "AND T0.CANCELED <> 'Y' " +
                                        "AND T0.DataSource <> 'N' " +
                                        "AND (T0.U_TipoMov = 'OFICINAS - PROMOTORES' OR T0.U_TipoMov = 'ADMON CONTRATOS - PROMOTOR') ORDER BY T0.DocEntry DESC");

                datos.codigoPlan = _oRec.Fields.Item("codigoPlan").Value;
                datos.nombrePlan = _oRec.Fields.Item("nombrePlan").Value;
                datos.codigoAsistente = _oRec.Fields.Item("codigoAsistente").Value;
                datos.nombreAsistente = _oRec.Fields.Item("nombreAsistente").Value;
                datos.prefijoPlan = _oRec.Fields.Item("prefijoPlan").Value;
                datos.empresa = _oRec.Fields.Item("empresa").Value;
                return datos;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener datos de la solicitud *ObtenerDatosSolicitud* : " + ex.Message);
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
        /// Una vez que se genere la factura y el pago correctamente el sistema insertara el docEntry de la factura en la tabla de Calculo de comisiones
        /// </summary>
        /// <param name="contrato">Numero contrato</param>
        /// <param name="docEntryFactura">DocEntry que se insertara en la tabla de calculo de comisiones</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        public static void InsertarDocentryFacturaCalculoComisiones(string contrato, string docEntryFactura, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _oRec.DoQuery(@"UPDATE  dbo.[@CALCULO_COMISIONES]
                                SET     U_DocEntryFactura = '" + docEntryFactura + "' " +
                                "WHERE   U_Contrato = '" + contrato + "'");
            }
            catch (Exception ex)
            {
                throw new Exception("Error al insertar el DocEntry de la factura *InsertarDocentryFacturaCalculoComisiones* : " + ex.Message);
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
        /// Valida si el plan existe en la base de datos
        /// </summary>
        /// <param name="contrato">contrato</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>True/False</returns>
        public static bool ValidarSiExistePlan(string contrato, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _oRec.DoQuery(@"SELECT  *
                                FROM    dbo.[@CALCULO_COMISIONES]
                                WHERE   U_Contrato = '" + contrato + "'");

                return _oRec.RecordCount == 0 ? true : false;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar si existe el plan *ValidarSiExistePlan* : " + ex.Message);
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
        /// Valida si el cobrador se encuentra 
        /// </summary>
        /// <param name="solicitudInt">Número de la solicitud</param>
        /// <param name="_Company"></param>
        /// <returns></returns>
        internal static bool ValidarSiEstaCobrador(string contrato, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _oRec.DoQuery(@"SELECT   U_Codigo_Cobrador
                               FROM     dbo.[@CALCULO_COMISIONES]
                               WHERE    U_Contrato = '" + contrato + "'");

                return string.IsNullOrEmpty(_oRec.Fields.Item("U_Codigo_Cobrador").Value) ? true : false;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar si existe el cobrador en el calculo de las comisiones *ValidarSiEstaCobrador* : " + ex.Message);
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
        /// Obtiene el costo del plan para generar la factura
        /// </summary>
        /// <param name="solicitudInterna">Numero de solicitud interna</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>Costo del plan</returns>
        public static double ObtenerCostoPaquete(string solicitudInternaPrefij, string codigoPlan, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  U_Costo
                                FROM    dbo.[@COMISIONES]
                                WHERE   U_Codigo_Plan = '" + codigoPlan + "' " +
                                        "AND U_Prefijo_Sol = '" + solicitudInternaPrefij + "' ");
                return Convert.ToDouble(_oRec.Fields.Item("U_Costo").Value.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener el monto del plan *ObtenerCostoPaquete* : " + ex.Message);
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
        /// Obtiene el saldo del contrato
        /// </summary>
        /// <param name="cardCode">Código de cliente</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>Monto del saldo de cuenta</returns>
        public static double ObtenerSaldoDeCuenta(string cardCode, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  Balance
                                FROM    dbo.OCRD
                                WHERE   CardCode = '" + cardCode + "' ");
                return Convert.ToDouble(_oRec.Fields.Item("Balance").Value.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener el saldo de cuenta *ObtenerSaldoDeCuenta* : " + ex.Message);
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
        /// Obtiene el esquema de Pago del asistente
        /// </summary>
        /// <param name="CodigoAsistente">Código del asistente</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>SUELDO/COMISION</returns>
        public static string ObtenerEsquemaPago(string CodigoAsistente, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT  CASE WHEN U_Esquema_pago LIKE '%SUELDO%' THEN 'SUELDO'
                                             ELSE U_Esquema_pago
                                        END AS Esquema
                                FROM    dbo.OHEM
                                WHERE   firstName = '" + CodigoAsistente + "'");
                return _oRec.Fields.Item("Esquema").Value.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener el esquema de pago *ObtenerEsquemaPago* : " + ex.Message);
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
        /// Obtiene el numero de las ayudas del asistente
        /// </summary>
        /// <param name="CodigoAsistente">Código del asistente</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <returns>Número de ayudas</returns>
        public static int ObtenerAyudasAsistente(string CodigoAsistente, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"DECLARE @CODIGO_ASISTENTE NVARCHAR(50)
                                SET @CODIGO_ASISTENTE = '" + CodigoAsistente + "'" +
                                "BEGIN " +
                                    "IF NOT EXISTS ( SELECT  * " +
                                                    "FROM    dbo.[@AYUDAS] " +
                                                    "WHERE   U_CodigoAsistente = @CODIGO_ASISTENTE ) " + 
                                        "BEGIN " +
                                            "INSERT  INTO dbo.[@AYUDAS] " +
                                                    "( Code , " +
                                                      "Name , " +
                                                      "U_CodigoAsistente , " +
                                                      "U_NumeroAyuda , " +
                                                      "U_NumeroAyudaAnt " +
		                                            ") " +
                                            "VALUES  ( ( SELECT  ISNULL(MAX(CAST(Code AS INT)), 0) + 1 AS DocEntry " +
                                                        "FROM    dbo.[@AYUDAS] WITH ( UPDLOCK ) " +
                                                      ") , " +
                                                      "( SELECT  ISNULL(MAX(CAST(Code AS INT)), 0) + 1 AS DocEntry " +
                                                        "FROM    dbo.[@AYUDAS] WITH ( UPDLOCK ) " +
                                                      ") , " +
                                                      "@CODIGO_ASISTENTE , " +
                                                      "1 , " +
                                                      "0   " +
		                                            ") " +
                                        "END " +
                                    "ELSE  " +
                                        "BEGIN " +
                                            "UPDATE  dbo.[@AYUDAS] " +
                                            "SET     U_NumeroAyuda = U_NumeroAyuda + 1 " +
                                            "WHERE   U_CodigoAsistente = @CODIGO_ASISTENTE " +
                                        "END " +
                                    "SELECT  U_NumeroAyuda AS NumeroAyudas " +
                                    "FROM    dbo.[@AYUDAS] " +
                                    "WHERE   U_CodigoAsistente = @CODIGO_ASISTENTE  " +			        
                                "END");
                return Convert.ToInt32(_oRec.Fields.Item("NumeroAyudas").Value.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener las ayudas del asistentes *ObtenerAyudasAsistente* : " + ex.Message);
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
        /// Valida si la solicitud ya se encuentra ingresada en algún documento
        /// </summary>
        /// <param name="solicitud">Número de solicitud proporcionada por el asistente</param>
        /// <param name="_Company"></param>
        /// <returns></returns>
        public static bool ValidarSiExisteSolicitud(string solicitud, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _oRec.DoQuery(@"SELECT  CardCode
                                FROM    OCRD
                                WHERE   U_SolicitudInt = '" + solicitud + "'");

                return _oRec.RecordCount == 0 ? true : false;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al validar si existe la solicitud *ValidarSiExisteSolicitud* : " + ex.Message);
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
        /// Valida si el contrato existe y si tiene mas contratos
        /// </summary>
        /// <param name="rfc">RFC del socio de negocio</param>
        /// <returns>Mensaje de los contratos creados o vacío</returns>
        public static string ValidarSiExisteSocioDeNegocio(string nombre, SAPbobsCOM.Company _Company)
        {
            try
            {
                List<string> contratos = null;
                string msgError = null;
                _oRec = null;
                contratos = new List<string>();
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//                _oRec.DoQuery(@"SELECT  T0.CardCode
//                                FROM    OCRD T0
//                                        LEFT JOIN dbo.CRD1 T1 ON T0.CardCode = T1.CardCode
//                                WHERE   CardName = '" + nombre + "' " +
//                                        "AND ISNULL(T1.Block, '') = '" + colonia + "'");

//                _oRec.DoQuery(@"SELECT  T0.CardCode AS CardCode
//                                    FROM    OCRD T0
//                                            LEFT JOIN dbo.CRD1 T1 ON T0.CardCode = T1.CardCode
//                                                                        AND T1.Address = 'DIRECCION 1'
//                                    WHERE   CardName = '" + nombre + "' " +
//                                            "AND ISNULL(T1.Block, '') = '" + colonia + "' " +
//                                            "AND ISNULL(T1.City, '') = '" + municipio + "' " +
//                                    "UNION " +
//                                    "SELECT  U_Contrato AS Cardcode " +
//                                    "FROM    dbo.[@CONTRATOS_PABS] " +
//                                    "WHERE   U_Nombre = '" + nombre + "' " +
//                                            "AND U_Colonia = '" + colonia + "' " +
//                                            "AND U_Municipio = '" + municipio + "' ");

//                _oRec.DoQuery(@"SELECT  CONCAT(T0.CardCode,',',T2.U_Contrato) AS CardCode        
//                                                    FROM    OCRD T0
//                                                            LEFT JOIN dbo.CRD1 T1 ON T0.CardCode = T1.CardCode
//                                                                                        AND T1.Address = 'DIRECCION 1'
//                                                            LEFT JOIN dbo.[@CONTRATOS_PABS] T2 ON T0.CardName = T2.U_Nombre
//                                                                                                    AND T1.Block = T2.U_Colonia
//                                                                                                    AND T1.City = T2.U_Municipio
//                                                    WHERE   CardName = '" + nombre + "' " +
//                                                         "AND ISNULL(T1.Block, '') = '" + colonia + "' " +
//                                                         "AND ISNULL(T1.City, '') = '" + municipio + "' ");

                _oRec.DoQuery(@"SELECT  T0.CardCode AS CardCode
                                    FROM    OCRD T0                                        
                                    WHERE   CardName = '" + nombre + "' " +                                   
                                "UNION " +
                                "SELECT  U_Contrato AS Cardcode " +
                                "FROM    dbo.[@CONTRATOS_PABS] " +
                                "WHERE   U_Nombre = '" + nombre + "' ");


                if (_oRec.RecordCount > 1)
                {
                    for (int i = 0; i < _oRec.RecordCount; i++)
                    {
                        contratos.Add(_oRec.Fields.Item("CardCode").Value.ToString());
                        _oRec.MoveNext();
                    }
                    var ListaContratos = string.Join(Environment.NewLine, contratos.Select(s => s.ToString()));
                    msgError = "Ya existen contratos con este nombre: " + nombre + Environment.NewLine + ListaContratos;
                }
                else
                {
                    msgError = "";
                }

                return msgError;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener el saldo de cuenta *ObtenerSaldoDeCuenta* : " + ex.Message);
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
        /// Carga la clase del xml 
        /// </summary>
        /// <returns>Clase de configuraciones</returns>
        public static bool CargarConfiguraciones()
        {
            try
            {
                Configuracion = new CONFIGURACION();
                Configuracion = Configuracion.CargarXML("Configuracion.xml");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"No se pudo cargar el archivo de configuraciones *ObtenerSaldoDeCuenta* : " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Serializa el xml de configuración
        /// </summary>
        /// <param name="obj">Objeto de la clase</param>
        /// <param name="NombreXml">Nombre del xml</param>
        /// <returns></returns>
        public static CONFIGURACION CargarXML(this CONFIGURACION obj, string NombreXml)
        {
            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(CONFIGURACION));
            try
            {
                if (System.IO.File.Exists(NombreXml))
                {
                    System.IO.StreamReader sr = new System.IO.StreamReader(NombreXml);
                    obj = (CONFIGURACION)serializer.Deserialize(sr);
                    sr.Close();
                }
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show(@"No se puede cargar el archivo *CargarXML* : " + ex.Message);
            }
            return obj;
        }

        /// <summary>
        /// Obtiene las oficinas por secretaria
        /// </summary>
        /// <param name="cmbBox">ComboBox a cargar</param>
        /// <param name="_Company">Company es el objeto de la API DI principal que representa 
        ///                         una única base de datos de SAP Business One empresa
        ///                         Este objeto le permite conectarse a la base de datos de la empresa y 
        ///                         crear objetos de negocio para su uso con la base de datos de la empresa</param>
        /// <param name="_oForma">Formulario</param>
        /// <returns>Oficinas</returns>
        public static SAPbouiCOM.ComboBox LlenarCargarOficinas(this SAPbouiCOM.ComboBox cmbBox, SAPbobsCOM.Company _Company, SAPbouiCOM.Form _oForma)
        {
            try
            {
                _oForma.Freeze(true);
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _oRec.DoQuery(@"SELECT  T0.U_Oficina AS CodigoOficina ,
                                        T1.WhsName AS NombreOficina
                                FROM    dbo.[@CAMBIO_OFICINA] T0
                                        INNER JOIN dbo.OWHS T1 ON T0.U_Oficina = T1.WhsCode
                                WHERE   T0.U_Usuario = '" + _Company.UserName.ToString() + "' " +
                                " ORDER BY T0.U_Oficina ASC");

                for(int oficina = 0; oficina < _oRec.RecordCount; oficina ++)
                {
                    if (_oRec.Fields.Item("NombreOficina").Value.ToString().Length > 80)
                    {
                        cmbBox.ValidValues.Add(_oRec.Fields.Item("CodigoOficina").Value.ToString(), _oRec.Fields.Item("NombreOficina").Value.ToString().Substring(0, 80));
                    }
                    else
                    {
                        cmbBox.ValidValues.Add(_oRec.Fields.Item("CodigoOficina").Value.ToString(), _oRec.Fields.Item("NombreOficina").Value.ToString());
                    }
                    _oRec.MoveNext();
                }
                cmbBox.ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription;
                SAPbouiCOM.Item item = cmbBox.Item;
                item.DisplayDesc = true;
                
                return cmbBox;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al cargar oficinas *LlenarCargarOficinas*");
            }
            finally
            {
                _oForma.Freeze(false);
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


        public static void ActualizarEsquemaComision(string contrato,string esquema, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _oRec.DoQuery(@"UPDATE dbo.OCRD SET U_Esquema_pago = '" + esquema + "' WHERE CardCode = '" + contrato + "'");
            }
            catch (Exception ex)
            {
                throw new Exception("Error al actualizar esquema *ActualizarEsquemaComision* : " + ex.Message);
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


        public static int ObtenerNumeroContratos(string codigoAsistenteContrato, SAPbobsCOM.Company _Company)
        {
            try
            {
                _oRec = null;
                _oRec = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                _oRec.DoQuery(@"SELECT TOP 1
                                        T1.U_CantidadContratos AS CantidadContratos
                                FROM    dbo.OHEM T0
                                        INNER JOIN dbo.[@CONFIG_ESQUEMAS] T1 ON T0.U_Esquema_pago = T1.U_Esquema
                                WHERE   T0.firstName = '" + codigoAsistenteContrato + "'");
                return Convert.ToInt32(_oRec.Fields.Item("CantidadContratos").Value.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception("Error al obtener el numero de contratos *ObtenerNumeroContratos* : " + ex.Message);
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

        #region CLASES

        /// <summary>
        /// Obtiene los datos de la tabla @DETALLE_COMISION y @COMISIONES
        /// </summary>
        public class DatosDetalleComsiones
        {
            #region NombresCodigosPuestos
            public string codigoRecomendado { get; set; }
            public string nombreRecomendado { get; set; }
            public string codigoLider { get; set; }
            public string nombreLider { get; set; }
            public string codigoSupervisor { get; set; }
            public string nombreSupervisor { get; set; }
            public string codigoCoordinador { get; set; }
            public string nombreCoordinador { get; set; }
            public string codigoCoordinador2 { get; set; }
            public string nombreCoordinador2 { get; set; }
            public string codigoGerente { get; set; }
            public string nombreGerente { get; set; }

            #endregion

            #region montosPuestos

            public double montoAsistenteSocial { get; set; }
            public double montoRecomendado { get; set; }
            public double montoLider { get; set; }
            public double montoSupervisor { get; set; }
            public double montoCoordinador { get; set; }
            public double montoCoordinador2 { get; set; }
            public double montoGerente { get; set; }
            public double montoFideicomiso { get; set; }
            public double montoBono { get; set; }
            public double montoInvInicial { get; set; }
            public double montoBonoCoordinador { get; set; }
            public double montoBonoCoordinador2 { get; set; }

            #endregion
        }

        /// <summary>
        /// Obtiene los datos de la transferencia relacionada con el pre - contrato
        /// </summary>
        public class DatosTransferencia
        {
            public int docEntryTransferencia { get; set; }
            public string codigoAsistente { get; set; }
            public string nombreAsistente { get; set; }
        }

        /// <summary>
        /// Obtiene los datos de los traspasos
        /// </summary>
        public class DatosTraspasos
        {
            public int DocEntryFactura { get; set; }
            public string nombreSN { get; set; }
            public string plan { get; set; }
            public double costoPlan { get; set; }
            public double saldo { get; set; }
            public string empresa { get; set; }
        }

        public class DatosSolicitud
        {
            public string codigoPlan { get; set; }
            public string nombrePlan { get; set; }
            public string codigoAsistente { get; set; }
            public string nombreAsistente { get; set; }
            public string prefijoPlan { get; set; }
            public string empresa { get; set; }
        }

        #endregion







     
    }
}
