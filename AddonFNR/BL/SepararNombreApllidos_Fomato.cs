using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddonFNR.BL
{
    public class SepararNombreApllidos_Fomato
    {
        #region METODOS

        /// <summary>
        /// Convierte el texto a mayúsculas
        /// </summary>
        /// <param name="nombreCompleto">Nombre a cambiar</param>
        /// <returns>Nombre convertido</returns>
        public static string FormatoMayusculas(string nombreCompleto)
        {
            try
            {

                nombreCompleto = nombreCompleto.ToUpper();
                nombreCompleto = nombreCompleto.Replace("Á", "A");
                nombreCompleto = nombreCompleto.Replace("É", "E");
                nombreCompleto = nombreCompleto.Replace("Í", "I");
                nombreCompleto = nombreCompleto.Replace("Ó", "O");
                nombreCompleto = nombreCompleto.Replace("Ú", "U");

                return nombreCompleto;

            }
            catch (Exception ex)
            {
                throw new Exception("Error en FormatoTextoRFC: " + ex.Message);
            }
        }

        /// <summary>
        /// Remueve palabras y da formato al nombre
        /// </summary>
        /// <param name="nombreCompleto">Nombre a valorar</param>
        /// <returns>nombre convertido</returns>
        public static string FormatoTextoRFC(string nombreCompleto)
        {
            try
            {

                nombreCompleto = nombreCompleto.ToUpper();
                nombreCompleto = nombreCompleto.Replace("Á", "A");
                nombreCompleto = nombreCompleto.Replace("É", "E");
                nombreCompleto = nombreCompleto.Replace("Í", "I");
                nombreCompleto = nombreCompleto.Replace("Ó", "O");
                nombreCompleto = nombreCompleto.Replace("Ú", "U");

                nombreCompleto = RemoverPalabras(nombreCompleto);
                nombreCompleto = RemoverNombres(nombreCompleto);
                nombreCompleto = SepararApellidos(nombreCompleto);

                return nombreCompleto;

            }
            catch (Exception ex)
            {
                throw new Exception("Error en FormatoTextoRFC: " + ex.Message);
            }
        }

        /// <summary>
        /// Remueve las palabras incorrectas
        /// </summary>
        /// <param name="nombre">Nombre a valorar</param>
        /// <returns>El nombre convertido</returns>
        private static string RemoverPalabras(string nombre)
        {
            object[] palabras = { " PARA ", " AND ", " CON ", " DEL ", " LAS ", " LOS ", " MAC ", " POR ", " SUS ", " THE ", " VAN ", 
                                    " VON ", " AL ", " DE ", " EL ", " EN ", " LA ", " MC ", " MI ", " OF ", " A ", " E ", " Y " };
            int i;

            nombre = (" " + nombre);
            for (i = palabras.GetLowerBound(0); (i <= palabras.Length - 1); i++)
            {
                nombre = nombre.Replace(palabras[i].ToString(), " ");
            }
            return nombre.Trim();
        }

        /// <summary>
        /// Remueve nombres no validos
        /// </summary>
        /// <param name="nombre">Nombre a valorar</param>
        /// <returns>El nombre convertido</returns>
        private static string RemoverNombres(string nombre)
        {
            object[] nombres = { " MA. ", " MA ", " J. ", " J " };
            int i;

            if (((nombre.IndexOf(" ") + 1)
                        > 0))
            {
                nombre = (" " + nombre);
                for (i = nombres.GetLowerBound(0); (i <= nombres.Length - 1); i++)
                {
                    nombre = nombre.Replace(nombres[i].ToString(), " ");
                }
            }
            return nombre.Trim();
        }

        /// <summary>
        /// Separa los apellidos
        /// </summary>
        /// <param name="nombre">Nombre completo con apellidos</param>
        /// <returns>El nombre separado</returns>
        private static string SepararApellidos(string nombre)
        {
            string[] nombreArr;
            string nuevaCadena = null;
            int i = 0;

            nombre = nombre.Replace("   ", " ");
            nombre = nombre.Replace("  ", " ");
            // Dividir el nombre por palabras en un arreglo
            nombreArr = nombre.Trim(' ').Split();

            // Analizar cada palabra dentro del arreglo
            //for (i = 0; (i <= UBound(nombreArr)); i++) {

            for (i = 0; (i <= nombreArr.Length - 1); i++)
            {
                switch (nombreArr[i].ToLower())
                {
                    case "de":
                    case "del":
                    case "la":
                    case "las":
                    case "los":
                    case "san":
                        nuevaCadena = (nuevaCadena
                                    + (nombreArr[i] + " "));
                        break;
                    default:
                        nuevaCadena = (nuevaCadena
                                    + (nombreArr[i] + "@"));
                        break;
                }
            }

            // Remover el �último carácter delimitador de la cadena
            if ((nuevaCadena.Substring((nuevaCadena.Length - 1)) == "@"))
            {
                nuevaCadena = nuevaCadena.Substring(0, (nuevaCadena.Length - 1));
            }

            return nuevaCadena;
        }

        /// <summary>
        /// Genera el RFC en base a los datos convertidos
        /// </summary>
        /// <param name="separarNombre">Nombre separado</param>
        /// <param name="dia">Dia</param>
        /// <param name="mes">Mes</param>
        /// <param name="year">Año</param>
        /// <returns>RFC armado</returns>        
        internal static string GenerarRFC(string[] separarNombre, string dia, string mes, string year)
        {
            try
            {

                //RFC.ObtieneRFC RFC = new RFC.ObtieneRFC();



                var rfc = "";
                if (separarNombre.Length == 3)
                {
                    //rfc = RFC.RFC13Pocisiones(separarNombre[2], separarNombre[1], separarNombre[0], year.Substring(2, 2) + "/" + mes.PadLeft(2, '0') + "/" + dia.PadLeft(2, '0'));


                    rfc = RfcFacil.RfcBuilder.ForNaturalPerson()
                                              .WithName(separarNombre[0])
                                              .WithFirstLastName(separarNombre[1])
                                              .WithSecondLastName(separarNombre[2])
                                              .WithDate(Convert.ToInt32(year), Convert.ToInt32(mes), Convert.ToInt32(dia))
                                              .Build().ToString();
                }
                else if (separarNombre.Length == 4)
                {
                    // rfc = RFC.RFC13Pocisiones(separarNombre[3], separarNombre[2], separarNombre[1] + separarNombre[0], year.Substring(2, 2) + "/" + mes.PadLeft(2, '0') + "/" + dia.PadLeft(2, '0'));
                    rfc = RfcFacil.RfcBuilder.ForNaturalPerson()
                                            .WithName(separarNombre[0] + " " + separarNombre[1])
                                            .WithFirstLastName(separarNombre[2])
                                            .WithSecondLastName(separarNombre[3])
                                            .WithDate(Convert.ToInt32(year), Convert.ToInt32(mes), Convert.ToInt32(dia))
                                            .Build().ToString();
                }
                else if (separarNombre.Length == 5)
                {

                    //rfc = RFC.RFC13Pocisiones(separarNombre[4], separarNombre[3], separarNombre[2] + separarNombre[1] + separarNombre[0], year.Substring(2, 2) + "/" + mes.PadLeft(2, '0') + "/" + dia.PadLeft(2, '0'));
                    rfc = RfcFacil.RfcBuilder.ForNaturalPerson()
                                            .WithName(separarNombre[0] + " " + separarNombre[1] + " " + separarNombre[2])
                                            .WithFirstLastName(separarNombre[3])
                                            .WithSecondLastName(separarNombre[4])
                                            .WithDate(Convert.ToInt32(year), Convert.ToInt32(mes), Convert.ToInt32(dia))
                                            .Build().ToString();
                }
                else if (separarNombre.Length == 6)
                {
                    rfc = RfcFacil.RfcBuilder.ForNaturalPerson()
                                          .WithName(separarNombre[0] + " " + separarNombre[1] + " " + separarNombre[2] + " " + separarNombre[3])
                                          .WithFirstLastName(separarNombre[4])
                                          .WithSecondLastName(separarNombre[5])
                                          .WithDate(Convert.ToInt32(year), Convert.ToInt32(mes), Convert.ToInt32(dia))
                                          .Build().ToString();
                }
                else if (separarNombre.Length == 7)
                {
                    rfc = RfcFacil.RfcBuilder.ForNaturalPerson()
                                          .WithName(separarNombre[0] + " " + separarNombre[1] + " " + separarNombre[2] + " " + separarNombre[3] + " " + separarNombre[4])
                                          .WithFirstLastName(separarNombre[5])
                                          .WithSecondLastName(separarNombre[6])
                                          .WithDate(Convert.ToInt32(year), Convert.ToInt32(mes), Convert.ToInt32(dia))
                                          .Build().ToString();
                }
                else if (separarNombre.Length == 8)
                {
                    rfc = RfcFacil.RfcBuilder.ForNaturalPerson()
                                          .WithName(separarNombre[0] + " " + separarNombre[1] + " " + separarNombre[2] + " " + separarNombre[3] + " " + separarNombre[4] + " " + separarNombre[5])
                                          .WithFirstLastName(separarNombre[6])
                                          .WithSecondLastName(separarNombre[7])
                                          .WithDate(Convert.ToInt32(year), Convert.ToInt32(mes), Convert.ToInt32(dia))
                                          .Build().ToString();
                }
                else if (separarNombre.Length == 9)
                {
                    rfc = RfcFacil.RfcBuilder.ForNaturalPerson()
                                          .WithName(separarNombre[0] + " " + separarNombre[1] + " " + separarNombre[2] + " " + separarNombre[3] + " " + separarNombre[4] + " " + separarNombre[5] + " " + separarNombre[6])
                                          .WithFirstLastName(separarNombre[7])
                                          .WithSecondLastName(separarNombre[8])
                                          .WithDate(Convert.ToInt32(year), Convert.ToInt32(mes), Convert.ToInt32(dia))
                                          .Build().ToString();
                }
                else if (separarNombre.Length == 10)
                {
                    rfc = RfcFacil.RfcBuilder.ForNaturalPerson()
                                          .WithName(separarNombre[0] + " " + separarNombre[1] + " " + separarNombre[2] + " " + separarNombre[3] + " " + separarNombre[4] + " " + separarNombre[5] + " " + separarNombre[6] + " " + separarNombre[7])
                                          .WithFirstLastName(separarNombre[8])
                                          .WithSecondLastName(separarNombre[9])
                                          .WithDate(Convert.ToInt32(year), Convert.ToInt32(mes), Convert.ToInt32(dia))
                                          .Build().ToString();
                }
                else
                {
                    rfc = "Error RFC";
                }


                return rfc.ToString();
            }
            catch (Exception)
            {

                throw;
            }
        }

        #endregion      
    }
}
 