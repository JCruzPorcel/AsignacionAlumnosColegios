using JCruzPorcel.Inscipcion_Colegio.Models.AlumnoData; // Utilizo el namespace para hacer referencia a los archivos que estoy utilizando, como la clase Alumno.
using JCruzPorcel.Inscipcion_Colegio.Models.ColegioData;
using Spectre.Console; // Esta libraria agrega nuevas funcionalidades de consola, permitiendo así la personalizacion y mejorando la experiencia de usuario.
using SpreadsheetLight; // Esta libreria se utiliza para poder acceder a los archivos Excel.
using System.Text.RegularExpressions; // Esta nos permite comprar cadenas de textos, y verificar si son iguales o contienen expresiones.

namespace JCruzPorcel.Inscipcion_Colegio.src // Utilizo este namespace para tener mas organización y que se vea quien es el autor.
{

    /// <summary>
    /// Programa para asignar alumnos a colegios según sus preferencias y las vacantes disponibles.
    /// Utiliza archivos de inscripciones y de colegios para obtener los datos necesarios.
    /// Luego genera archivos de vacantes para cada colegio con los alumnos asignados.
    /// Hecho en Visual Studio 2022 en C# por Perez Porcel Juan Cruz.
    /// </summary> 

    // Al ejecutarse se crearan las carpetas y archivos necesarios o faltantes dentro de la ruta del proyecto en la carpeta data/GeneratedFiles

    internal class Program // La clase Program se declara como "internal" para restringir su acceso desde programas o clases externas al proyecto actual. Esta clase contiene la lógica principal del programa.
    {
        #region Variables
        // Nombre de las Carpetas y Archivos
        const string FOLDER = "data";
        const string INSCRIPCIONES_FILE_PATH = "Inscripciones.xlsx";
        const string ANGUIL_FILE_PATH = "Colegio_Anguil.xlsx";
        const string TOAY_FILE_PATH = "Colegio_Toay.xlsx";
        const string SANTA_ROSA_FILE_PATH = "Colegio_SantaRosa.xlsx";
        const string OUTPUT_FILE_FOLDER = "GeneratedFiles";
        #endregion

        #region Main Logic
        static void Main(string[] args)
        {
            // Obtiene la ruta actual del proyecto
            string basePath = Directory.GetParent(Directory.GetCurrentDirectory())?.Parent?.Parent?.FullName ?? string.Empty;

            if (string.IsNullOrEmpty(basePath))
            {
                return; // Si la ruta del proyecto es nula o vacía, el programa no se ejecutará para evitar conflictos
            }

            // Ruta del archivo de plantilla
            // Combina la ruta del proyecto con la carpeta y el nombre de los archivos.
            string inscripcionesFilePath = @Path.Combine(basePath, FOLDER, INSCRIPCIONES_FILE_PATH);  // El @ en "Path.Combine" es para tomar de forma literal la cadena de texto (String).
            string anguilFilePath = @Path.Combine(basePath, FOLDER, ANGUIL_FILE_PATH);        // No es necesario, pero lo hice para evitar posibles errores con las extenciones o las diagonales.                                                                                               
            string toayFilePath = @Path.Combine(basePath, FOLDER, TOAY_FILE_PATH);           // Por ejemplo: .xlsx, / o \.
            string santaRosaFilePath = @Path.Combine(basePath, FOLDER, SANTA_ROSA_FILE_PATH);
            string outputFilePath = @Path.Combine(basePath, FOLDER, OUTPUT_FILE_FOLDER);

            List<Alumno> alumnos = LeerArchivoInscripciones(inscripcionesFilePath);
            List<Colegio> colegios = LeerArchivoColegios(toayFilePath, santaRosaFilePath, anguilFilePath);

            string? option;
            Colegio? colegioSeleccionado = null; // Tipo anulable

            do
            {
                option = AnsiConsole.Prompt(
                    new SelectionPrompt<string>()
                    .Title("[underline purple]¿A qué colegio desea asignar?[/]")
                    .AddChoices("[yellow]Anguil[/]", "[yellow]Santa Rosa[/]", "[yellow]Toay[/]", "[red]Salir[/]"));

                AsignarAlumnosAColegios(alumnos, colegios, option);

                colegioSeleccionado = colegios.FirstOrDefault(colegio =>
                    Regex.IsMatch(option, @"\b" + Regex.Escape(colegio.Nombre) + @"\b", RegexOptions.IgnoreCase));


                if (colegioSeleccionado is not null)
                {
                    try
                    {
                        string outputFileName = $"Vacantes_{colegioSeleccionado.Nombre}.xlsx";
                        CrearArchivoVacantesColegio(colegioSeleccionado.AlumnosAsignados, outputFileName, outputFilePath);

                        AnsiConsole.Markup($"[yellow]Proceso finalizado. Se han generado los archivos de vacantes para el colegio[/] [underline blue]{colegioSeleccionado.Nombre}[/].");
                        AnsiConsole.Markup("\n\n[yellow]Se generaron todos los archivos de vacantes con[/] [green]éxito[/].");
                        AnsiConsole.Markup("\n\nPresiona cualquier tecla para continuar...\n\n");
                    }
                    catch (Exception ex)
                    {
                        AnsiConsole.Markup($"[underline red]Ocurrió un error al generar los archivos de vacantes:[/] {ex.Message}");
                    }
                    Console.ReadKey();
                    Console.Clear();
                }
                else
                {
                    AnsiConsole.Markup("[underline red]No se encontró el colegio seleccionado.[/]");
                }
            } while (option != "[red]Salir[/]");


            Console.Clear();
            Environment.Exit(0);

            // Esperar a que el usuario presione una tecla para salir
            /*Console.WriteLine("\n\nPresiona cualquier tecla para salir...");
            Console.ReadKey();*/

            #endregion

            #region Method's
            static List<Alumno> LeerArchivoInscripciones(string path) // Se encargaría de leer el archivo "Inscripciones.xlsx" y extraer los datos de los alumnos inscritos.
            {
                // Crear List de Alumnos
                List<Alumno> alumnos = new List<Alumno>();

                // Cargar el archivo de plantilla
                SLDocument document = new SLDocument(path);

                // Obtener las estadísticas de la hoja de cálculo
                SLWorksheetStatistics stats = document.GetWorksheetStatistics();

                // Iterar a través de las filas
                for (int rowIndex = 2; rowIndex <= stats.EndRowIndex; rowIndex++)
                {
                    byte columnNroInscripcion = 1; // Columna 1 (A)
                    byte columnNombre = 2; // Columna 2 (B)
                    byte columnEdad = 3; // Columna 3 (C)
                    byte columnGrado = 4; // Columna 4 (D)
                    byte columnPreferencia = 5; // Columna 5 (E)


                    int nroInscripcion = document.GetCellValueAsInt32(rowIndex, columnNroInscripcion);
                    string nombre = document.GetCellValueAsString(rowIndex, columnNombre);
                    int edad = document.GetCellValueAsInt32(rowIndex, columnEdad);
                    int grado = document.GetCellValueAsInt32(rowIndex, columnGrado);
                    string preferencia = document.GetCellValueAsString(rowIndex, columnPreferencia);


                    // Imprimir los valores de las celdas
                    /*Console.WriteLine($"Número de inscripción: {nroInscripcion}");
                      Console.WriteLine($"Nombre: {nombre}");
                      Console.WriteLine($"Edad:  {edad}");
                      Console.WriteLine($"Grado: {grado}");
                      Console.WriteLine($"Preferencia: {preferencia}");

                      Console.WriteLine("--------"); // Separador entre filas*/

                    Alumno alumno = new Alumno(nroInscripcion, nombre, edad, grado, preferencia); // Crear Alumno con los datos recopilados

                    alumnos.Add(alumno); // Agregar Alumno a la lista
                }

                return alumnos; // Retornar la lista
            }

            static List<Colegio> LeerArchivoColegios(string toayFilePath, string santaRosaFilePath, string anguilFilePath) // Obtiene la información sobre los grados y las vacantes disponibles en cada colegio.
            {
                List<Colegio> colegios = new List<Colegio>();

                List<string> filePaths = new List<string> { toayFilePath, santaRosaFilePath, anguilFilePath };
                List<string> nombresColegios = new List<string> { "Toay", "Santa Rosa", "Anguil" };

                for (int i = 0; i < filePaths.Count; i++)
                {
                    string filePath = filePaths[i];
                    string nombreColegio = nombresColegios[i];

                    Colegio colegio = new Colegio(nombreColegio);

                    // Cargar el archivo de plantilla
                    SLDocument document = new SLDocument(filePath);

                    // Obtener las estadísticas de la hoja de cálculo
                    SLWorksheetStatistics stats = document.GetWorksheetStatistics();

                    // Iterar a través de las filas
                    for (int rowIndex = 2; rowIndex <= stats.EndRowIndex; rowIndex++)
                    {
                        byte columnGrado = 1; // Columna 1 (A)
                        byte columnVacantes = 2; // Columna 2 (B)

                        int grado = document.GetCellValueAsInt32(rowIndex, columnGrado);
                        int vacantes = document.GetCellValueAsInt32(rowIndex, columnVacantes);

                        colegio.AgregarGrado(grado, vacantes);
                    }


                    colegios.Add(colegio);
                }

                return colegios;
            }

            // Se cambio la forma en al que se Asignan los Alumnos a los colegios. Se agrego string colegioSeleccionado, asi se ajusta al nuevo codigo y libreria agregada Spectre.Console.
            static void AsignarAlumnosAColegios(List<Alumno> alumnos, List<Colegio> colegios, string colegioSeleccionado)
            {
                foreach (Alumno alumno in alumnos)
                {
                    foreach (Colegio colegio in colegios)
                    {
                        if (Regex.IsMatch(colegioSeleccionado, @"\b" + Regex.Escape(colegio.Nombre) + @"\b", RegexOptions.IgnoreCase))
                        {
                            // Esto verifica que los nombres de los colegios sean los mismos que su preferencia comprobando si contiene en alguna parte la misma cadena de strings
                            if (Regex.IsMatch(alumno.Preferencia, @"\b" + Regex.Escape(colegio.Nombre) + @"\b", RegexOptions.IgnoreCase) && colegio.TieneVacantes(alumno.Grado))
                            {
                                colegio.AgregarAlumno(alumno);
                                break;
                            }
                        }

                        /*if (Regex.IsMatch(alumno.Preferencia, @"\b" + Regex.Escape(colegio.Nombre) + @"\b", RegexOptions.IgnoreCase) // Esto verifica que los nombres de los colegios sean los mismos que su preferencia
                            && colegio.TieneVacantes(alumno.Grado))                                                                   // comprobando si contiene en alguna parte la misma cadena de strings
                           
                        {
                            colegio.AgregarAlumno(alumno);
                            break;
                        }*/
                    }
                }
            }

            static void CrearArchivoVacantesColegio(List<Alumno> alumnos, string outputFileName, string path)
            {
                string pathFile = path;

                SLDocument worksheet = new SLDocument();

                System.Data.DataTable dataTable = new System.Data.DataTable();

                // Columns 
                dataTable.Columns.Add("Nro Inscripción", typeof(int));
                dataTable.Columns.Add("Nombre", typeof(string));
                dataTable.Columns.Add("Edad", typeof(int));
                dataTable.Columns.Add("Grado", typeof(int));

                foreach (Alumno alumno in alumnos)
                {
                    dataTable.Rows.Add(alumno.NroInscripcion, alumno.Nombre, alumno.Edad, alumno.Grado);
                }

                if (!Directory.Exists(pathFile))
                {
                    Directory.CreateDirectory(pathFile);
                }

                worksheet.ImportDataTable(1, 1, dataTable, true);

                worksheet.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Vacantes"); // Renombra la hoja de Excel a "Vacante".

                worksheet.SaveAs(@Path.Combine(pathFile, outputFileName));
            }

        }
        #endregion
    }
}
