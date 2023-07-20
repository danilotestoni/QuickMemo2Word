using System;
using System.IO;
using System.IO.Compression;
using Newtonsoft.Json;
using Xceed.Words.NET;

namespace QuickMemo2Word {
    static class Program {
         static void Main()
        {
            // Obtener la ruta de la carpeta donde se ejecuta el programa
            string rutaCarpeta = Path.Combine(Directory.GetCurrentDirectory(), "WorkingDirectory");

            // Recorrer la carpeta y procesar los archivos .lqm encontrados
            try
            {
                // Obtener la lista de archivos en la carpeta actual
                string[] archivosLQM = Directory.GetFiles(rutaCarpeta, "*.lqm");

                if (archivosLQM.Length == 0)
                {
                    Console.WriteLine("No se encontraron archivos .lqm en la carpeta.");
                    return;
                }

                Console.WriteLine("Archivos .lqm encontrados: {0}", archivosLQM.Length);

                foreach (string archivolqm in archivosLQM)
                {
                    LqmFunctions lqmFunctions = new();
                    Console.WriteLine("Procesando archivo .lqm: {0}", archivolqm);
                    lqmFunctions.ProcesarArchivoLqm(archivolqm);
                }
                Console.WriteLine("Procesmiento ficheros .lqm finalizado. Pulse una tecla para terminar...");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener los archivos .lqm: {ex.Message}");
            }
        }
        
    }
}
