using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace QuickMemo2Word {
    public class LqmFunctions {

        public void ProcesarArchivoLqm(string archivoLqm) {
            // Comprobar si el archivo .lqm existe
            if (!File.Exists(archivoLqm)) {
                Console.WriteLine($"El archivo {archivoLqm} no existe.");
                return;
            }

            try {
                // Crear una copia del archivo .lqm con extensión .zip
                string archivoZip = Path.ChangeExtension(archivoLqm, ".zip");
                File.Copy(archivoLqm, archivoZip, true);

                // Obtener el nombre del archivo sin extensión
                string nombreArchivo = Path.GetFileNameWithoutExtension(archivoLqm);

                // Crear una carpeta con el mismo nombre del archivo
                string rutaCarpetaDestino = Path.Combine(Directory.GetCurrentDirectory(), nombreArchivo);
                string rutaNuevaCarpeta = CrearCarpeta(rutaCarpetaDestino);
                Console.WriteLine($"Ruta de la nueva carpeta: {rutaNuevaCarpeta}");

                // Descomprimir el archivo .zip dentro de la carpeta creada
                DescomprimirArchivoZip(archivoZip, rutaNuevaCarpeta);

                // Eliminar el archivo .zip
                File.Delete(archivoZip);

                //Procesar los archivos dentro de la rutaNuevaCarpeta
                MemoInfoModel memoInfo = LeerArchivoMemoInfo(Path.Combine(rutaNuevaCarpeta, "memoinfo.jlqm"));

                string rutaCarpetaWord = CrearCarpeta(memoInfo.Category.CategoryName);

                // Ruta y nombre del nuevo archivo Word que se va a crear
                string rutaArchivoWord = Path.Combine(rutaCarpetaWord, nombreArchivo + ".docx");

                // Crear el archivo Word y agregar contenido
                try {
                    // Crear un nuevo documento Word
                    using (DocX documento = DocX.Create(rutaArchivoWord)) {
                        // Agregar el título centrado
                        Xceed.Document.NET.Formatting formatoTitulo = new Xceed.Document.NET.Formatting {
                            Size = 18D,
                            Bold = true,
                        };

                        documento.InsertParagraph(memoInfo.Category.CategoryName, false, formatoTitulo);
                        documento.InsertParagraph().SpacingAfter(20D);

                        // Agregar párrafos con diferentes estilos
                        Xceed.Document.NET.Formatting formatoParrafoNormal = new Xceed.Document.NET.Formatting {
                            Size = 12D,
                            Bold = false,
                            Italic = false
                        };

                        // Agregar imágenes entre párrafos
                        string rutaImagenMemo = Path.Combine(rutaNuevaCarpeta, "images/" + memoInfo.Memo.PreviewImage);

                        if (File.Exists(rutaImagenMemo)) {
                            Xceed.Document.NET.Picture imagenMemo = documento.AddImage(rutaImagenMemo).CreatePicture();
                            var parrafoImagenMemo = documento.InsertParagraph();
                            parrafoImagenMemo.AppendPicture(imagenMemo).Alignment = Xceed.Document.NET.Alignment.center;
                        }

                        documento.InsertParagraph().SpacingAfter(40D);

                        foreach (var memoObject in memoInfo.MemoObjectList) {
                            if (memoObject.DescRaw != null) {
                                documento.InsertParagraph(memoObject.DescRaw.ToString(), false, formatoParrafoNormal);
                                documento.InsertParagraph().SpacingAfter(20D);
                            }

                            if (memoObject.FileName != null) {
                                string rutaImagenObject = Path.Combine(rutaNuevaCarpeta, "images/" + memoObject.FileName);

                                if (File.Exists(rutaImagenObject)) {
                                    Xceed.Document.NET.Picture imagenObject = documento.AddImage(rutaImagenObject).CreatePicture();
                                    var parrafoImagenObject = documento.InsertParagraph();
                                    parrafoImagenObject.AppendPicture(imagenObject).Alignment = Xceed.Document.NET.Alignment.center;
                                    documento.InsertParagraph().SpacingAfter(20D);
                                }
                            }
                        }

                        // Guardar el documento
                        documento.Save();

                        //Eliminar carpeta que ya no es necesaria
                        EliminarCarpeta(rutaNuevaCarpeta);
                    }
                } catch (Exception ex) {
                    Console.WriteLine($"Error al crear el archivo Word: {ex.Message}");
                }
                Console.WriteLine("Procesamiento exitoso. Fichero en: {0}", rutaArchivoWord);
            } catch (Exception ex) {
                Console.WriteLine($"Error al procesar el archivo: {ex.Message}");
            }
        }

        static MemoInfoModel LeerArchivoMemoInfo(string rutaArchivo) {
            MemoInfoModel objetoMemo = new();
            try {
                string contenidoArchivo = File.ReadAllText(rutaArchivo);
                if (!string.IsNullOrEmpty(contenidoArchivo)) {
                    objetoMemo = JsonConvert.DeserializeObject<MemoInfoModel>(contenidoArchivo);
                } else {
                    Console.WriteLine($"Error al deserializar el JSON: el contenido del archivo es nulo o vacío");
                }
                return objetoMemo;
            } catch (Exception ex) {
                Console.WriteLine($"Error al leer el archivo o deserializar el JSON: {ex.Message}");
                return objetoMemo;
            }
        }

        static void DescomprimirArchivoZip(string archivoZip, string directorioDestino) {
            try {
                ZipFile.ExtractToDirectory(archivoZip, directorioDestino);
                Console.WriteLine("Descompresión exitosa.");
            } catch (Exception ex) {
                Console.WriteLine($"Error al descomprimir el archivo: {ex.Message}");
            }
        }

        static string CrearCarpeta(string nombreCarpeta) {
            string rutaNuevaCarpeta = Path.Combine(Directory.GetCurrentDirectory(), nombreCarpeta);

            try {
                Directory.CreateDirectory(rutaNuevaCarpeta);
                Console.WriteLine($"Carpeta \"{nombreCarpeta}\" creada exitosamente.");
                return rutaNuevaCarpeta;
            } catch (Exception ex) {
                Console.WriteLine($"Error al crear la carpeta: {ex.Message}");
                return string.Empty;
            }
        }

        static void EliminarCarpeta(string rutaCarpeta) {
            try {
                if (Directory.Exists(rutaCarpeta)) {
                    Directory.Delete(rutaCarpeta, true);
                } else {
                    Console.WriteLine("La carpeta no existe en la ruta especificada.");
                }
            } catch (Exception ex) {
                Console.WriteLine($"Error al eliminar la carpeta: {ex.Message}");
            }
        }
    }
}
