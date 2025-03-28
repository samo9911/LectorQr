using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

class Program
{
    static HashSet<string> scannedQRs = new HashSet<string>(); // Conjunto para almacenar los QR escaneados
    static int qrCount = 0; // Contador de cantidad de QR escaneados
    static string scannedQRFilePath = "scanned_qrs.txt"; // Archivo donde se guardarán los QR escaneados
    static string excelFilePath = "QRData.xlsx"; // Archivo donde se guardarán los datos en Excel

    static void Main()
    {
        // Cargar los QR escaneados previamente desde el archivo
        LoadScannedQRs();

        // Iniciar la lectura indefinida de QR
        Console.WriteLine("Escanea el QR de datos:");

        while (true)
        {
            // Segundo escaneo: QR de datos
            string qrData = ScanQRCode();

            if (string.IsNullOrEmpty(qrData))
            {
                Console.WriteLine("QR no leído.");
                continue; // Continúa esperando por un nuevo QR si no se lee correctamente
            }

            // Verificar si el QR ya fue leído
            if (scannedQRs.Contains(qrData))
            {
                Console.WriteLine("QR repetido.");
                continue; // Si el QR ya fue escaneado, salta al siguiente ciclo
            }
            else
            {
                // Validar el QR según las reglas
                if (IsValidQR(qrData))
                {
                    // Si es válido, lo agregamos al conjunto y lo guardamos en el archivo Excel
                    scannedQRs.Add(qrData);
                    try
                    {
                        qrCount++; // Aumentar el contador de QR escaneados
                        SaveToExcel(qrData, qrCount); // Guardar el QR con la fecha y el código
                        SaveScannedQR(qrData); // Guardar el QR escaneado en el archivo de texto
                        Console.WriteLine("QR agregado exitosamente.");
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine("Error al guardar el archivo Excel: " + ex.Message);
                    }
                }
                else
                {
                    Console.WriteLine("QR inválido. Asegúrate de que tenga exactamente 102 caracteres, 8 '#' , 4 '*' y 1 '='.");
                }
            }
        }
    }

    // Función para capturar la entrada del escáner QR (se asume que el scanner actúa como teclado)
    static string ScanQRCode()
    {
        try
        {
            string scannedData = string.Empty;

            // Leer la entrada hasta que el escáner envíe un salto de línea o se complete el QR
            ConsoleKeyInfo keyInfo;
            while (true)
            {
                keyInfo = Console.ReadKey(intercept: true); // Lee las teclas sin mostrarlas en la consola

                // Si se detecta un salto de línea, significa que el QR ha terminado
                if (keyInfo.Key == ConsoleKey.Enter)
                {
                    break;
                }

                // Agregar la tecla leída al texto del QR
                scannedData += keyInfo.KeyChar;
            }

            Console.WriteLine("\nQR leído: " + scannedData);
            return scannedData;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error al leer QR: " + ex.Message);
            return null;
        }
    }

    // Función para verificar si el QR es válido: los caracteres especiales #, *, = y longitud de 102 caracteres
    static bool IsValidQR(string qrData)
    {
        // Verificar que el QR tenga exactamente 102 caracteres
        if (qrData.Length != 102)
        {
            return false; // El QR no es válido si no tiene exactamente 102 caracteres
        }

        // Contar las ocurrencias de cada carácter especial
        int countHash = 0;
        int countAsterisk = 0;
        int countEqual = 0;

        foreach (char c in qrData)
        {
            if (c == '#') countHash++;
            else if (c == '*') countAsterisk++;
            else if (c == '=') countEqual++;
        }

        // Verificar que los caracteres estén en las cantidades correctas
        if (countHash == 8 && countAsterisk == 4 && countEqual == 1)
        {
            return true; // El QR es válido
        }
        else
        {
            return false; // El QR es inválido si no cumple con las cantidades exactas
        }
    }

    // Función para almacenar los datos en un archivo Excel
    static void SaveToExcel(string qrData, int qrId)
    {
        try
        {
            var fileInfo = new FileInfo(excelFilePath);

            // Intentar abrir y escribir en el archivo Excel
            using (var package = new ExcelPackage(fileInfo))
            {
                // Verificar si la hoja "DatosQR" ya existe
                var worksheet = package.Workbook.Worksheets["DatosQR"];

                // Si la hoja no existe, crearla
                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add("DatosQR");

                    // Agregar encabezados a las columnas
                    worksheet.Cells[1, 1].Value = "Número de QR";
                    worksheet.Cells[1, 2].Value = "Fecha";
                    worksheet.Cells[1, 3].Value = "Código QR";
                }

                // Encuentra la primera fila vacía en la hoja
                int row = worksheet.Dimension?.Rows + 1 ?? 2;

                // Escribir la data en la siguiente fila vacía
                worksheet.Cells[row, 1].Value = qrId; // Número de QR (ID único)
                worksheet.Cells[row, 2].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); // Fecha y hora
                worksheet.Cells[row, 3].Value = qrData; // El código QR

                package.Save(); // Guarda los cambios en el archivo
                Console.WriteLine("QR guardado en Excel.");
            }
        }
        catch (IOException ex)
        {
            Console.WriteLine("Error al guardar el archivo Excel: " + ex.Message);
        }
        catch (UnauthorizedAccessException ex)
        {
            Console.WriteLine("Error de acceso al archivo Excel: " + ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error inesperado al guardar el archivo Excel: " + ex.Message);
        }
    }

    // Función para cargar los QR escaneados desde el archivo de texto
    static void LoadScannedQRs()
    {
        if (File.Exists(scannedQRFilePath))
        {
            var lines = File.ReadAllLines(scannedQRFilePath);
            foreach (var line in lines)
            {
                scannedQRs.Add(line);
            }
        }
    }

    // Función para guardar un QR escaneado en el archivo de texto
    static void SaveScannedQR(string qrData)
    {
        // Agregar el QR al archivo de texto
        try
        {
            File.AppendAllLines(scannedQRFilePath, new[] { qrData });
        }
        catch (IOException ex)
        {
            Console.WriteLine("Error al guardar en el archivo de texto: " + ex.Message);
        }
    }
}
