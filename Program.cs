using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Diagnostics;

class Program
{
    static HashSet<string> scannedQRs = new HashSet<string>(); // Conjunto para almacenar los QR comunes validados
    static int qrCount = 1; // Contador de cantidad de QR escaneados, empieza en 1
    static string scannedQRFilePath = "scanned_qrs.txt"; // Archivo donde se guardarán los QR escaneados
    static string excelFilePath = "QRData.xlsx"; // Archivo donde se guardarán los datos en Excel de los QR comunes
    static string adminQR = "SERGIO_ORTIZ_001"; // QR de Administrador (debe ser un QR específico)
    static HashSet<string> userQRs = new HashSet<string>(); // QR de usuarios que pueden acceder
    static HashSet<string> adminQRs = new HashSet<string>(); // QR de administradores
    static bool programUnlocked = false; // Flag para verificar si el programa está desbloqueado

    static void Main()
    {
        // Cargar los QR escaneados previamente desde los archivos de administradores y usuarios
        LoadQRsFromFiles();

        Console.WriteLine("Escanea un QR para iniciar:");

        string scannedQR = ScanQRCode(); // Escanear el primer QR

        // Verificar si es un QR Administrador o Usuario
        if (adminQRs.Contains(scannedQR))
        {
            // Si es QR Administrador, desbloqueamos el acceso
            Console.WriteLine("Acceso Administrador desbloqueado.");
            AdminMenu(); // Mostrar menú Administrador
            programUnlocked = true;
        }
        else if (userQRs.Contains(scannedQR))
        {
            // Si es QR Usuario, desbloqueamos el acceso
            Console.WriteLine("Acceso Usuario desbloqueado.");
            UserMenu(); // Mostrar menú Usuario
            programUnlocked = true;
        }
        else
        {
            // Si el QR no es Administrador ni Usuario, muestra un error
            Console.WriteLine("Error: No se puede acceder sin un QR válido de Administrador o Usuario.");
        }

        // Si el programa está desbloqueado, permitir escanear QR comunes
        if (programUnlocked)
        {
            ScanAndValidateQRCommon();
        }
    }

    static void LoadQRsFromFiles()
    {
        // Cargar administradores desde archivo
        if (File.Exists("admin_qrs.txt"))
        {
            var adminLines = File.ReadAllLines("admin_qrs.txt");
            foreach (var line in adminLines)
            {
                adminQRs.Add(line);
            }
        }

        // Cargar usuarios desde archivo
        if (File.Exists("user_qrs.txt"))
        {
            var userLines = File.ReadAllLines("user_qrs.txt");
            foreach (var line in userLines)
            {
                userQRs.Add(line);
            }
        }
        else
        {
            // Si no existe el archivo 'user_qrs.txt', crearlo
            Console.WriteLine("No se encontró el archivo de usuarios. Creando archivo...");
            File.WriteAllText("user_qrs.txt", "");  // Crear un archivo vacío si no existe
        }
    }

    // Función para escanear y validar un QR común
    static void ScanAndValidateQRCommon()
    {
        Console.WriteLine("Escanea un QR Común para validar:");

        string scannedQR = ScanQRCode(); // Escanear el QR

        if (IsValidQR(scannedQR))
        {
            // Verificar si el QR es repetido
            if (scannedQRs.Contains(scannedQR))
            {
                Console.WriteLine("QR repetido. No se puede agregar.");
            }
            else
            {
                // Si es válido y no repetido, guardarlo en el archivo Excel
                SaveQRToExcel(scannedQR, DateTime.Now, "Cliente");
                scannedQRs.Add(scannedQR); // Agregar a la lista de QR escaneados
                Console.WriteLine("QR válido y almacenado exitosamente.");
            }
        }
        else
        {
            // Si el QR es inválido, mostrar un mensaje
            Console.WriteLine("QR inválido. No se puede agregar.");
        }
    }

    // Función para validar un QR común (verificar longitud y caracteres)
    static bool IsValidQR(string qrData)
    {
        if (qrData.Length != 102) return false;

        int countHash = 0, countAsterisk = 0, countEqual = 0;
        foreach (char c in qrData)
        {
            if (c == '#') countHash++;
            else if (c == '*') countAsterisk++;
            else if (c == '=') countEqual++;
        }

        return (countHash == 8 && countAsterisk == 4 && countEqual == 1);
    }

    // Función para escanear el QR desde el teclado
    static string ScanQRCode()
    {
        try
        {
            string scannedData = string.Empty;

            ConsoleKeyInfo keyInfo;
            while (true)
            {
                keyInfo = Console.ReadKey(intercept: true); // Lee las teclas sin mostrarlas en la consola

                if (keyInfo.Key == ConsoleKey.Enter)
                {
                    break;
                }

                scannedData += keyInfo.KeyChar; // Agregar el carácter escaneado
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

    // Menú para el QR Usuario
    static void UserMenu()
    {
        while (true)
        {
            Console.WriteLine("\nMenú Usuario");
            Console.WriteLine("1. Escanear QR Común");
            Console.WriteLine("2. Ver QR Registrados (Escribe 'VER')");
            Console.WriteLine("3. Salir");
            string choice = Console.ReadLine();

            // Validar la opción seleccionada
            switch (choice)
            {
                case "1":
                    ScanAndValidateQRCommon();  // Escanear y validar un QR común
                    break;

                case "2":
                    ViewQRClients();  // Ver el archivo Excel con los QR registrados
                    break;

                case "3":
                    Console.WriteLine("Saliendo del menú Usuario.");
                    return;  // Salir del menú

                default:
                    Console.WriteLine("Opción no válida. Intenta nuevamente.");
                    break;
            }
        }
    }

    // Menú para el QR Administrador
    static void AdminMenu()
    {
        while (true)
        {
            Console.WriteLine("\nMenú Administrador");
            Console.WriteLine("1. Agregar QR Usuario");
            Console.WriteLine("2. Modificar QR Usuario (Cambiar permisos de Administrador)");
            Console.WriteLine("3. Eliminar QR Usuario");
            Console.WriteLine("4. Ver QR Clientes (Escribe 'VER')");
            Console.WriteLine("5. Asignar Permiso Administrativo a Usuario");
            Console.WriteLine("6. Salir");
            string choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    AddQRUser();
                    break;
                case "2":
                    ModifyQRUser();
                    break;
                case "3":
                    DeleteQRUser();
                    break;
                case "4":
                    ViewQRClients();  // Ver el archivo Excel con los QR registrados
                    break;
                case "5":
                    AssignAdminPermissions();
                    break;
                case "6":
                    Console.WriteLine("Saliendo del menú Administrador.");
                    return;  // Salir del menú
                default:
                    Console.WriteLine("Opción no válida. Intenta nuevamente.");
                    break;
            }
        }
    }

    // Función para agregar un QR de usuario
    static void AddQRUser()
    {
        Console.WriteLine("Escanea el QR Usuario (debe ser un QR de 102 caracteres):");
        string qrData = ScanQRCode();

        if (IsValidQR(qrData))
        {
            userQRs.Add(qrData);
            SaveScannedQR(qrData, "user_qrs.txt"); // Guardar el QR en el archivo de texto
            SaveQRToExcel(qrData, DateTime.Now, "Usuario");
            Console.WriteLine("QR Usuario agregado exitosamente.");
        }
        else
        {
            Console.WriteLine("QR inválido. Asegúrate de que tenga exactamente 102 caracteres.");
        }
    }

    // Función para modificar un QR de usuario
    static void ModifyQRUser()
    {
        Console.WriteLine("Modificar QR Usuario: ");
        Console.WriteLine("Lista de Usuarios:");
        foreach (var qr in userQRs)
        {
            Console.WriteLine(qr);
        }
        Console.WriteLine("Escribe el QR que deseas modificar:");

        string qrData = Console.ReadLine();
        if (userQRs.Contains(qrData))
        {
            // Aquí puedes agregar la lógica para cambiar permisos (de Usuario a Administrador o viceversa)
            if (adminQRs.Contains(qrData))
            {
                adminQRs.Remove(qrData);
                Console.WriteLine("QR removido de administradores.");
            }
            else
            {
                adminQRs.Add(qrData);
                Console.WriteLine("QR agregado como administrador.");
            }

            // Guardar cambios
            SaveQRToFile("admin_qrs.txt", adminQRs);
            SaveQRToFile("user_qrs.txt", userQRs);
        }
        else
        {
            Console.WriteLine("QR no encontrado.");
        }
    }

    // Función para eliminar un QR de usuario
    static void DeleteQRUser()
    {
        Console.WriteLine("Eliminar QR Usuario: ");
        Console.WriteLine("Lista de Usuarios:");
        foreach (var qr in userQRs)
        {
            Console.WriteLine(qr);
        }
        Console.WriteLine("Escribe el QR que deseas eliminar:");

        string qrData = Console.ReadLine();
        if (userQRs.Contains(qrData))
        {
            userQRs.Remove(qrData);
            SaveQRToFile("user_qrs.txt", userQRs);
            Console.WriteLine("QR de usuario eliminado.");
        }
        else
        {
            Console.WriteLine("QR no encontrado.");
        }
    }

    // Función para asignar permisos administrativos a un QR de usuario
    static void AssignAdminPermissions()
    {
        Console.WriteLine("Asignar permisos administrativos a un QR Usuario: ");
        // Lógica para asignar permisos administrativos.
    }

    // Función para guardar el QR en un archivo
    static void SaveQRToFile(string filePath, HashSet<string> qrSet)
    {
        try
        {
            File.WriteAllLines(filePath, qrSet);
        }
        catch (IOException ex)
        {
            Console.WriteLine("Error al guardar en el archivo: " + ex.Message);
        }
    }

    // Función para guardar el QR en el archivo Excel
    static void SaveQRToExcel(string qrData, DateTime timestamp, string tipoQR)
    {
        var fileInfo = new FileInfo(excelFilePath);
        using (var package = new ExcelPackage(fileInfo))
        {
            var worksheet = package.Workbook.Worksheets["DatosQR"];
            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add("DatosQR");
                worksheet.Cells[1, 1].Value = "ID QR";
                worksheet.Cells[1, 2].Value = "Fecha";
                worksheet.Cells[1, 3].Value = "QR";
                worksheet.Cells[1, 4].Value = "Tipo QR";
            }

            int row = worksheet.Dimension?.Rows + 1 ?? 2;
            worksheet.Cells[row, 1].Value = qrCount++; // Nuevo ID
            worksheet.Cells[row, 2].Value = timestamp.ToString("yyyy-MM-dd HH:mm:ss"); // Fecha y hora
            worksheet.Cells[row, 3].Value = qrData; // El QR escaneado
            worksheet.Cells[row, 4].Value = tipoQR; // Tipo de QR (Administrador, Usuario, Cliente)

            package.Save();
        }
    }

    // Función para ver los QR almacenados en el archivo Excel
    static void ViewQRClients()
    {
        Console.WriteLine("\nQR Clientes almacenados:");
        if (File.Exists(excelFilePath))
        {
            // Abrir el archivo Excel con el programa predeterminado
            Process.Start(new ProcessStartInfo(excelFilePath) { UseShellExecute = true });
        }
        else
        {
            Console.WriteLine("No hay archivo Excel o no contiene QR Clientes.");
        }
    }

    // Función para guardar el QR escaneado en el archivo de texto
    static void SaveScannedQR(string qrData, string filePath)
    {
        try
        {
            File.AppendAllLines(filePath, new[] { qrData });
        }
        catch (IOException ex)
        {
            Console.WriteLine("Error al guardar en el archivo de texto: " + ex.Message);
        }
    }
}
