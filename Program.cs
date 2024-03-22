using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text.RegularExpressions;
using OfficeOpenXml;


class Program
{
    static List<string> excludePaths = new List<string>();
    static bool verbose = false;
    static bool suppress = false;
    static bool opsecMode = false;

    static void Main(string[] args)
    {

        PrintBanner();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set the EPPlus license context to NonCommercial to suppress trial warnings - Disclaimer: I am not responsible for improper commercial use :)
        var directoriesToProcess = new List<string>();
        var searchPatterns = new List<string>();
        var fileNames = new List<string>();

        foreach (var arg in args)
        {
            string key = arg.Contains("=") ? arg.Substring(0, arg.IndexOf('=')) : arg;
            string value = arg.Contains("=") ? arg.Substring(arg.IndexOf('=') + 1).Trim('"') : "";

            switch (key.ToLower())
            {
                case "--directory":
                    directoriesToProcess.AddRange(ParseInputToList(value));
                    break;
                case "--exclude":
                    excludePaths.AddRange(ParseInputToList(value));
                    break;
                case "--type":
                    searchPatterns.AddRange(ParseInputToList(value));
                    break;
                case "--name":
                    fileNames.AddRange(ParseInputToList(value));
                    break;
                case "--help":
                    DisplayHelpMenu();
                    return;
                case "--verbose":
                    verbose = true;
                    break;
                case "--suppress":
                    suppress = true;
                    break;
                case "--opsec":
                    opsecMode = true;
                    break;
            }
        }

        // Debug output - prints options selected by the user
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"directory: {string.Join(", ", directoriesToProcess)}");
        Console.WriteLine($"exclude: {string.Join(", ", excludePaths)}");
        Console.WriteLine($"file type: {string.Join(", ", searchPatterns)}");
        Console.WriteLine($"name: {string.Join(", ", fileNames)}");
        Console.WriteLine($"verbose: {verbose}");
        Console.WriteLine($"suppress: {suppress}");
        Console.WriteLine($"opsec: {opsecMode}");
        Console.ResetColor();

        if (!directoriesToProcess.Any())
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("No directories specified for processing.");
            Console.ResetColor();
            DisplayHelpMenu();
            return;
        }

        foreach (var directory in directoriesToProcess)
        {
            if (!Directory.Exists(directory))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Directory does not exist: {directory}");
                Console.ResetColor();
                continue;
            }

            var matchedFiles = TraverseDirectory(directory, searchPatterns, fileNames);
            OutputToConsoleAndExcel(matchedFiles, directory);
        }

        Console.ResetColor(); // Ensures the terminal colour is reset to default after execution has concluded
    }

    static void PrintBanner()
    {
        string banner = "SharpMole";
        Console.ForegroundColor = ConsoleColor.Blue;
        Console.WriteLine(new String('=', banner.Length));
        Console.WriteLine(banner);
        Console.WriteLine(new String('=', banner.Length));
        Console.ResetColor();
    }

    static List<string> ParseInputToList(string input)
    {
        if (File.Exists(input))
        {
            try
            {
                return File.ReadAllLines(input).ToList();
            }
            catch (IOException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error reading file: {ex.Message}");
                Console.ResetColor();
                Environment.Exit(1);
            }
        }
        else if (input.Contains(","))
        {
            // Split by commas and trim spaces from each path
            return input.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(path => path.Trim()).ToList();
        }
        return new List<string> { input.Trim() }; // Ensure single input path is also trimmed
    }

    static List<string> TraverseDirectory(string rootDirectory, List<string> searchPatterns, List<string> fileNames)
    {
        List<string> filesFound = new List<string>();
        var options = new EnumerationOptions { IgnoreInaccessible = true, RecurseSubdirectories = true };
        var directoriesToExclude = excludePaths.Select(Path.GetFullPath).ToList(); // Normalize the paths

        foreach (var file in Directory.EnumerateFiles(rootDirectory, "*", options))
        {
            var fullFilePath = Path.GetFullPath(file); // Normalize the file path
            var fileInfo = new FileInfo(fullFilePath); // Declare fileInfo here with the normalized path
            if (directoriesToExclude.Any(exclude => fullFilePath.StartsWith(exclude, StringComparison.OrdinalIgnoreCase)))
            {
                if (verbose)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"Excluding file: {fullFilePath}");
                    Console.ResetColor();
                }
                continue;
            }


            if (searchPatterns.Any() && !searchPatterns.Any(pattern => fileInfo.Extension.Equals($".{pattern}", StringComparison.OrdinalIgnoreCase))) continue;
            if (fileNames.Any() && !fileNames.Any(name => fileInfo.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)) continue;

            if (opsecMode && !CanAccessFile(fileInfo.FullName)) continue; // Skip files that cannot be accessed under opsecMode

            if (verbose)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Enumerating file: {fileInfo.FullName}");
                Console.ResetColor();
            }

            filesFound.Add(fileInfo.FullName);
        }

        return filesFound;
    }



    static void OutputToConsoleAndExcel(List<string> files, string directory)
    {
        if (!suppress)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Processing {files.Count} files in {directory}");
            Console.ResetColor();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Permissions");
                worksheet.Cells[1, 1].Value = "File Path";
                worksheet.Cells[1, 2].Value = "Read";
                worksheet.Cells[1, 3].Value = "Write";
                worksheet.Cells[1, 4].Value = "Read & Write";

                int row = 2;
                foreach (var filePath in files)
                {
                    var fileInfo = new FileInfo(filePath); // Declare fileInfo here for each file path
                    bool canRead = false, canWrite = false;

                    if (opsecMode && !CanAccessFile(filePath)) // Integrated opsecMode check
                    {
                        continue; // Skip files that cannot be accessed under opsecMode
                    }

                    try
                    {
                        var accessControl = fileInfo.GetAccessControl();
                        var rules = accessControl.GetAccessRules(true, true, typeof(SecurityIdentifier));

                        foreach (FileSystemAccessRule rule in rules)
                        {
                            if ((rule.FileSystemRights & FileSystemRights.ReadData) > 0 && rule.AccessControlType == AccessControlType.Allow)
                                canRead = true;
                            if ((rule.FileSystemRights & FileSystemRights.WriteData) > 0 && rule.AccessControlType == AccessControlType.Allow)
                                canWrite = true;
                        }
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Access denied when processing security information for file: {filePath}");
                        Console.WriteLine($"Error: {ex.Message}");
                        Console.ResetColor();
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"An exception occurred while processing file: {filePath}");
                        Console.WriteLine($"Error: {ex.ToString()}");
                        Console.ResetColor();
                    }
                    finally
                    {
                        if (canRead || canWrite) // Condition to include file in the spreadsheet
                        {
                            worksheet.Cells[row, 1].Value = filePath;
                            worksheet.Cells[row, 2].Value = canRead ? "Y" : "N";
                            worksheet.Cells[row, 3].Value = canWrite ? "Y" : "N";
                            worksheet.Cells[row, 4].Value = (canRead && canWrite) ? "Y" : "N";
                            row++;
                        }
                    }
                }

                if (!suppress)
                {
                    var cleanDirectory = MakeValidFileName(directory);
                    var excelFileName = MakeValidFileName(directory);
                    package.SaveAs(new FileInfo(excelFileName));
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Excel file saved: {excelFileName}");
                    Console.ResetColor();
                }
            }
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"Suppressed writing output for {files.Count} files in {directory}");
            Console.ResetColor();
        }
    }


    static bool CanAccessFile(string filePath)
    {
        try
        {
            // Use the GetAccessControl method from FileInfo
            var fileInfo = new FileInfo(filePath);
            var accessControlList = fileInfo.GetAccessControl();
            var accessRules = accessControlList.GetAccessRules(true, true, typeof(SecurityIdentifier));

            var windowsIdentity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(windowsIdentity);

            foreach (FileSystemAccessRule rule in accessRules)
            {
                if (principal.IsInRole((SecurityIdentifier)rule.IdentityReference))
                {
                    if ((rule.FileSystemRights & FileSystemRights.ReadData) != 0)
                    {
                        if (rule.AccessControlType == AccessControlType.Allow)
                            return true;
                        if (rule.AccessControlType == AccessControlType.Deny)
                            return false;
                    }
                }
            }
            // If no deny rules are found for the current user, assume access is allowed.
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            // Exception thrown when the user does not have access to the target file.
            return false;
        }
        catch
        {
            // Handle any other exception (for example, file not found), assume no access granted.
            return false;
        }
    }

    static string MakeValidFileName(string name)
    {
        // Normalize network paths
        if (name.StartsWith("\\\\"))
        {
            name = name.TrimStart('\\').Replace('\\', '_').Replace('$', ''); // Removes $ symbol and replaces \ with _
        }
        else if (Path.IsPathRooted(name) && name.EndsWith(":\\"))
        {
            // If the path is a root drive path (e.g., "C:\"), then just return the drive letter followed by "_Drive_permissions.xlsx"
            // For root directories, remove the colon and trailing backslash and append "_Drive"
            name = name.TrimEnd('\\').TrimEnd(':') + "_Drive_permissions";
        }

        // Replace invalid filename characters with an underscore
        string invalidChars = new string(Path.GetInvalidFileNameChars());
        foreach (char invalidChar in invalidChars)
        {
            name = name.Replace(invalidChar, '_');
        }

        name = name + ".xlsx"; //Ensures the returned filename always includes a .XLSX extension
        return name;
    }

    static void DisplayHelpMenu()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  --directory=<path | comma seperated list | file>    Specify the root directory(s) for the search or a file containing a list of directories.");
        Console.WriteLine("  --exclude=<path | comma seperated list | file>      Specify paths to exclude or a file containing a list of paths to exclude.");
        Console.WriteLine("  --type=<extension | comma seperated list | file>    Specify file extensions to search for or a file containing a list of extensions.");
        Console.WriteLine("  --name=<filename | comma seperated list | file>     Specify file names to search for or a file containing a list of file names.");
        Console.WriteLine("  --help                     Display this help message.");
        Console.WriteLine("  --verbose                  Print detailed enumeration to the console.");
        Console.WriteLine("  --suppress                 Suppress writing to output files.");
        Console.WriteLine("\nExamples:");
        Console.WriteLine("  --directory=C:\\Users\\ --type=txt");
        Console.WriteLine("  --directory=directories.txt --exclude=excludes.txt --type=extensions.txt --name=names.txt --verbose --suppress");
        Console.WriteLine("  --directory=C:\\ --exclude=excludes.txt --type=extensions.txt --name=\"example1,example2\" --verbose --suppress");
    }
}
