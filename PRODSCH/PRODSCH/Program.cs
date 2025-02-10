using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.Diagnostics;  // For Process.Start (opens up exel sheet)

public class Program
{
    public static void Main()
    {
        // Go to folder in Downloads where I saved PRODSCH files
        string folderPath = @"C:\Users\lleuterio3\Downloads\PRODSCH files";

        // Get all the files in that dir
        string[] files = Directory.GetFiles(folderPath);

        // Select the most recent file, based on write time
        string latestFile = files.MaxBy(File.GetLastWriteTime) ?? throw new InvalidOperationException("No files found in the directory.");

        // Create Excel workbook and its sheets
        var workbook = new XLWorkbook();
        var main = CreateWorksheet(workbook, "Sheet1", 36.45);
        var FA = CreateWorksheet(workbook, "FA", 30.86);
        var ST = CreateWorksheet(workbook, "ST", 34.57);
        var AR = CreateWorksheet(workbook, "AR", 28.57);
        var HR = CreateWorksheet(workbook, "HR", 31.30);
        var OIT = CreateWorksheet(workbook, "OIT", 36.57);

        // Start adding in data at row 
        int row = 2, rowFA = 2, rowST = 2, rowAR = 2, rowHR = 2, rowOIT = 2;

        // List of prefixes to check
        string[] validPrefixes = { "FA", "ST", "AR", "HR", "OIT" };

        // Create a dictionary to map the prefixes to their sheets and rows
        var prefixToSheet = new Dictionary<string, (IXLWorksheet sheet, int row, string lastDayOfWeek)>
        {
            { "FA", (FA, rowFA, string.Empty) },
            { "ST", (ST, rowST, string.Empty) },
            { "AR", (AR, rowAR, string.Empty) },
            { "HR", (HR, rowHR, string.Empty) },
            { "OIT", (OIT, rowOIT, string.Empty) }
        };

        // Capture these patterns:
        // QUEUED       {Chain:OIT_C_ADASTRA_XFER} Thu Jan 30 2025 09:45 
        // {Module:OIT_M_GZPEMAL} Thu Jan 30 2025 10:39 EST5EDT (GMT-5.0) (Dls)
        string pattern = @"\s*(QUEUED\s+)?\{(Chain|Module):(\S+)\}\s+(\w+\s+\w+\s+\d+\s+\d{4}\s+\d{1,2}:\d{2})";

        // Read file line by line. Using StreamReader for larger files
        using (StreamReader reader = new StreamReader(latestFile))
        {
            string line;

            // To keep track of day of week for main worksheet
            string dayOfWeek = string.Empty;

            // Read the file line by line to see if matches the pattern
            while ((line = reader.ReadLine()) != null)
            {
                // Check for a match with the pattern
                Match match = Regex.Match(line, pattern);

                if (match.Success)
                {
                    // Set variables accordingly
                    string type = match.Groups[2].Value; // "Chain" or "Module"
                    string name = match.Groups[3].Value; // Chain/Module name
                    string dateTime = match.Groups[4].Value; // Date and time

                    // Extract day of the week from dateTime (first 3 characters of the date part)
                    string currentDayOfWeek = dateTime.Substring(0, 3); // Ex: "Thu", "Mon"

                    // To add header if new day of week
                    if (dayOfWeek != currentDayOfWeek)
                    {
                        // Add header to main sheet
                        AddHeader(main, row, dateTime);
                            
                        // Update row since now we have added two extra rows
                        row = row + 2;
                            
                        // Update dayOfWeek to be the new day to look for
                        dayOfWeek = currentDayOfWeek;
                    }

                    //Console.WriteLine(currentDayOfWeek);

                    // If it's a "QUEUED" chain, process it
                    // If it's a module, process it. Checks to make sure the reading line does not say "INACTIVE" in front of it
                    if ((line.Contains("QUEUED") && type == "Chain") || (line.Contains("{Module:") && !line.Contains("INACTIVE")))
                    {
                        //Console.WriteLine($"Chain: {name}");
                        //Console.WriteLine($"Date/Time: {dateTime}");
 
                        // Add data to Excel
                        main.Cell(row, 1).Value = name;
                        main.Cell(row, 2).Value = dateTime;
                        row++; // Move to next row

                        // Add each application to appropriate sheet
                        // Loop through the dictionary and check if the name starts with the prefix
                        foreach (var prefix in prefixToSheet.Keys)          // iterates over all the keys (FA, ST, etc) in the prefixToSheet dictionary
                        {
                            if (name.StartsWith(prefix))                    // Checks to see if the name contains one of the prefixes
                            {
                                // Tuple deconstruction
                                var (sheet, row1, day) = prefixToSheet[prefix];  

                                // Writes the name into the first column in the corresponding sheet
                                sheet.Cell(row1, 1).Value = name;

                                // Writes the dateTime into the second column in the correcponding sheet
                                sheet.Cell(row1, 2).Value = dateTime;

                                // See if it is a new day of the week. if strings match, skip this step
                                if (day != currentDayOfWeek)
                                {
                                    // Add header to the relating sheet
                                    AddHeader(sheet, row1, dateTime);

                                    // Update row1 since now we have added two extra rows
                                    row1 = row1 + 2;

                                    // Update dayOfWeek to be the new day to look for
                                    day = currentDayOfWeek;
                                }

                                // Increment the row for the current sheet
                                prefixToSheet[prefix] = (sheet, row1 + 1, day);
                                break;
                            }
                        }
                    }
                }
            }

            // Save the workbook to a file
            String outputFilePath = @"C:\Users\lleuterio3\Documents\PRODSCH_C#.xlsx";
            workbook.SaveAs(outputFilePath);

            Console.WriteLine("\nExcel file created successfully!");

            // Open the newly created Excel file
            try
            {
                // Start the process to open the file
                // Needed to launch file with appropriate application
                Process.Start(new ProcessStartInfo(outputFilePath) { UseShellExecute = true });

                Console.WriteLine("\nExcel file created and opened successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while opening the file: {ex.Message}");
            }
        }
    }

    // Creates excel sheet w/ formatting
    public static IXLWorksheet CreateWorksheet(XLWorkbook workbook, string name, double width)
    {
        // Create worksheet
        var worksheet = workbook.AddWorksheet(name);

        // Add headers to the first row
        worksheet.Cell(1, 1).Value = "Job/Process Flow";
        worksheet.Cell(1, 2).Value = "Start Time";
        worksheet.Row(1).Style.Font.Bold = true;                                // Make the entire row bold
        worksheet.Row(1).Style.Font.Underline = XLFontUnderlineValues.Single;   // Underline the entire row

        // Freeze the first row
        worksheet.SheetView.FreezeRows(1);

        // Set column to width
        worksheet.Column(1).Width = width;  // job name
        worksheet.Column(2).Width = 20;     // time

        return worksheet;
    }

    public static void AddHeader(IXLWorksheet worksheet, int row, string date)
    {
        // Split the dateTime variable to only get the date part
        string[] dateParts = date.Split(' ');                       // Split the string into parts based on spaces
        string dateOnly = string.Join(" ", dateParts.Take(4));      // get only the first 4 parts of the date ex: "Thu Jan 30 2025"

        // Add a row above twice
        worksheet.Row(row).InsertRowsAbove(2);

        // Add the dateTime above in col 1
        // row + 1 is for adjustment since row is still set to where data was originally placed
        worksheet.Cell(row + 1, 1).Value = dateOnly;

        // Bold and underline it
        worksheet.Row(row + 1).Style.Font.Bold = true;
        worksheet.Row(row + 1).Style.Font.Underline = XLFontUnderlineValues.Single;
    }

}


