using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.Diagnostics;  // For Process.Start

class Program
{
    static void Main()
    {
        // Go to folder in Downloads where I saved PRODSCH files
        string folderPath = @"C:\Users\lleuterio3\Downloads\PRODSCH files";

        // Get all the files in that dir
        string[] files = Directory.GetFiles(folderPath);

        // Select the most recent file, based on write time
        string latestFile = files
            .OrderByDescending(file => File.GetLastWriteTime(file))
            .FirstOrDefault();

        if (latestFile != null)
        {
            // Create Execl sheet
            var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Sheet1");
            var FA = workbook.AddWorksheet("FA");
            var ST = workbook.AddWorksheet("ST");
            var AR = workbook.AddWorksheet("AR");
            var HR = workbook.AddWorksheet("HR");
            var OIT = workbook.AddWorksheet("OIT");

            // Add headers to the first row
            worksheet.Cell(1, 1).Value = "Job/Process Flow";
            worksheet.Cell(1, 2).Value = "Start Time";
            worksheet.Row(1).Style.Font.Bold = true;                                 // Make the entire row bold
            worksheet.Row(1).Style.Font.Underline = XLFontUnderlineValues.Single;    // Underline the entire row

            FA.Cell(1, 1).Value = "Job/Process Flow";
            FA.Cell(1, 2).Value = "Start Time";
            FA.Row(1).Style.Font.Bold = true;  
            FA.Row(1).Style.Font.Underline = XLFontUnderlineValues.Single;

            ST.Cell(1, 1).Value = "Job/Process Flow";
            ST.Cell(1, 2).Value = "Start Time";
            ST.Row(1).Style.Font.Bold = true; 
            ST.Row(1).Style.Font.Underline = XLFontUnderlineValues.Single;

            AR.Cell(1, 1).Value = "Job/Process Flow";
            AR.Cell(1, 2).Value = "Start Time";
            AR.Row(1).Style.Font.Bold = true;  
            AR.Row(1).Style.Font.Underline = XLFontUnderlineValues.Single;

            HR.Cell(1, 1).Value = "Job/Process Flow";
            HR.Cell(1, 2).Value = "Start Time";
            HR.Row(1).Style.Font.Bold = true; 
            HR.Row(1).Style.Font.Underline = XLFontUnderlineValues.Single;

            OIT.Cell(1, 1).Value = "Job/Process Flow";
            OIT.Cell(1, 2).Value = "Start Time";
            OIT.Row(1).Style.Font.Bold = true;  
            OIT.Row(1).Style.Font.Underline = XLFontUnderlineValues.Single;

            // Freeze the first row
            worksheet.SheetView.FreezeRows(1);
            FA.SheetView.FreezeRows(1);
            ST.SheetView.FreezeRows(1);
            AR.SheetView.FreezeRows(1);
            HR.SheetView.FreezeRows(1);
            OIT.SheetView.FreezeRows(1);

            // Set column to auto width
            worksheet.Column(1).Width = 36.45;
            worksheet.Column(2).Width = 20;

            FA.Column(1).Width = 30.86;
            FA.Column(2).Width = 20;

            ST.Column(1).Width = 34.57;
            ST.Column(2).Width = 20;

            AR.Column(1).Width = 28.57;
            AR.Column(2).Width = 20;

            HR.Column(1).Width = 31.30;
            HR.Column(2).Width = 20;

            OIT.Column(1).Width = 36.57;
            OIT.Column(2).Width = 20;

            // List of prefixes to check
            string[] validPrefixes = { "FA", "ST", "AR", "HR", "OIT" };

            // Start adding in data at row 
            int row = 2, rowFA = 2, rowST = 2, rowAR = 2, rowHR = 2, rowOIT = 2;

            // Create a dictionary to map the prefixes to their sheets and rows
            var prefixToSheet = new Dictionary<string, (IXLWorksheet sheet, int row, string lastDayOfWeek)>
            {
                { "FA", (FA, rowFA, string.Empty) },
                { "ST", (ST, rowST, string.Empty) },
                { "AR", (AR, rowAR, string.Empty) },
                { "HR", (HR, rowHR, string.Empty) },
                { "OIT", (OIT, rowOIT, string.Empty) }
            };

            // REGEX
            // \s* zero or more white space characters
            // \s+ one or more white space characters
            // \S+ one or more non-whitespace characters
            // \w+ one or more word characters (days of the week and month name)
            // \d+ is the day (in numbers)
            // \d{4} is the year (2025)
            // \d{2}:\d{2} is the time in HH:MM (09:45)
            // ? makes the group optional. Will match if it exists and will match if it is absent

            // Parenthesis sets the Capture Groups
            // (\S+) first Capture Group - name of chain or module
            // (\w+\s+\w+\s+\d+\s+\d{4}\s+\d{2}:\d{2}) second Capture Group - date and time

            // Capture these patterns
            // QUEUED       {Chain:OIT_C_ADASTRA_XFER} Thu Jan 30 2025 09:45 
            // {Module:OIT_M_GZPEMAL} Thu Jan 30 2025 10:39 EST5EDT (GMT-5.0) (Dls)
            string pattern = @"\s*(QUEUED\s+)?\{(Chain|Module):(\S+)\}\s+(\w+\s+\w+\s+\d+\s+\d{4}\s+\d{1,2}:\d{2})";

            // Read file line by line. Using StreamReader for larger files
            using (StreamReader reader = new StreamReader(latestFile))
            {
                string line;

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

                        //Console.WriteLine(currentDayOfWeek);

                        // If it's a "QUEUED" chain, process it
                        if (line.Contains("QUEUED") && type == "Chain")
                        {
                            //Console.WriteLine($"Chain: {name}");
                            //Console.WriteLine($"Date/Time: {dateTime}");

                            // Add data to Excel
                            worksheet.Cell(row, 1).Value = name;
                            worksheet.Cell(row, 2).Value = dateTime;
                            row++; // Move to next row

                            // Add each application to appropriate sheet
                            // Loop through the dictionary and check if the name starts with the prefix
                            foreach (var prefix in prefixToSheet.Keys)          // iterates over all the keys (FA, ST, etc) in the prefixToSheet dictionary
                            {
                                if (name.StartsWith(prefix))                      // Checks to see if the name contains one of the prefixes
                                {
                                    // Access the correct sheet and add data
                                    // tuple deconstruction allows you to store multiple values in a single object, and you can extract them individually
                                    // prefixToSheet[prefix] accesses the value (the tuple (IXLWorksheet sheet, int row)) associated with the current prefix in the dictionary
                                    // Ex: if prefix is OIT, then prefixToSheet["OIT'] might return a tuple like (OIT, 2)
                                    // Needed to add 'day' to keep track of each sheet's last day of the week
                                    var (sheet, row1, day) = prefixToSheet[prefix];  

                                    // Writes the name into the first column in the corresponding sheet
                                    sheet.Cell(row1, 1).Value = name;

                                    // Writes the dateTime into the second column in the correcponding sheet
                                    sheet.Cell(row1, 2).Value = dateTime;

                                    // See if it is a new day of the week. if strings match, skip this step
                                    if (day != currentDayOfWeek)
                                    {
                                        // Update dayOfWeek to be the new day to look for
                                        day = currentDayOfWeek;

                                        // Insert a row above twice
                                        sheet.Row(row1).InsertRowsAbove(2);

                                        // Split the dateTime variable to only get the date part
                                        string[] dateParts = dateTime.Split(' ');                   // Split the string into parts based on spaces
                                        string dateOnly = string.Join(" ", dateParts.Take(4));      // get onyl the first 4 parts of the date ex: "Thu Jan 30 2025"

                                        // Add the dateTime above in col 1
                                        // row1 + 1 is for adjustment since row1 is still set to where data was originally placed
                                        sheet.Cell(row1+1, 1).Value = dateOnly;

                                        // Bold and underline it
                                        sheet.Row(row1+1).Style.Font.Bold = true;
                                        sheet.Row(row1+1).Style.Font.Underline = XLFontUnderlineValues.Single;

                                        // Update row1 since now we have added two extra rows
                                        row1 = row1 + 2;
                                    }

                                    // Increment the row for the current sheet
                                    prefixToSheet[prefix] = (sheet, row1 + 1, day);
                                    break;
                                }
                            }
                        }
                        // If it's a module, process it. Checks to make sure the reading line does not say "INACTIVE" in front of it
                        else if (line.Contains("{Module:") && !line.Contains("INACTIVE"))
                        {
                            // For Modules (no QUEUED check needed)
                            //Console.WriteLine($"Module: {name}");
                            //Console.WriteLine($"Date/Time: {dateTime}");

                            // Add data to Excel
                            worksheet.Cell(row, 1).Value = name;
                            worksheet.Cell(row, 2).Value = dateTime;
                            row++; // Move to next row

                            // Add each application to appropriate sheet
                            // Loop through the dictionary and check if the name starts with the prefix
                            foreach (var prefix in prefixToSheet.Keys)
                            {
                                if (name.StartsWith(prefix))
                                {
                                    // Access the correct sheet and add data
                                    var (sheet, row1, day) = prefixToSheet[prefix];
                                    sheet.Cell(row1, 1).Value = name;
                                    sheet.Cell(row1, 2).Value = dateTime;

                                    // See if it is a new day of the week. if strings match, skip this step
                                    if (day != currentDayOfWeek)
                                    {
                                        // Update dayOfWeek to be the new day to look for
                                        day = currentDayOfWeek;

                                        // Insert a row above twice
                                        sheet.Row(row1).InsertRowsAbove(2);

                                        // Split the dateTime variable to only get the date part
                                        string[] dateParts = dateTime.Split(' ');                
                                        string dateOnly = string.Join(" ", dateParts.Take(4));

                                        // Add the dateTime above in col 1
                                        // row1 + 1 is for adjustment since row1 is still set to where data was originally placed
                                        sheet.Cell(row1 + 1, 1).Value = dateOnly;

                                        // Bold and underline it
                                        sheet.Row(row1 + 1).Style.Font.Bold = true;
                                        sheet.Row(row1 + 1).Style.Font.Underline = XLFontUnderlineValues.Single;

                                        // Update row1 since now we have added two extra rows
                                        row1 = row1 + 2;
                                    }

                                    // Increment the row for the current sheet
                                    prefixToSheet[prefix] = (sheet, row1 + 1, day);
                                    break;
                                }
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
                // Needed to launch file with appropriate application!
                Process.Start(new ProcessStartInfo(outputFilePath) { UseShellExecute = true });

                Console.WriteLine("\nExcel file created and opened successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while opening the file: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("No files with numbers found in the folder.");
        }


        

    }
}


