using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;


namespace Get_a_collection_of_all_running_Excel_instances
{

    internal class Program
    {
        static void Main()
        {

            Application excel;

            ExcelAppCollection myApps = new ExcelAppCollection();
            Console.WriteLine("Session ID " + myApps.SessionID);
            excel = myApps.PrimaryInstance;
            //var oExcels = new List<Process>();

            List<Process> oExcels = (List<Process>)myApps.GetProcesses();

            Console.WriteLine("Number of Excel processes found: " + oExcels.Count);
            Console.WriteLine();

            foreach (Process process in oExcels)
            {
                excel = myApps.FromProcess(process);
                Console.WriteLine("Process ID " + process.Id);
                Console.WriteLine("Excel Workbooks count " + excel.Workbooks.Count);
                Console.WriteLine();

                foreach (Workbook workbook in excel.Workbooks)
                {
                    Console.WriteLine("Process ID: " + process.Id + " Workbook Name: " + workbook.Name);
                }

                Console.WriteLine();

            }

            Console.ReadLine(); 

        }
    }
}
