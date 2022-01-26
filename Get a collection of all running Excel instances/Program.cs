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

            Application oExcel;

            ExcelAppCollection myApps = new ExcelAppCollection();
            Console.WriteLine("Session ID " + myApps.SessionID);
            //oExcel = myApps.PrimaryInstance;
            //var oExcels = new List<Process>();

            List<Process> oExcels = (List<Process>)myApps.GetProcesses();

            Console.WriteLine("Number of Excel processes found: " + oExcels.Count);
            Console.WriteLine();

            foreach (Process process in oExcels)
            {
                oExcel = myApps.FromProcess(process);
                Console.WriteLine("Process ID " + process.Id);
                if (oExcel != null)
                {
                    Console.WriteLine("Excel Workbooks count " + oExcel.Workbooks.Count);
                    Console.WriteLine();

                    foreach (Workbook workbook in oExcel.Workbooks)
                    {
                        Console.WriteLine("Process ID: " + process.Id + " Workbook Name: " + workbook.Name);
                    }
                }
                else
                {
                    Console.WriteLine("Excel is in task manager but not visible - not correctly closed?");
                }
                
                //TODO: dispose correctly excel object

                Console.WriteLine();

            }

            Console.ReadLine(); 

        }
    }
}
