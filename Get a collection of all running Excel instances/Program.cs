using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Get_a_collection_of_all_running_Excel_instances
{

    internal class Program
    {
        static void Main()
        {

            ExcelAppCollection myApps = new ExcelAppCollection();
            Console.WriteLine("Session ID " + myApps.SessionID);
            //oExcel = myApps.PrimaryInstance;
            
            Console.WriteLine("Getting Excel processes");
            List<Process> ExcelProcesses = (List<Process>)myApps.GetProcesses();

            Console.WriteLine("Number of Excel processes found: {0}", ExcelProcesses.Count);
            Console.WriteLine();

            Application ExcelAppication;

            foreach (Process process in ExcelProcesses)
            {
                Console.WriteLine("Process ID {0}" , process.Id);
                ExcelAppication = myApps.FromProcess(process);
                
                if (ExcelAppication != null)
                {
                    Console.WriteLine("Excel Workbooks count {0}" , ExcelAppication.Workbooks.Count);
                    Console.WriteLine();


                    foreach (Workbook oWorkbook in ExcelAppication.Workbooks)
                    {

                        Console.WriteLine("Closing workbook {0}", oWorkbook.Name);
                        if (oWorkbook.Path != string.Empty)
                        {
                            oWorkbook.Close(true, Missing.Value, Missing.Value);
                        }
                        else
                        {

                            string savePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), oWorkbook.Name + ".xlsx");
                            Console.WriteLine("Workbook first time opened - savied {0}", savePath);
                            oWorkbook.SaveAs(savePath, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                            oWorkbook.Close(true, Missing.Value, Missing.Value);
                        }

                        Console.WriteLine("Releasing workbook object");
                        ReleaseAll(oWorkbook);

                    }

                    Console.WriteLine("Releasing Excel object {0}", process.Id);
                    ExcelAppication.Quit();
                    ReleaseAll(ExcelAppication);
                    Console.WriteLine();

                }
                else
                {
                    Console.WriteLine("Excel is in task manager but not visible. Kill it with fire!");
                    process.Kill();
                    Console.WriteLine();
                }

            }
        
            Console.ReadLine(); 
        }

        static void ReleaseAll(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }



    }

}
