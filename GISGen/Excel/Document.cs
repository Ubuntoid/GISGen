using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace GISGen.Excel
{
    public class Document
    {
        private ExcelPackage _excelPackage;
        public ExcelWorksheets Worksheets;
        //public List<ExcelWorksheet> WorksheetList;

        public Document()
        {
        }

        public Document(string fileName)
        {
            FromFile(fileName);
        }

        public void Save()
        {
            _excelPackage?.Save();
        }

        public void Close()
        {
            _excelPackage?.Dispose();
        }

        public void FromFile(string fileName)
        {
            
            var fileInfo = new FileInfo(fileName);
            _excelPackage = new ExcelPackage(fileInfo);
            
            
            //_excelPackage.Save();
            //Worksheets = GetWorksheets();
            Worksheets = _excelPackage.Workbook.Worksheets;

        }

       // private ExcelWorksheets GetWorksheets()
       // {
       //     while (true)
       //     {
       //         try
       //         {
       //             return _excelPackage.Workbook.Worksheets;
       //         }
       //         catch (Exception)
       //         {
       //             continue;
       //         }
       //     }
          
       //}
    }

}