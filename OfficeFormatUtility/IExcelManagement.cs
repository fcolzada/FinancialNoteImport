using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace OfficeFormatUtility
{
    interface IExcelManagement
    {
        ExcelDocument OpenDocument(String path);
        DataTable GetFirstSheet(string path);
    }
}
