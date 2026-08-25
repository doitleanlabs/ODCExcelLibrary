using OutSystems.ExternalLibraries.SDK;

namespace OutSystems.ExternalLib.Excel {

        [OSInterface(Name = "Excel_Library_Custom", Description = "Library for Microsoft Excel", IconResourceName = "ExcelLibraryCustom.resources.excel.png")]
        public interface IExcelLibrary
        {

                // ============================================================
                // Interface For Workbook
                // ============================================================

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Workbook_Create", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Create a New Workbook Excel File")]
                public byte[] Workbook_Create(Worksheet[] worksheets, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Workbook_SetProperties", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Set Workbook Document Properties")]
                public byte[] Workbook_SetProperties(byte[] excelBinary, WorkbookProperties workbookProperties, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "Worksheets", OriginalName = "Workbook_GetWorksheet", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Get All Worksheets Properties in the Spreadsheet")]
                public WorksheetProperties[] Workbook_GetWorksheet(byte[] excelBinary, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);


                // ============================================================
                // Interface For Worksheet
                // ============================================================

                [OSAction(ReturnName = "Worksheet", OriginalName = "Worksheet_GetProperties", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Get Worksheet Properties in the Spreadsheet")]
                public WorksheetProperties Worksheet_GetProperties(byte[] excelBinary, string sheetName, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Worksheet_Add", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Add Single Worksheet on Existing Excel File")]
                public byte[] Worksheet_Add(byte[] excelBinary, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Worksheet_AddList", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Add Multiple Worksheets on Existing Excel File")]
                public byte[] Worksheet_AddList(byte[] excelBinary, Worksheet[] worksheets, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Worksheet_AutofitColumns", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Set Autofit Cell Column")]
                public byte[] Worksheet_AutofitColumns(byte[] excelBinary, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Worksheet_Calculate", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Trigger to Calculate Formula on Specific Worksheet")]
                public byte[] Worksheet_Calculate(byte[] excelBinary, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Worksheet_Protect", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Protect Workbook Excel File")]
                public byte[] Worksheet_Protect(byte[] excelBinary, string password, bool? isAllowAutoFilter = false, bool? isAllowDeleteColumns = false, bool? isAllowDeleteRows = false,
                bool? isAllowEditObject = false, bool? isAllowFormatCells = false, bool? isAllowFormatColumns = false, bool? isAllowFormatRows = false, bool? isAllowInsertColumns = false,
                bool? isAllowInsertHyperlinks = false, bool? isAllowInsertRows = false, bool? isAllowPivotTables = false, bool? isAllowSelectLockedCells = false, bool? isAllowSelectUnLockedCells = false,
                bool? isAllowSort = false, bool? isProtected = false, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Worksheet_AddAutoFilter", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Add Automatic Filter on Existing Excel File")]
                public byte[] Worksheet_AddAutoFilter(byte[] excelBinary, int startCellRow = 0, int startCellColumn = 0, int endCellRow = 0, int endCellColumn = 0, string? cellName = null, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Worksheet_Delete", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Delete Worksheet")]
                public byte[] Worksheet_Delete(byte[] excelBinary, int sheetIndex = 0, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Worksheet_Rename", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Rename Worksheet")]
                public byte[] Worksheet_Rename(byte[] excelBinary, string newSheetName, int sheetIndex = 0, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Worksheet_Hide_Show", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Hide or Show Worksheet")]
                public byte[] Worksheet_Hide_Show(byte[] excelBinary, int sheetIndex = 0, string? sheetName = null, bool isShow = false, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFileResult", OriginalName = "Worksheet_Copy", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Copy Worksheet from another Excel File")]
                public byte[] Worksheet_Copy(byte[] ExcelSource, [OSParameter(DataType = OSDataType.Text, Description = "Comma separated supported for multiple sheet name")] string SourceSheetName, byte[] ExcelDestination, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);


                // ============================================================
                // Interface For Cell
                // ============================================================

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Cell_Write", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Write Value to Cell")]
                public byte[] Cell_Write(byte[] excelBinary, CellWrite[] cellWrites, CellCopy? cellCopy = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "CellValue", OriginalName = "Cell_Read", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Read Value from Cell")]
                public string Cell_Read(byte[] excelBinary, [OSParameter(DataType = OSDataType.Integer, Description = "Row Index start from 1")] int cellRow = 0, [OSParameter(DataType = OSDataType.Integer, Description = "Column Index start from 1")] int cellColumn = 0, string? cellName = null, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Cell_Merge", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Merge Cell")]
                public byte[] Cell_Merge(byte[] excelBinary, CellMerge[] cellMerges, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Cell_UnMerge", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "UnMerge Cell")]
                public byte[] Cell_UnMerge(byte[] excelBinary, CellMerge[] cellUnMerges, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Cell_Copy", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Copy Single Cell Value to Single / Multiple Destination Cells")]
                public byte[] Cell_Copy(byte[] excelBinary, CellCopy cellCopy, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "Cells", OriginalName = "Cell_FindByValue", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Find Cells for specific Text Value")]
                public CellFindResult[] Cell_FindByValue(byte[] excelBinary, string cellValue, bool isContain = false, string? cellRange = null, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Cell_Write_RichText", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Write Value to Cell in RichText Format")]
                public byte[] Cell_Write_RichText(byte[] excelBinary, CellWriteRichText[] cellWriteRichTexts, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);


                // ============================================================
                // Interface For Comments
                // ============================================================

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Comment_Add", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Add Comments")]
                public byte[] Comment_Add(byte[] excelBinary, CommentAdd[] commentAdds, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Comment_Delete", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Delete Comments")]
                public byte[] Comment_Delete(byte[] excelBinary, CommentDelete[] commentDeletes, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                // ============================================================
                // Interface For Column
                // ============================================================

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Column_Delete", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Delete Column on specific Column Index")]
                public byte[] Column_Delete(byte[] excelBinary, [OSParameter(DataType = OSDataType.Integer, Description = "Column Index start from 1")] int colIndex, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Column_Hide_Show", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Hide or Show Column on specific Column Index")]
                public byte[] Column_Hide_Show(byte[] excelBinary, [OSParameter(DataType = OSDataType.Integer, Description = "Column Index start from 1")] int colIndex, bool isShow = false, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Column_Insert", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Inserts New Columns into the Spreadsheet. Existing Columns below the Position are Shifted Down.")]
                public byte[] Column_Insert(byte[] excelBinary, [OSParameter(DataType = OSDataType.Integer, Description = "Column Index start from 1")] int colIndex, int colNewAdd = 1, int colWidth = 64, bool isCopyFormatFromSource = false, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Column_Width", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Set Column Width on specific Column Index")]
                public byte[] Column_Width(byte[] excelBinary, [OSParameter(DataType = OSDataType.Integer, Description = "Column Index start from 1")] int colIndex, int colWidth = 64, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                // ============================================================
                // Interface For Row
                // ============================================================

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Row_Delete", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Delete Row on specific Row Index")]
                public byte[] Row_Delete(byte[] excelBinary, [OSParameter(DataType = OSDataType.Integer, Description = "Row Index start from 1")] int rowIndex, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Row_Hide_Show", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Hide or Show Row on specific Row Index")]
                public byte[] Row_Hide_Show(byte[] excelBinary, [OSParameter(DataType = OSDataType.Integer, Description = "Row Index start from 1")] int rowIndex, bool isShow = false, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Row_Insert", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Inserts New Rows into the Spreadsheet. Existing Rows below the Position are Shifted Down.")]
                public byte[] Row_Insert(byte[] excelBinary, [OSParameter(DataType = OSDataType.Integer, Description = "Row Index start from 1")] int rowIndex, int rowNewAdd = 1, int rowHeight = 20, bool isCopyFormatFromSource = false, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Row_Height", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Set Row Height on specific Row Index")]
                public byte[] Row_Height(byte[] excelBinary, [OSParameter(DataType = OSDataType.Integer, Description = "Row Index start from 1")] int rowIndex, int rowHeight = 20, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Row_AutoHeight", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Set Row Auto Height")]
                public byte[] Row_AutoHeight(byte[] excelBinary, [OSParameter(DataType = OSDataType.Integer, Description = "Row Index start from 1")] int rowStart, int rowEnd, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                // ============================================================
                // Public Method Implementation Interface - Data Validations
                // ============================================================

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Data_Validation_Integer", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Data Validation Integer")]
                public byte[] Data_Validation_Integer(byte[] excelBinary, CellDataValidation cellDataValidation, DataValidation dataValidation, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Data_Validation_Decimal", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Data Validation Decimal")]
                public byte[] Data_Validation_Decimal(byte[] excelBinary, CellDataValidation cellDataValidation, DataValidation dataValidation, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Data_Validation_List", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Data Validation with Dropdown")]
                public byte[] Data_Validation_List(byte[] excelBinary, CellDataValidation cellDataValidation, DataValidationListItem dataValidationListItem, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                // ============================================================
                // Interface For Range Functions
                // ============================================================

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Range_Format", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Cell Range Formating")]
                public byte[] Range_Format(byte[] excelBinary, RangeFormat[] rangeFormats, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Range_BorderFormat", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Cell Range Border Formatting")]
                public byte[] Range_BorderFormat(byte[] excelBinary, RangeBorderFormat[] rangeBorderFormats, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "RangeCellValue", OriginalName = "Range_CellRead", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Read Cell Value with Cell Range")]
                public RangeCellValue[] Range_CellRead(byte[] excelBinary, RangeCellRead[] rangeCellReads, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "CellRange", OriginalName = "Range_FromAddress", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Convert Given Address to the Cell Range format")]
                public CellRange Range_FromAddress([OSParameter(DataType = OSDataType.Text, Description = "Text address, e.g. AB47 or A11:AB47")] string Address);


                // ============================================================
                // Interface For Miscellaneous
                // ============================================================

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Data_WriteJSON", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Load JSON Format into Excel Table")]
                public byte[] Data_WriteJSON(byte[] excelBinary, DataWriteJSON[] dataWriteJSONs, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(OriginalName = "Data_WriteJSON_URL", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Load JSON Format into Excel Table from serialized payload URL")]
                public void Data_WriteJSON_URL(
                        byte[] excelBinary,
                        [OSParameter(DataType = OSDataType.Text, OriginalName = "GET_BINARY_URL", Description = "URL to download serialized DataWriteJSON array")]
                        string GET_BINARY_URL,
                        [OSParameter(DataType = OSDataType.Text, OriginalName = "UPDATE_BINARY_URL", Description = "Optional URL to upload the updated Excel binary with HTTP PUT")]
                        string UPDATE_BINARY_URL,
                        [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")]
                        string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFile", OriginalName = "Image_Insert", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Inserts Images into the Spreadsheet")]
                public byte[] Image_Insert(byte[] excelBinary, byte[] imageFile, int imageSizePercent = 100, int imageWidth = 0, int imageHeight = 0, int cellRow = 0, int cellColumn = 0, string? cellName = null, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ExcelFileOutput", OriginalName = "Excel_Merge", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Merge Excel Files")]
                public byte[] Excel_Merge(ExcelMerge[] ExcelFiles, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ImageFiles", OriginalName = "Image_GetAll_OverCell", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Get All Images (Over Cell Only) from the Spreadsheet")]
                public ExcelImages[] Image_GetAll_OverCell(byte[] excelBinary, [OSParameter(DataType = OSDataType.Text, Description = "Empty SheetName = All Worksheet")] string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ImageFiles", OriginalName = "Image_GetAll_InCell", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Get All Images (In Cell Only) from the Spreadsheet")]
                public ExcelImages[] Image_GetAll_InCell(byte[] excelBinary, [OSParameter(DataType = OSDataType.Text, Description = "Empty SheetName = All Worksheet")] string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);

                [OSAction(ReturnName = "ImageFile", OriginalName = "Image_Get", IconResourceName = "ExcelLibraryCustom.resources.excel.png", Description = "Get Image from the Spreadsheet")]
                public ExcelImages Image_Get(byte[] excelBinary, [OSParameter(DataType = OSDataType.Integer, Description = "Row Index start from 1")] int cellRow = 0, [OSParameter(DataType = OSDataType.Integer, Description = "Column Index start from 1")] int cellColumn = 0, [OSParameter(DataType = OSDataType.Text, Description = "Single Cell Name. e.g. AB47")] string? cellName = null, string? sheetName = null, [OSParameter(Description = "Optional EPPlus commercial license key. Leave empty to use EPPlus under the free non-commercial license.")] string? licenseKey = null);
        }
}