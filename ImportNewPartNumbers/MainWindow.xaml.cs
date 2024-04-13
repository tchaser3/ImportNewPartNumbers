/* Title:           Updating Part Information
 * Date:            2-10-20
 * Author:          Terry Holmes
 * 
 * Description:     This is used for updating part numbers */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NewPartNumbersDLL;
using NewEventLogDLL;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportNewPartNumbers
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //settting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        PartNumberClass ThePartNumberClass = new PartNumberClass();
        EventLogClass TheEventLogClass = new EventLogClass();

        FindPartFromMasterPartListByJDEPartNumberDataSet TheFindPartFromMasterPartListByJDEPartNumberDataSet = new FindPartFromMasterPartListByJDEPartNumberDataSet();
        FindPartFromMasterPartListByPartNumberDataSet TheFindPartFromMasterPartListByPartNumberDataset = new FindPartFromMasterPartListByPartNumberDataSet();
        FindPartByPartNumberDataSet TheFindPartByPartNumberDataSet = new FindPartByPartNumberDataSet();
        FindPartByJDEPartNumberDataSet TheFindPartByJDEPartNumberDataSet = new FindPartByJDEPartNumberDataSet();
        ImportExcelPartsDataSet TheImportExcelPartsDataSet = new ImportExcelPartsDataSet();
        PartInformationDataSet ThePartInformationDataSet = new PartInformationDataSet();

        //setting global variables
        int gintPartID;
        string gstrOldPartNumber;
        string gstrOldJDEPartNumber;
        string gstrNewPartNumber;
        string gstrNewJDEPartNumber;
        string gstrDescrpition;
        bool gblnProcess;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.CloseTheProgram();
        }
        private bool ImportExcelSheet()
        {
            bool blnFatalError = false;
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            string strPartNumber;
            string strJDEPartNumber;
            string strPartDescription;
            DateTime datPunchDate = DateTime.Now;
            int intPartID;
            int intRecordsReturned;

            try
            {
                TheImportExcelPartsDataSet.importexcelparts.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 5000; intCounter < intNumberOfRecords; intCounter++)
                {
                        
                    strPartNumber = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2);
                    strJDEPartNumber = Convert.ToString((range.Cells[intCounter, 3] as Excel.Range).Value2);
                    strPartDescription = Convert.ToString((range.Cells[intCounter, 2] as Excel.Range).Value2);

                    TheFindPartByPartNumberDataSet = ThePartNumberClass.FindPartByPartNumber(strPartNumber);

                    intPartID = intCounter * -1;

                    intRecordsReturned = TheFindPartByPartNumberDataSet.FindPartByPartNumber.Rows.Count;

                    if(intRecordsReturned > 0)
                    {
                        intPartID = TheFindPartByPartNumberDataSet.FindPartByPartNumber[0].PartID;
                    }

                    ImportExcelPartsDataSet.importexcelpartsRow NewPartRow = TheImportExcelPartsDataSet.importexcelparts.NewimportexcelpartsRow();

                    NewPartRow.PartID = intPartID;
                    NewPartRow.JDEPartNumber = strJDEPartNumber;
                    NewPartRow.PartNumber = strPartNumber;
                    NewPartRow.PartDescription = strPartDescription;

                    TheImportExcelPartsDataSet.importexcelparts.Rows.Add(NewPartRow);
                }

                PleaseWait.Close();
                dgrResults.ItemsSource = TheImportExcelPartsDataSet.importexcelparts;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import New Part Number // Main Window // Import Excel Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }


            return blnFatalError;
        }

        private void btnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            //setting local variables
            bool blnFatalError = false;
            int intCounter;
            int intNumberOfRecords;
            int intRecordsReturned;

            PleaseWait PleaseWait = new PleaseWait();
            PleaseWait.Show();

            try
            {
                TheImportExcelPartsDataSet.importexcelparts.Rows.Clear();
                ThePartInformationDataSet.partinformation.Rows.Clear();

                blnFatalError = ImportExcelSheet();

                if (blnFatalError == true)
                    throw new Exception();

                intNumberOfRecords = TheImportExcelPartsDataSet.importexcelparts.Rows.Count - 1;

                dgrResults.ItemsSource = TheImportExcelPartsDataSet.importexcelparts;

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import New Part Numbers // Main Window // Import Excel Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

            PleaseWait.Close();

        }
        private bool AddRecordsToDB()
        {
            bool blnFatalError = false;

            try
            {
                PartInformationDataSet.partinformationRow NewPartRow = ThePartInformationDataSet.partinformation.NewpartinformationRow();

                NewPartRow.NewJDEPartNumber = gstrNewJDEPartNumber;
                NewPartRow.PartID = gintPartID;
                NewPartRow.NewPartNumber = gstrNewPartNumber;
                NewPartRow.OldJDEPartNumber = gstrOldJDEPartNumber;
                NewPartRow.OldPartNumber = gstrOldPartNumber;
                NewPartRow.PartDescription = gstrDescrpition;
                NewPartRow.Process = gblnProcess;

                ThePartInformationDataSet.partinformation.Rows.Add(NewPartRow);
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import New Part Numbers // Main Window // Add Records To DG " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            int intPartID;
            string strPartNumber;
            string strJDEPartNumber;
            string strPartDesciption;
            int intCounter;
            int intNumberOfRecords;
            bool blnFatalError = false;
            bool blnProcess;

            try
            {
                intNumberOfRecords = TheImportExcelPartsDataSet.importexcelparts.Rows.Count;

                for (intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                {
                    intPartID = TheImportExcelPartsDataSet.importexcelparts[intCounter].PartID;
                    strPartNumber = TheImportExcelPartsDataSet.importexcelparts[intCounter].PartNumber;
                    strJDEPartNumber = TheImportExcelPartsDataSet.importexcelparts[intCounter].JDEPartNumber;
                    strPartDesciption = TheImportExcelPartsDataSet.importexcelparts[intCounter].PartDescription;

                  
                    if (intPartID < 1)
                    {
                         blnFatalError = ThePartNumberClass.InsertPartIntoPartNumbers(intPartID, strPartNumber, strJDEPartNumber, strPartDesciption, 0);

                         if (blnFatalError == true)
                             throw new Exception();
                    }
                    else if (intPartID > 0)
                    {
                        blnFatalError = ThePartNumberClass.UpdatePartInformation(intPartID, strJDEPartNumber, strPartDesciption, true, 0);

                        if (blnFatalError == true)
                            throw new Exception();
                    }
                        
                    
                }

                TheMessagesClass.InformationMessage("Part Numbers are Updated");


                ThePartInformationDataSet.partinformation.Rows.Clear();

                dgrResults.ItemsSource = ThePartInformationDataSet.partinformation;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import New Part Numbers // Main Window // Process Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
