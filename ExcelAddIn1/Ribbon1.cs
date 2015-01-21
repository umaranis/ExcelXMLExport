using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace ExcelAddIn1
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelAddIn1.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnTextButton(Office.IRibbonControl control)
        {
            object missing = System.Type.Missing;
            System.Collections.Specialized.StringCollection columns = new System.Collections.Specialized.StringCollection();
            System.Windows.Forms.SaveFileDialog dlg = new System.Windows.Forms.SaveFileDialog();
            dlg.Filter = "Xml files (*.xml)|*.xml|All files (*.*)|*.*";
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                StreamWriter f = new StreamWriter(dlg.FileName);

                Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
                Excel.Range r;

                //Extracting the list of columns
                for (int i = 65; i <= 90; i++)
                {
                    r = sheet.get_Range(((char)i).ToString() + "1", missing);
                    string value = r.Value2;
                    if (value != null)
                    {
                        columns.Add(value);
                    }
                    else
                    {
                        break;
                    }
                }

                f.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                f.WriteLine("<"+sheet.Name+">");

                //traversing through the rows
                int rowNo = 2;
                string [] rowValues = new string[columns.Count];
                do
                {

                    for (int i = 65; i < 65 + columns.Count; i++)
                    {
                        r = sheet.get_Range(((char)i).ToString() + rowNo.ToString(), missing);
                        rowValues[i - 65] = Convert.ToString(r.Value2);
                    }

                    if (IsRowNull(rowValues))
                    {
                        break;
                    }
                    else
                    {
                        f.WriteLine("<Row>");
                        for (int j = 0; j < rowValues.Length; j++)
                        {
                            f.Write("<" + columns[j] + ">");
                            f.Write(rowValues[j]);
                            f.Write("</" + columns[j] + ">");
                        }
                        f.WriteLine("</Row>");
                        
                    }                    

                    rowNo++;
                } while (true);

                f.WriteLine("</" + sheet.Name + ">");
                f.Close();
            }
        }

        private bool IsRowNull(string[] rowValues)
        {
            for (int i = 0; i < rowValues.Length; i++)
            {
                if (rowValues[i] != null) return false;
            }
            return true;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
