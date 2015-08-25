using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;

namespace ALITECS.SWx.SESE
{
    /// <summary>
    /// Summary description for SESE.
    /// </summary>
    [Guid("ff6692f4-1c7b-4058-b865-8c45967bca16"), ComVisible(true)]
    [SwAddin(
        Description = "SolidWorks Easy STL Export",
        Title = "SESE",
        LoadAtStartup = true
        )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables

        private int addinID = 0;

        private ISldWorks iSwApp = null;

        // Public Properties
        public ISldWorks SwApp
        {
            get { return iSwApp; }
        }

        #endregion Local Variables

        #region SolidWorks Registration

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute

            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion Get Custom Attribute: SwAddinAttribute

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion SolidWorks Registration

        #region ISwAddin Implementation

        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            iSwApp.AddMenuItem4((int)swDocumentTypes_e.swDocPART, addinID, "Export STLs", -1, "ExportSTLs", "", "", "");

            return true;
        }

        public bool DisconnectFromSW()
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }

        #endregion ISwAddin Implementation

        #region UI Callbacks

        public void ExportSTLs()
        {
            int errors = 0;
            int warnings = 0;

            ModelDoc2 mdoc = iSwApp.ActiveDoc;
            PartDoc pdoc = iSwApp.ActiveDoc;

            SelectData swSelectData = default(SelectData);
            swSelectData = (SelectData)mdoc.SelectionManager.CreateSelectData();

            // set export coordinate system to one which is named "STL"
            mdoc.Extension.SetUserPreferenceString((int)swUserPreferenceStringValue_e.swFileSaveAsCoordinateSystem, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, "STL");

            // get all solid bodies of the part
            object[] bodies = pdoc.GetBodies2(0, true);

            // the the full path of the opened document
            string path = mdoc.GetPathName();

            if (bodies.Length > 1) // if there are more than one body in the part
            {
                foreach (Body2 body in bodies)
                {
                    body.Select2(false, swSelectData); // select the body

                    // save an STL in the same directory as the document but named after the body
                    mdoc.Extension.SaveAs(Path.GetDirectoryName(path) + "\\" + body.Name + ".stl", 0, 3, null, errors, warnings);
                }
            }
            else // if there is only one body
            {
                // save an STL in the same directory as the document named after the document
                mdoc.Extension.SaveAs(Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + ".stl", 0, 3, null, errors, warnings);
            }
        }

        #endregion UI Callbacks
    }
}