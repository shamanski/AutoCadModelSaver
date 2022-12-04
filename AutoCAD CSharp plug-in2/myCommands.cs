
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows;
using System.Windows.Controls;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

[assembly: CommandClass(typeof(AutoCAD_CSharp_plug_in2.MyCommands))]

namespace AutoCAD_CSharp_plug_in2
{

    public class MyCommands
    {
        [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
        static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);
        static PaletteSet _ps = null;

        [CommandMethod("WPFP")]
        public void ShowWPFPalette()
        {
            if (_ps == null)

            {

                _ps = new PaletteSet("WPF Palette");

                _ps.Size = new System.Drawing.Size(400, 600);

                _ps.DockEnabled =

                  (DockSides)((int)DockSides.Left + (int)DockSides.Right);


                UserControl1 uc = new UserControl1();

                _ps.AddVisual("AddVisual", uc);



                UserControl1 uc2 = new UserControl1();

                ElementHost host = new ElementHost();

                host.AutoSize = true;

                host.Dock = DockStyle.Fill;

                host.Child = uc2;

                _ps.Add("Add ElementHost", host);

            }

            _ps.KeepFocus = true;
            _ps.Visible = true;
        }

        [CommandMethod("1111")]
        public void MyCommand() // This method can have any name
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            List<ObjectId> ids = new List<ObjectId>();

            Editor ed = doc.Editor;
            Database db = doc.Database;
            var prOpt = new PromptSelectionOptions();
            prOpt.MessageForAdding = "Выбор блоков панелей";
            var selRes = ed.GetSelection(prOpt);
            if (selRes.Status != PromptStatus.OK)
                return;

            // Фильтр блоков панелей (по имени блоков).
            using (var t = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId idEnt in selRes.Value.GetObjectIds())
                {
                    if (idEnt.ObjectClass.Name == "AcDbBlockReference")
                    {
                        var blRef = t.GetObject(idEnt, OpenMode.ForRead) as BlockReference;
                        if (GetEffectiveBlockName(blRef) == "A4")
                        {
                            ids.Add(idEnt);
                            var extents = blRef.Bounds;
                            var ext1 = extents.Value.MinPoint;
                            var ext2 = extents.Value.MaxPoint;
                            Point3d first; Point3d second;
                            first = ext1;
                            second = ext2;
                            ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 1)), rbTo = new ResultBuffer(new TypedValue(5003, 2));
                            double[] firres = new double[] { 0, 0, 0 };
                            double[] secres = new double[] { 0, 0, 0 };
                            // Transform the first point...
                            acedTrans(first.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres);
                            // ... and the second
                            acedTrans(second.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres);
                            // We can safely drop the Z-coord at this stage
                            Extents2d window = new Extents2d(firres[0], firres[1], secres[0], secres[1]);
                            PlotLayout(window);
                            ed.WriteMessage(blRef.Position.ToString());
                        }
                    }
                }
                t.Commit();
            }
        }

        public static string GetEffectiveBlockName(BlockReference blref)
        {
            string res = string.Empty;
            if (blref.IsDynamicBlock)
            {
                var btr = blref.DynamicBlockTableRecord.GetObject(OpenMode.ForRead) as BlockTableRecord;
                res = btr.Name;
                btr.Close();
            }
            else
            {
                res = blref.Name;
            }
            return res;
        }

        [CommandMethod("PlotLayout")]
        public static void PlotLayout(Extents2d window)
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Reference the Layout Manager
                LayoutManager acLayoutMgr = LayoutManager.Current;

                // Get the current layout and output its name in the Command Line window
                Layout acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout),
                                                    OpenMode.ForRead) as Layout;
                // Get the PlotInfo from the layout
                using (PlotInfo acPlInfo = new PlotInfo())
                {
                    acPlInfo.Layout = acLayout.ObjectId;

                    // Get a copy of the PlotSettings from the layout
                    using (PlotSettings acPlSet = new PlotSettings(acLayout.ModelType))
                    {
                        acPlSet.CopyFrom(acLayout);

                        // Update the PlotSettings object
                        PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                        // Set the plot type
                        acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                        acPlSetVdr.SetPlotWindowArea(acPlSet, window);
                        // Set the plot scale
                        acPlSetVdr.SetUseStandardScale(acPlSet, true);
                        acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);

                        // Center the plot
                        acPlSetVdr.SetPlotCentered(acPlSet, true);

                        acPlSetVdr.SetPlotConfigurationName(acPlSet, "DWG to PDF.pc3", "ANSI_A_(8.50_x_11.00_Inches)");
                        ///TMP
                        Stream s = File.Open(@"c:\plot\myplot.txt", FileMode.Create);
                        BinaryFormatter b = new BinaryFormatter();
                        b.Serialize(s, acPlSetVdr.GetCanonicalMediaNameList(acPlSet));
                        s.Close();
                        
                        acPlInfo.OverrideSettings = acPlSet;

                        // Validate the plot info
                        using (PlotInfoValidator acPlInfoVdr = new PlotInfoValidator())
                        {
                            acPlInfoVdr.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                            acPlInfoVdr.Validate(acPlInfo);

                            // Check to see if a plot is already in progress
                            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                            {
                                using (PlotEngine acPlEng = PlotFactory.CreatePublishEngine())
                                {
                                    // Track the plot progress with a Progress dialog
                                    using (PlotProgressDialog acPlProgDlg = new PlotProgressDialog(false, 1, true))
                                    {
                                        using ((acPlProgDlg))
                                        {
                                            // Define the status messages to display 
                                            // when plotting starts
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Plot Progress");
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");

                                            // Set the plot progress range
                                            acPlProgDlg.LowerPlotProgressRange = 0;
                                            acPlProgDlg.UpperPlotProgressRange = 100;
                                            acPlProgDlg.PlotProgressPos = 0;

                                            // Display the Progress dialog
                                            acPlProgDlg.OnBeginPlot();
                                            acPlProgDlg.IsVisible = true;

                                            // Start to plot the layout
                                            acPlEng.BeginPlot(acPlProgDlg, null);

                                            // Define the plot output
                                            acPlEng.BeginDocument(acPlInfo, acDoc.Name, null, 1, true, @"c:\plot\myplot.pdf");

                                            // Display information about the current plot
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.Status, "Plotting: " + acDoc.Name + " - " + acLayout.LayoutName);

                                            // Set the sheet progress range
                                            acPlProgDlg.OnBeginSheet();
                                            acPlProgDlg.LowerSheetProgressRange = 0;
                                            acPlProgDlg.UpperSheetProgressRange = 100;
                                            acPlProgDlg.SheetProgressPos = 0;

                                            // Plot the first sheet/layout
                                            using (PlotPageInfo acPlPageInfo = new PlotPageInfo())
                                            {
                                                acPlEng.BeginPage(acPlPageInfo, acPlInfo, true, null);
                                            }

                                            acPlEng.BeginGenerateGraphics(null);
                                            acPlEng.EndGenerateGraphics(null);

                                            // Finish plotting the sheet/layout
                                            acPlEng.EndPage(null);
                                            acPlProgDlg.SheetProgressPos = 100;
                                            acPlProgDlg.OnEndSheet();

                                            // Finish plotting the document
                                            acPlEng.EndDocument(null);

                                            // Finish the plot
                                            acPlProgDlg.PlotProgressPos = 100;
                                            acPlProgDlg.OnEndPlot();
                                            acPlEng.EndPlot(null);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

}
