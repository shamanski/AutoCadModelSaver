using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AutoCAD_CSharp_plug_in2
{
   public class AcadApp
    {
        private Autodesk.AutoCAD.ApplicationServices.Document doc;
        private Editor ed;
        private Database db;
        [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
        static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);
        public AcadApp()
        {
            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            ed = doc.Editor;
            db = doc.Database;
        }
        public void PrintAll()
        {
            int i = 1;
            foreach (var block in GetBlockReferences(new string[] { "A4", "A3" }))
            {
                Print(block, $"{i++}.pdf");
            };
        }

        public void Print(BlockReference block, string fileName)
        {
            PlotLayout(GetWindow(block), fileName);
        }

        public IEnumerable<BlockReference> GetBlockReferences(IEnumerable<string> blockNames)
        {
            var selRes = Select();
            using (var t = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId idEnt in selRes.Value.GetObjectIds())
                {
                    if (idEnt.ObjectClass.Name == "AcDbBlockReference")
                    {
                        var blRef = t.GetObject(idEnt, OpenMode.ForRead) as BlockReference;
                        if (blockNames.Contains(GetEffectiveBlockName(blRef)))
                        {
                            yield return blRef;
                        }
                    }
                }
                t.Commit();
            }

        }

        private PromptSelectionResult Select()
        {
            var prOpt = new PromptSelectionOptions();
            prOpt.MessageForAdding = "Выделите регион для поиска листов";
            var selRes = ed.GetSelection(prOpt);
            if (selRes.Status != PromptStatus.OK)
                return null;
            return selRes;
        }

        private Extents2d GetWindow(BlockReference blRef)
        {
            var extents = blRef.Bounds;
            var ext1 = extents.Value.MinPoint;
            var ext2 = extents.Value.MaxPoint;
            Point3d first; Point3d second;
            first = ext1;
            second = ext2;
            ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 1)), rbTo = new ResultBuffer(new TypedValue(5003, 2));
            double[] firres = new double[] { 0, 0, 0 };
            double[] secres = new double[] { 0, 0, 0 };
            acedTrans(first.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres);
            acedTrans(second.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres);
            return new Extents2d(firres[0], firres[1], secres[0], secres[1]);
        }

        private string GetEffectiveBlockName(BlockReference blref)
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

        public void PlotLayout(Extents2d window, string fileName)
        {
            // Get the current document and database, and start a transaction

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                doc.LockDocument();
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

                        // Set the plot device to use
                        acPlSetVdr.SetPlotConfigurationName(acPlSet, "DWG to PDF.pc3", "ANSI_A_(8.50_x_11.00_Inches)");


                        // Set the plot info as an override since it will
                        // not be saved back to the layout
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

                                            acPlEng.BeginPlot(acPlProgDlg, null);

                                            // Define the plot output
                                            acPlEng.BeginDocument(acPlInfo, doc.Name, null, 1, true, fileName);
                                            using (PlotPageInfo acPlPageInfo = new PlotPageInfo())
                                            {
                                                acPlEng.BeginPage(acPlPageInfo, acPlInfo, true, null);
                                            }

                                            acPlEng.BeginGenerateGraphics(null);
                                            acPlEng.EndGenerateGraphics(null);
                                            acPlEng.EndPage(null);
                                            acPlEng.EndDocument(null);
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
