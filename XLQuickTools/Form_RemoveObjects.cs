using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace XLQuickTools
{
    public partial class RemoveObjectsForm : Form
    {
        public RemoveObjectsForm()
        {
            InitializeComponent();
        }

        // On Load
        private void RemoveObjectsForm_Load(object sender, EventArgs e)
        {
            // Check all options on load
            CbShapes.Checked = true;
            CbCharts.Checked = true;
            CbActiveX.Checked = true;
            CbFormControls.Checked = true;
            CbComments.Checked = true;
            CbNotes.Checked = true;
        }

        // Cancel Button
        private void RemoveObjectsForm_Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        // Ok Button
        private void RemoveObjectsForm_Ok_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;
            if (activeSheet == null) return;

            try
            {
                excelApp.ScreenUpdating = false;
                excelApp.EnableEvents = false;

                // Comments and Notes must be removed BEFORE shapes

                // Modern Threaded Comments (Excel 365)
                if (CbComments.Checked)
                {
                    RemoveThreadedComments(activeSheet);
                }

                // Legacy Notes
                if (CbNotes.Checked)
                {
                    RemoveNotes(activeSheet);
                }

                // Shapes, Charts, ActiveX Controls, Form Controls
                RemoveShapes(activeSheet);

                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error removing objects:\r\n\r\n" + ex.Message,
                    "Remove Objects",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                // Restore Excel settings
                excelApp.ScreenUpdating = true;
                excelApp.EnableEvents = true;
            }
        }

        // Remove Shapes
        private void RemoveShapes(Excel.Worksheet worksheet)
        {
            if (!CbShapes.Checked &&
                !CbCharts.Checked &&
                !CbActiveX.Checked &&
                !CbFormControls.Checked)
            {
                return;
            }

            for (int i = worksheet.Shapes.Count; i >= 1; i--)
            {
                Excel.Shape shape = null;

                try
                {
                    shape = worksheet.Shapes.Item(i);

                    Office.MsoShapeType shapeType = shape.Type;

                    switch (shapeType)
                    {
                        // Charts
                        case Office.MsoShapeType.msoChart:

                            if (CbCharts.Checked)
                            {
                                shape.Delete();
                            }

                            break;

                        // ActiveX Controls
                        case Office.MsoShapeType.msoOLEControlObject:

                            if (CbActiveX.Checked)
                            {
                                shape.Delete();
                            }

                            break;

                        // Embedded OLE Objects
                        case Office.MsoShapeType.msoEmbeddedOLEObject:

                            if (CbActiveX.Checked)
                            {
                                shape.Delete();
                            }

                            break;

                        // Form Controls
                        case Office.MsoShapeType.msoFormControl:

                            if (CbFormControls.Checked)
                            {
                                shape.Delete();
                            }

                            break;

                        // Notes and Comments
                        case Office.MsoShapeType.msoComment:

                            break;

                        // Everything Else
                        default:

                            if (CbShapes.Checked)
                            {
                                shape.Delete();
                            }

                            break;
                    }
                }
                catch
                {
                    // Processing the remaining objects
                }
                finally
                {
                    ReleaseComObject(shape);
                }
            }
        }

        // Remove Threaded Comments (Excel 365)
        private void RemoveThreadedComments(Excel.Worksheet worksheet)
        {
            object threadedComments = null;

            try
            {
                /*
                 * CommentsThreaded is not exposed by some versions of
                 * Microsoft.Office.Interop.Excel.
                 *
                 * Excel 365 itself supports it, so use COM late binding
                 * instead of requiring a newer Interop assembly.
                 */

                threadedComments = worksheet.GetType().InvokeMember(
                    "CommentsThreaded",
                    BindingFlags.GetProperty,
                    null,
                    worksheet,
                    null);

                if (threadedComments == null)
                {
                    return;
                }

                dynamic comments = threadedComments;

                // Delete backwards so the collection can safely shrink
                for (int i = comments.Count; i >= 1; i--)
                {
                    dynamic comment = null;

                    try
                    {
                        comment = comments.Item(i);

                        if (comment != null)
                        {
                            comment.Delete();
                        }
                    }
                    catch
                    {
                        // Continue deleting remaining comments
                    }
                    finally
                    {
                        ReleaseComObject(comment);
                    }
                }
            }
            catch
            {
                // If CommentsThreaded is not available, continue
            }
            finally
            {
                ReleaseComObject(threadedComments);
            }
        }

        // Remove Legacy Notes
        private void RemoveNotes(Excel.Worksheet worksheet)
        {
            Excel.Range noteCells = null;
            Excel.Areas areas = null;

            try
            {
                // Explicitly delete the Comment object
                noteCells = worksheet.Cells.SpecialCells(
                    Excel.XlCellType.xlCellTypeComments);

                if (noteCells == null)
                {
                    return;
                }

                areas = noteCells.Areas;

                for (int areaIndex = 1; areaIndex <= areas.Count; areaIndex++)
                {
                    Excel.Range area = null;

                    try
                    {
                        area = areas[areaIndex];

                        foreach (Excel.Range cell in area.Cells)
                        {
                            DeleteNoteFromCell(cell);
                        }
                    }
                    catch
                    {
                        // Continue with the next area
                    }
                    finally
                    {
                        ReleaseComObject(area);
                    }
                }
            }
            catch (COMException)
            {
                // SpecialCells throws when no cells contain Notes
            }
            catch
            {
                // Ignore unexpected errors while processing Notes
            }
            finally
            {
                ReleaseComObject(areas);
                ReleaseComObject(noteCells);
            }
        }

        // Delete the Note on a single cell
        private void DeleteNoteFromCell(Excel.Range cell)
        {
            Excel.Comment comment = null;

            try
            {
                comment = cell.Comment;

                // Orphaned indicator
                if (comment == null)
                {
                    try
                    {
                        comment = cell.AddComment(string.Empty);
                    }
                    catch
                    {
                        // Nothing further can be done for this cell
                        return;
                    }
                }

                if (comment != null)
                {
                    comment.Delete();
                }
            }
            catch
            {
                // Continue with the next cell
            }
            finally
            {
                ReleaseComObject(comment);
                ReleaseComObject(cell);
            }
        }

        // Release COM Object
        private void ReleaseComObject(object obj)
        {
            try
            {
                if (obj != null && Marshal.IsComObject(obj))
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
            catch
            {
                // Ignore COM cleanup errors
            }
        }
    }
}