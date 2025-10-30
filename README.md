# How to Improve the Performance of Clipboard Paste Operation in WinForms GridControl?

This example demonstrates that how to improve the performance for clipboard paste operation in [WinForms GridControl](https://www.syncfusion.com/winforms-ui-controls/grid-control).

Pasting performance can be improved by using the [Model.ClipboardPaste](https://help.syncfusion.com/cr/windowsforms/Syncfusion.Windows.Forms.Grid.GridModel.html#Syncfusion_Windows_Forms_Grid_GridModel_ClipboardPaste) event and pasting the copied text manually for the selected range. This method of pasting avoids the iteration of cells while getting the values from the Clipboard.

``` csharp
//Event Triggering
this.Model.ClipboardPaste += model_ClipboardPaste;
 
//Event Customization
public void model_ClipboardPaste(object sender, GridCutPasteEventArgs e)
{    
    GridRangeInfoList rangeList;
    GridModel model = sender as GridModel;
    model.Selections.GetSelectedRanges(out rangeList, true);
    GridRangeInfo range = rangeList.GetOuterRange(rangeList.ActiveRange);
    //Getting data from clipboard.
    string psz = GetClipboardText();
    this.PasteTextFromBuffer(psz, range, e.ClipboardFlags);
    e.Handled = true;
}
 
private string GetClipboardText()
{
    string buffer = null;
    IDataObject iData = null;
    if (GridUtil.IsSet(this.Model.CutPaste.ClipboardFlags, GridDragDropFlags.Styles | GridDragDropFlags.Text))
    {
        iData = Clipboard.GetDataObject();
    }
    if (GridUtil.IsSet(this.Model.CutPaste.ClipboardFlags, GridDragDropFlags.Text)
                    && iData != null)
    {
        if (iData.GetDataPresent(DataFormats.UnicodeText))
        {
            buffer = iData.GetData(DataFormats.UnicodeText) as string;
        }
        else if (iData.GetDataPresent(DataFormats.Text))
        {
            buffer = iData.GetData(DataFormats.Text) as string;
        }
    }
    return buffer;
}
 
private bool PasteTextFromBuffer(string psz, GridRangeInfo range, int dragDropFlags)
{
    bool canceled = false;
 
    OperationFeedback op = new OperationFeedback(Model);
    try
    {
        op.AllowRollback = true;
        op.AllowNestedProgress = false;
 
        int rowIndex, colIndex;
 
        Model.ConfirmChanges();
        Model.CommandStack.BeginTrans("Paste");
 
        rowIndex = range.Top;
        colIndex = range.Left;
 
 
        int nLastCol = colIndex;
        int size = psz.Length;
 
        try
        {
            string[] copiedValue = psz.Split(new[] { "\r\n" }, StringSplitOptions.None);
 
            for (int i = 0; i < copiedValue.Length; i++)
            {
                string[] value = copiedValue[i].Split(new[] { "\t" }, StringSplitOptions.None);
 
                for (int j = 0; j < value.Length; j++)
                {
                    GridStyleInfo style = null;
                    style = Model[rowIndex, colIndex];
                    this.Model.TextDataExchange.PasteTextRowCol(rowIndex, colIndex, value[j]);
                    colIndex++;
                }
                rowIndex++;
                colIndex = range.Left;
            }
            if (size > 0)
            {
                op.PercentComplete = (int)((rowIndex * colIndex) * 100 / size);
            }
 
            if (op.ShouldCancel)
            {
                throw new GridUserCanceledException();
            }
        }
        catch (GridUserCanceledException ex)
        {
            TraceUtil.TraceExceptionCatched(ex);
            if (!ExceptionManager.RaiseExceptionCatched(this, ex))
            {
                throw;
            }
 
            canceled = true;
        }
 
        if (canceled && op.RollbackConfirmed)
        {
            Model.CommandStack.Rollback();
        }
        else
        {
            Model.CommandStack.CommitTrans();
        }
        //// Also formula refresh cells that have references to the pasted cells.
        Model.Refresh();
 
        return !canceled;
    }
    finally
    {
        op.Close();
        Model.EndUpdate();
    }
}
```

Take a moment to peruse the WinForms [GridControl - Clipboard support](https://help.syncfusion.com/windowsforms/grid-control/copy-and-paste#paste) documentation, where you can find about GridControl clipboard operation with code examples.