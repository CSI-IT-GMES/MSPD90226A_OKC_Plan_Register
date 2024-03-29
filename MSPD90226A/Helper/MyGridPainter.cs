// Developer Express Code Central Example:
// How to merge cells horizontally in GridView
// 
// This example demonstrates how to merge cells located in the same row. The main
// idea is to paint merged cell manually.
// You can find a helper class in this
// example, which can be easily connected to your existing GridView.
// 
// You can find sample updates and versions for different programming languages here:
// http://www.devexpress.com/example=E2472

using DevExpress.Utils.Drawing;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.Drawing;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace CSI
{
    public class MyGridPainter : GridPainter
    {


        public MyGridPainter(GridView view)
            : base(view)
        {

        }

        private bool _IsCustomPainting;
        public bool IsCustomPainting
        {
            get { return _IsCustomPainting; }
            set { _IsCustomPainting = value; }
        }

        public void DrawMergedCell(MyMergedCell cell, PaintEventArgs e)
        {
            int delta = cell.Column1.VisibleIndex - cell.Column2.VisibleIndex;
            if (Math.Abs(delta) > 1)
                return;
            GridViewInfo vi = View.GetViewInfo() as GridViewInfo;
            GridCellInfo gridCellInfo1 = vi.GetGridCellInfo(cell.RowHandle, cell.Column1);
            GridCellInfo gridCellInfo2 = vi.GetGridCellInfo(cell.RowHandle, cell.Column2);
            if (gridCellInfo1 == null || gridCellInfo2 == null)
                return;
            Rectangle targetRect = Rectangle.Union(gridCellInfo1.Bounds, gridCellInfo2.Bounds);
            gridCellInfo1.Bounds = targetRect;
            gridCellInfo1.CellValueRect = targetRect;
            gridCellInfo2.Bounds = targetRect;
            gridCellInfo2.CellValueRect = targetRect;
            if (delta < 0)
                gridCellInfo1 = gridCellInfo2;
            if (gridCellInfo1.ViewInfo == null)   //yjkim modify
                return;
            Rectangle bounds = gridCellInfo1.ViewInfo.Bounds;
            //bounds.Location = new System.Drawing.Point(bounds.Location.X + 1, bounds.Location.Y + 1);
            bounds.Width = targetRect.Width - 2;
            bounds.Height = targetRect.Height - 2;
            gridCellInfo1.ViewInfo.Bounds = bounds;
            gridCellInfo1.ViewInfo.CalcViewInfo(e.Graphics);
            IsCustomPainting = true;
            GraphicsCache cache = new GraphicsCache(e.Graphics);
            gridCellInfo1.Appearance.FillRectangle(cache, gridCellInfo1.Bounds);
            DrawRowCell(new GridViewDrawArgs(cache, vi, vi.ViewRects.Bounds), gridCellInfo1);
            IsCustomPainting = false; ;
        }

    }

}
