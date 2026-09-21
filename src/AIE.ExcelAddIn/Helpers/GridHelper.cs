using System;
using System.Drawing;
using System.Windows.Forms;

namespace AIE.ExcelAddIn.Helpers
{
    public static class GridHelper
    {
        public static void PaintMergedHeader(object sender, DataGridViewCellPaintingEventArgs e, DataGridView dgv, 
            int startCol1, int endCol1, string header1,
            int startCol2, int endCol2, string header2)
        {
            if (e.RowIndex == -1 && e.ColumnIndex >= 0)
            {
                e.Handled = true;

                // Determine if this column is part of a merged header
                bool isMerged1 = e.ColumnIndex >= startCol1 && e.ColumnIndex <= endCol1;
                bool isMerged2 = e.ColumnIndex >= startCol2 && e.ColumnIndex <= endCol2;
                bool isMerged = isMerged1 || isMerged2;

                // Save original clip and set clip to current cell to prevent bleeding
                Region oldClip = e.Graphics.Clip;
                e.Graphics.SetClip(e.CellBounds);

                // 1. Draw background
                using (Brush backBrush = new SolidBrush(e.CellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(backBrush, e.CellBounds);
                }

                int midY = e.CellBounds.Top + (e.CellBounds.Height / 2);

                // 2. Draw border
                using (Pen gridPen = new Pen(dgv.GridColor))
                {
                    // Top border (always drawn)
                    e.Graphics.DrawLine(gridPen, e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Right, e.CellBounds.Top);
                    // Bottom border (always drawn)
                    e.Graphics.DrawLine(gridPen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right, e.CellBounds.Bottom - 1);
                    
                    if (isMerged)
                    {
                        // Draw horizontal divider line in the middle
                        e.Graphics.DrawLine(gridPen, e.CellBounds.Left, midY, e.CellBounds.Right, midY);
                        
                        // Right border: only draw full-height for the LAST column in the group
                        // For intermediate columns, only draw in the BOTTOM half (sub-header area)
                        int endCol = isMerged1 ? endCol1 : endCol2;
                        if (e.ColumnIndex == endCol)
                        {
                            // Last column in merged group: full right border
                            e.Graphics.DrawLine(gridPen, e.CellBounds.Right - 1, e.CellBounds.Top, e.CellBounds.Right - 1, e.CellBounds.Bottom);
                        }
                        else
                        {
                            // Intermediate column: only bottom half right border (sub-header separator)
                            e.Graphics.DrawLine(gridPen, e.CellBounds.Right - 1, midY, e.CellBounds.Right - 1, e.CellBounds.Bottom);
                        }
                    }
                    else
                    {
                        // Non-merged columns: full right border
                        e.Graphics.DrawLine(gridPen, e.CellBounds.Right - 1, e.CellBounds.Top, e.CellBounds.Right - 1, e.CellBounds.Bottom);
                    }
                }

                // 3. Draw text
                using (Brush foreBrush = new SolidBrush(e.CellStyle.ForeColor))
                {
                    StringFormat format = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center,
                        Trimming = StringTrimming.EllipsisCharacter
                    };

                    if (isMerged)
                    {
                        // Draw sub-header in the bottom half
                        Rectangle subRect = new Rectangle(e.CellBounds.Left, midY, e.CellBounds.Width, e.CellBounds.Height / 2);
                        
                        string text = e.Value?.ToString() ?? "";
                        if (text.StartsWith("ĐM ")) text = text.Substring(3);
                        if (text.StartsWith("CP ")) text = text.Substring(3);
                        if (text == "Khác (%)") text = "Khác";

                        e.Graphics.DrawString(text, e.CellStyle.Font, foreBrush, subRect, format);

                        // Draw main header across the merged group (clipping slices it per cell)
                        int startCol = isMerged1 ? startCol1 : startCol2;
                        int endCol = isMerged1 ? endCol1 : endCol2;
                        string mainHeader = isMerged1 ? header1 : header2;

                        int startX = e.CellBounds.Left;
                        for (int i = startCol; i < e.ColumnIndex; i++)
                        {
                            startX -= dgv.Columns[i].Width;
                        }

                        int totalWidth = 0;
                        for (int i = startCol; i <= endCol; i++)
                        {
                            totalWidth += dgv.Columns[i].Width;
                        }
                        
                        Rectangle mainRect = new Rectangle(
                            startX, 
                            e.CellBounds.Top, 
                            totalWidth, 
                            e.CellBounds.Height / 2);
                        
                        e.Graphics.DrawString(mainHeader, e.CellStyle.Font, foreBrush, mainRect, format);
                    }
                    else
                    {
                        // Normal header: draw centered in full bounds
                        format.FormatFlags = StringFormatFlags.NoClip;
                        e.Graphics.DrawString(e.Value?.ToString(), e.CellStyle.Font, foreBrush, e.CellBounds, format);
                    }
                }
                // Restore clip
                e.Graphics.Clip = oldClip;
            }
        }

        public class MergedHeaderGroup
        {
            public int StartCol { get; set; }
            public int EndCol { get; set; }
            public string HeaderText { get; set; } = string.Empty;

            public MergedHeaderGroup(int startCol, int endCol, string headerText)
            {
                StartCol = startCol;
                EndCol = endCol;
                HeaderText = headerText;
            }
        }

        public static void PaintMultiMergedHeaders(object sender, DataGridViewCellPaintingEventArgs e, DataGridView dgv, 
            params MergedHeaderGroup[] groups)
        {
            if (e.RowIndex == -1 && e.ColumnIndex >= 0)
            {
                e.Handled = true;

                MergedHeaderGroup matchedGroup = null;
                if (groups != null)
                {
                    for (int g = 0; g < groups.Length; g++)
                    {
                        if (e.ColumnIndex >= groups[g].StartCol && e.ColumnIndex <= groups[g].EndCol)
                        {
                            matchedGroup = groups[g];
                            break;
                        }
                    }
                }

                bool isMerged = matchedGroup != null;

                Region oldClip = e.Graphics.Clip;
                e.Graphics.SetClip(e.CellBounds);

                // 1. Draw background
                using (Brush backBrush = new SolidBrush(e.CellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(backBrush, e.CellBounds);
                }

                int midY = e.CellBounds.Top + (e.CellBounds.Height / 2);

                // 2. Draw border
                using (Pen gridPen = new Pen(dgv.GridColor))
                {
                    // Top border
                    e.Graphics.DrawLine(gridPen, e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Right, e.CellBounds.Top);
                    // Bottom border
                    e.Graphics.DrawLine(gridPen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right, e.CellBounds.Bottom - 1);

                    if (isMerged)
                    {
                        // Horizontal divider
                        e.Graphics.DrawLine(gridPen, e.CellBounds.Left, midY, e.CellBounds.Right, midY);

                        if (e.ColumnIndex == matchedGroup.EndCol)
                        {
                            e.Graphics.DrawLine(gridPen, e.CellBounds.Right - 1, e.CellBounds.Top, e.CellBounds.Right - 1, e.CellBounds.Bottom);
                        }
                        else
                        {
                            e.Graphics.DrawLine(gridPen, e.CellBounds.Right - 1, midY, e.CellBounds.Right - 1, e.CellBounds.Bottom);
                        }
                    }
                    else
                    {
                        e.Graphics.DrawLine(gridPen, e.CellBounds.Right - 1, e.CellBounds.Top, e.CellBounds.Right - 1, e.CellBounds.Bottom);
                    }
                }

                // 3. Draw text
                using (Brush foreBrush = new SolidBrush(e.CellStyle.ForeColor))
                {
                    StringFormat format = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center,
                        Trimming = StringTrimming.EllipsisCharacter
                    };

                    if (isMerged)
                    {
                        // Sub-header in bottom half
                        Rectangle subRect = new Rectangle(e.CellBounds.Left, midY, e.CellBounds.Width, e.CellBounds.Height / 2);
                        string text = e.Value?.ToString() ?? "";
                        e.Graphics.DrawString(text, e.CellStyle.Font, foreBrush, subRect, format);

                        // Main header in top half
                        int startX = e.CellBounds.Left;
                        for (int i = matchedGroup.StartCol; i < e.ColumnIndex; i++)
                        {
                            startX -= dgv.Columns[i].Width;
                        }

                        int totalWidth = 0;
                        for (int i = matchedGroup.StartCol; i <= matchedGroup.EndCol; i++)
                        {
                            totalWidth += dgv.Columns[i].Width;
                        }

                        Rectangle mainRect = new Rectangle(startX, e.CellBounds.Top, totalWidth, e.CellBounds.Height / 2);
                        e.Graphics.DrawString(matchedGroup.HeaderText, e.CellStyle.Font, foreBrush, mainRect, format);
                    }
                    else
                    {
                        format.FormatFlags = StringFormatFlags.NoClip;
                        e.Graphics.DrawString(e.Value?.ToString(), e.CellStyle.Font, foreBrush, e.CellBounds, format);
                    }
                }

                e.Graphics.Clip = oldClip;
            }
        }

        /// <summary>
        /// Tự động co giãn toàn bộ dòng và cột cho DataGridView phù hợp với dữ liệu hiện có
        /// </summary>
        public static void AutoFit(this DataGridView dgv)
        {
            if (dgv == null || dgv.Columns.Count == 0) return;
            try
            {
                dgv.SuspendLayout();
                if (dgv.Rows.Count > 0)
                {
                    dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                    dgv.AutoResizeRows(DataGridViewAutoSizeRowsMode.DisplayedCells);
                }
                else
                {
                    dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.ColumnHeader);
                }

                // Đảm bảo các cột tên / nội dung có độ rộng tối thiểu dễ nhìn
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    if (col.Visible)
                    {
                        string header = col.HeaderText ?? "";
                        string name = col.Name ?? "";
                        if (header.Contains("Tên") || header.Contains("Nội dung") || name.Contains("Ten") || name.Contains("NoiDung"))
                        {
                            if (col.Width < 220) col.Width = 220;
                        }
                        else if (col.Width < 60)
                        {
                            col.Width = 60;
                        }
                    }
                }
            }
            catch { }
            finally
            {
                dgv.ResumeLayout();
            }
        }
    }
}
