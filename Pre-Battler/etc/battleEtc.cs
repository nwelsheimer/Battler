using Infragistics.Win.UltraWinGrid;
using System.Windows.Forms;

namespace Pre_Battler
{
    public static class battleEtc
    { 
    public static void GridNavigation(UltraGrid grid, KeyEventArgs e)
    {
        switch (e.KeyCode)
        {
            case Keys.Up:
                grid.PerformAction(UltraGridAction.ExitEditMode, false, false);
                grid.PerformAction(UltraGridAction.AboveCell, false, false);
                e.Handled = true;
                grid.PerformAction(UltraGridAction.EnterEditMode, false, false);
                break;
            case Keys.Down:
                grid.PerformAction(UltraGridAction.ExitEditMode, false, false);
                grid.PerformAction(UltraGridAction.BelowCell, false, false);
                e.Handled = true;
                grid.PerformAction(UltraGridAction.EnterEditMode, false, false);
                break;
            case Keys.Enter:
                grid.PerformAction(UltraGridAction.ExitEditMode, false, false);
                grid.PerformAction(UltraGridAction.BelowCell, false, false);
                e.Handled = true;
                grid.PerformAction(UltraGridAction.EnterEditMode, false, false);
                break;
        }
    }

        public static string GridLayout(UltraGrid grid, int mode, string settings = "")
        {
            byte[] loa;
            System.IO.MemoryStream lo;

            switch (mode)
            {
                case 1://We are doing a save
                    lo = new System.IO.MemoryStream();
                    grid.DisplayLayout.Save(lo, PropertyCategories.All);
                    loa = lo.ToArray();
                    settings = System.Convert.ToBase64String(loa);
                    return settings;
                case 2://We are doing a load
                    loa = System.Convert.FromBase64String(settings);
                    lo = new System.IO.MemoryStream(loa);
                    lo.Seek(0, System.IO.SeekOrigin.Begin);
                    try { grid.DisplayLayout.Load(lo, PropertyCategories.All); } catch { }
                    return "SUCCESS";
            }
            return "FAILED";
        }
    }
}