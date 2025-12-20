using System;
using System.Windows.Forms; // <--- Importante: Sem isso o DataGridView dá erro
using System.Reflection;

namespace Escala
{
    public static class ExtensionMethods
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo? pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
            if (pi != null)
            {
                pi.SetValue(dgv, setting, null);
            }
        }
    }
}