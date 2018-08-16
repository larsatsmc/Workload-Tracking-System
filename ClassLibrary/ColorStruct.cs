using DevExpress.XtraGrid.Columns;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary
{
    public class ColorStruct
    {
        public int ProjectID { get; set; }

        public int RowHandle { get; set; }

        public string Column { get; set; }

        public Color Color { get; set; }

        public int ColorARGB { get; set; }

        public ColorStruct ()
        {

        }

        public ColorStruct(object projectID, object column, object aRGBColor)
        {
            this.ProjectID = Convert.ToInt32(projectID);
            this.Column = ConvertObjectToString(column);
            this.ColorARGB = Convert.ToInt32(aRGBColor);
            this.Color = Color.FromArgb(this.ColorARGB);
        }

        public ColorStruct(object projectID, string column, Color color, int aRGBColor)
        {
            this.ProjectID = Convert.ToInt32(projectID);
            this.Column = ConvertObjectToString(column);
            this.ColorARGB = Convert.ToInt32(aRGBColor);
        }

        private string ConvertObjectToString(object obj)
        {
            if (obj != null)
            {
                return obj.ToString();
            }
            else
            {
                return "";
            }
        }
    }
}
