using System;
using System.Drawing;

namespace ClassLibrary
{
    public class ColorStruct
    {
        public int ID { get; set; }
        public int ProjectID { get; set; }

        public int ProjectNumber { get; set; }

        public int RowHandle { get; set; }

        public string ColumnFieldName { get; set; }

        public Color Color { get { return Color.FromArgb(this.ARGBColor); } }

        public int ARGBColor { get; set; }

        public ColorStruct ()
        {

        }

        public ColorStruct(object projectID, object column, object aRGBColor, object projectID2)
        {
            this.ProjectID = Convert.ToInt32(projectID);
            this.ColumnFieldName = ConvertObjectToString(column);
            this.ARGBColor = Convert.ToInt32(aRGBColor);
            //this.Color = Color.FromArgb(this.ARGBColor);
            this.ProjectNumber = Convert.ToInt32(projectID2);
        }

        public ColorStruct(object projectID, string column, Color color, int aRGBColor)
        {
            this.ProjectID = Convert.ToInt32(projectID);
            this.ColumnFieldName = ConvertObjectToString(column);
            this.ARGBColor = Convert.ToInt32(aRGBColor);
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
