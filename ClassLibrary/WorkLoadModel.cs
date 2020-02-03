using System;

namespace ClassLibrary
{
    public class WorkLoadModel
    {
        public string ToolNumber { get; set; }
        public int MWONumber { get; set; }
        public int ProjectNumber { get; set; }
        public string Stage { get; set; }
        public string Customer { get; set; }
        public string PartName { get; set; }
        public int DeliveryInWeeks { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? FinishDate { get; set; }
        public DateTime? AdjustedDeliveryDate { get; set; }
        public int MoldCost { get; set; }
        public string Engineer { get; set; }
        public string Designer { get; set; }
        public string ToolMaker { get; set; }
        public string RoughProgrammer { get; set; }
        public string FinisherProgrammer { get; set; }
        public string ElectrodeProgrammer { get; set; }
        public string Manifold { get; set; }
        public string MoldBase { get; set; }
        public string GeneralNotes { get; set; }
    }
}
