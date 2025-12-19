using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CAVISTAAPI.Models
{
    public class ReportModel
    {
        public string ReportID { get; set; }
        public string TableID { get; set; }
        public string ColumnID { get; set; }
        public string ColumnName { get; set; }
        public string DisplayColumnName { get; set; }
        public string Width { get; set; }
        public string Visible { get; set; }
        public string Alignment { get; set; }
        public string DisplayIndex { get; set; }
        public string IsHiddenColumn { get; set; }
        public string Total { get; set; }
        public string TotalYN { get; set; }
    }
    public class ReportColumnDataModel
    {
        public string field { get; set; }
        public string header { get; set; }
        public string type { get; set; }
        public int width { get; set; }
        public string align { get; set; }        
        public bool visible { get; set; }        
        public bool EnableColumnMenu { get; set; }
        public bool ShowinColumnOption { get; set; }        
        public string total { get; set; }
        public string TotalYN { get; set; }
        public bool EnableSum { get; set; }
        public bool EnableAvg { get; set; }
        public string precision { get; set; }
        public bool ClickPopup { get; set; }
    }
    public class ReportParameters
    {
        public string ParameterID { get; set; }
        public string ReportID { get; set; }
        public string ParameterName { get; set; }
        public string ParameterType { get; set; }
        public string IsMandatory { get; set; }
        public string ParamOrder { get; set; }
        public string AutolistName { get; set; }
        public string ProcedureName { get; set; }
        public string SendFiltersDetail { get; set; }
        public List<ReportFilters> lstvFilters { get; set; }
    }
    public class ReportFilters
    {
        public string Param1 { get; set; }
        public string Param2 { get; set; }
        public string Param4 { get; set; }
        public string Param3 { get; set; }
        public string Param5 { get; set; }
    }
}