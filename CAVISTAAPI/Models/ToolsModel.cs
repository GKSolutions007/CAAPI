using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CAVISTAAPI.Models
{
    public class ToolsModel
    {
    }
    public class MapRoleModel
    {
        public MapRoleModel()
        {
            children = new List<MapRoleModel>();
        }
        public int UserID { get; set; }
        public IEnumerable<SelectListItem> UserName { get; set; }
        public string id { get; set; }
        public string text { get; set; }
        public clsState state { get; set; }
        public List<MapRoleModel> children { get; set; }
    }
    public class clsState
    {
        public bool selected { get; set; }

        public clsState()
        {
            selected = false;
        }
    }

    public class checkedIDs
    {
        public int IDs { get; set; }
    }
    public class SaveMapRoleRequest
    {
        public List<checkedIDs> checked_ids { get; set; }
        public string RoleName { get; set; }
        public string UID { get; set; }
    }
}