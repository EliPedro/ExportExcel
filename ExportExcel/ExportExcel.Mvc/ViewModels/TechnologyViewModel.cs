using ExportExcel.Mvc.Models;
using System.Collections.Generic;

namespace ExportExcel.Mvc.ViewModels
{
    public class TechnologyViewModel
    {
        public List<Technology> Technologies
        {
            get
            {
                return StaticData.Technologies;
            }
        }
    }
}