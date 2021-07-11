using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnergyAnalysis.Models
{
    /// <summary>
    /// 标定元素
    /// </summary>
    public class CalibrationElement
    {
        public string id { get; set; }
        public string symbol { get; set; }
        public string name { get; set; }
        public string Ka { get; set; }
        public string Kb { get; set; }
        public string La { get; set; }
        public string Lb { get; set; }
        public string Lg { get; set; }
        public string Ll { get; set; }
    }
}
