using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnergyAnalysis.Models
{
    public class Element
    {
        public string id { get; set; }//元素ID
        public List<float> energy { get; set; }//能谱数据
        public string dieTime { get; set; }//死区时间
        public string probeTemperature { get; set; }//探头温度
        public string batteryVoltage { get; set; }//电池电压
        public string collectionTime { get; set; }//采集时间
    }
}
