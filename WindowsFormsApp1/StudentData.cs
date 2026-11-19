using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class StudentData
    {
        public List<object> KisZhValues { get; set; } = new List<object>();
        public List<object> NagyZhValues { get; set; } = new List<object>();
        public List<object> BeadandoValues { get; set; } = new List<object>();

        public double TotalEarnedPoints { get; set; } = 0;
        public double PointPercentage { get; set; } = 0;
        public int ZeroCount { get; set; } = 0;
    }
}
