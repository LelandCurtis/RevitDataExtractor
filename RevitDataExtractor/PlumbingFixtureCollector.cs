using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitDataExtractor
{
    internal class PlumbingFixtureCollector
    {
        public string Family { get; set; }
        public string TypeMark { get; set; }
        public string Description { get; set; }
        public string Type { get; set; }
        public string Manufacturer { get; set; }
        public int Count { get; set; }
    }
}
