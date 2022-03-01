using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LatexToMathType.Models
{
    public class MathTypeModel
    {
        public string latex { get; set; }
        public string ole { get; set; }
        public string wmf { get; set; }
        public string lQid { get; set; }
        public string type { get; set; }
    }
}