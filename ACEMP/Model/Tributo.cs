using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACEMP.Model
{
    class Tributo
    {
        public float pis { get; set; }
        public float cofins { get; set; }
        public float csll { get; set; }
        public float irpj { get; set; }

        public Tributo(float pis, float cofins, float csll, float irpj)
        {
            this.pis = pis;
            this.cofins = cofins;
            this.csll = csll;
            this.irpj = irpj;
        }
    }
}
