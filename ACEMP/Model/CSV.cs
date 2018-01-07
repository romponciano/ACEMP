using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACEMP.Model
{
    class CSV
    {
        public DataTable csvFinal { get; set; }
        public string nomeEmpresa { get; set; }
        public string mes { get; set; }
        public string ano { get; set; }
        public string cnpj { get; set; }
        public Boolean temExterior { get; set; }
        public Tributo tributo { get; set; }
        public List<int> clientesExterior { get; set; }

        public CSV(DataTable csvFinal, string nomeEmpresa, string mes, string ano, string cnpj)
        {
            this.csvFinal = csvFinal;
            this.nomeEmpresa = nomeEmpresa;
            this.mes = mes;
            this.ano = ano;
            this.cnpj = cnpj;
            this.temExterior = false;
            clientesExterior = new List<int>();
        }
    }
}
