using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Capstone.Models
{
    [Serializable]
    public class Company
    {
        public string? Code { get; set; }
        public string? Name { get; set; }
        public string? City { get; set; }
        public string? IndependentAuditingFirm { get; set; }
    }
}
