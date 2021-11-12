using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ivi.Visa;

namespace AutocalMyp
{
    public class _instrumentb
    {
        public string modelo { get; set; }        
        public string resource { get; set; }
        public string informe { get; set; }
        public string modo { get; set; }
        public IMessageBasedSession device { get; set; }
        public string file { get; set; }
    }
}
