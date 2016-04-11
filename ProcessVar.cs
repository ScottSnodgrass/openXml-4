using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElancoPimsDdsParser
{
    // Process variables kinda/sorta line up with Process (in Operation.Phase.Process)
    // But there can be several of these vars within a single Process (or AspenScript)
    class ProcessVar
    {
        public string Name { get; set; }
        public string Units { get; set; }

        public ProcessVar(string name)
        {
            Name = name;
        }
    }
}
