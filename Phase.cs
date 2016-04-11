using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElancoPimsDdsParser
{
    class Phase
    {
        List<AspenScript> aspenScriptList = new List<AspenScript>();
        List<ProcessVar> processVarList = new List<ProcessVar>();
        public string Name { get; private set; }

        public Phase(string name)
        {
            Name = name;
        }
        public void addAspenScript(AspenScript script)
        {
            aspenScriptList.Add(script);
        }
        public void addProcessVar(ProcessVar pv)
        {
            processVarList.Add(pv);
        }
    }
}
