using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ElancoPimsDdsParser
{
    class Operation
    {
        List<Phase> phaseList = new List<Phase>();
        public string Name { get; private set; }

        public Operation(string name)
        {
            Name = name;
        }

        public void addPhase(Phase phase)
        {
            this.phaseList.Add(phase);
        }

        public List<Phase> getPhases()
        {
            return phaseList;
        }

        public bool addAspenScript(string[] idStrings, XElement tableNode)
        {
            bool foundPhase = false;

            AspenScript script = new AspenScript(idStrings);

            bool success = script.parseTable(tableNode);
            if (!success)
            {
                Console.WriteLine("ERROR - parsing script table!");
                return false;
            }
            Console.WriteLine("script text: \n{0}", script.ScriptText);
            foreach (var phase in phaseList)
            {
                if (idStrings[1].Equals(phase.Name))
                {
                    foundPhase = true;
                    phase.addAspenScript(script);
                    break;
                }
            }
            if (!foundPhase)
            {
                Console.WriteLine("ERROR - found OP but not Phase({0} for script!", idStrings[1]);
            }
            return foundPhase;
        }
    }
}
