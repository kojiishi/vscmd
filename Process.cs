using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vscmd {
    partial class Debugger {
        public IEnumerable<Process> LocalProcesses {
            get { return ((IEnumerable<dynamic>)this.Object.LocalProcesses)
                .Select(p => new Process(p)); }
        }
    }

    public class Process : ComObject {
        public Process(object obj) : base(obj) {
        }

        public string Name { get { return this.Object.Name; } }
        public int ProcessID { get { return this.Object.ProcessID; } }

        public void Attach() {
            this.Object.Attach();
        }
    }
}
