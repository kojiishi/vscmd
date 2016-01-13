using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vscmd {
    partial class VisualStudio {
        public Debugger Debugger { get { return new Debugger(this.Object.Debugger); } }
    }

    public partial class Debugger : ComObject {
        public Debugger(object obj) : base(obj) {
        }
    }
}
