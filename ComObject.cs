using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vscmd
{
    public class ComObject
    {
        public ComObject(object obj)
        {
            this.Object = obj;
        }

        protected dynamic Object { get; }

        public T PropertyValue<T>(string name)
        {
            return (T)this.Object.Properties.Item(name).Value;
        }
        public void SetPropertyValue<T>(string name, T value)
        {
            this.Object.Properties.Item(name).Value = value;
        }
    }

    public class ComChildObject<TParent> : ComObject
    {
        public ComChildObject(object obj, TParent parent)
            : base(obj)
        {
            this.Parent = parent;
        }

        public TParent Parent { get; }
    }
}
