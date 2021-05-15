using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vscmd {
  public class ComObject {
    public ComObject(object obj) {
      Debug.Assert(obj != null);
      this.Object = obj;
    }

    protected dynamic Object { get; }

    public T PropertyValue<T>(string name) {
      return (T)this.Properties.Item(name).Value;
    }

    public void SetPropertyValue<T>(string name, T value) {
      this.Properties.Item(name).Value = value;
    }

    public void WritePropertyNames() {
      foreach (dynamic property in this.Properties) {
        Console.WriteLine(property.Name);
      }
    }

    dynamic Properties {
      get {
        Debug.Assert((object)this.Object.Properties != null);
        return this.Object.Properties;
      }
    }
  }

  public class ComChildObject<TParent> : ComObject {
    public ComChildObject(object obj, TParent parent)
        : base(obj) {
      this.Parent = parent;
    }

    public TParent Parent { get; }
  }
}
