using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vscmd {
  partial class VisualStudio {
    partial class Project {
      public Configuration ActiveConfiguration {
        get { return this.CreateConfiguration(this.Object.ConfigurationManager.ActiveConfiguration); }
      }

      protected virtual Configuration CreateConfiguration(object configuration) {
        return new Configuration(configuration, this);
      }
    }

    public class Configuration : ComChildObject<Project> {
      internal Configuration(object configuration, Project parent)
          : base(configuration, parent) { }

      public Project Project { get { return this.Parent; } }
      public string Name { get { return this.Object.ConfigurationName; } }

      public virtual string DebugStartArguments {
        get { throw this.Project.KindNotSupportedException(); }
        set { throw this.Project.KindNotSupportedException(); }
      }

      public virtual string DebugStartProgram {
        get { throw this.Project.KindNotSupportedException(); }
        set { throw this.Project.KindNotSupportedException(); }
      }
    }

    partial class CPlusPlusProject {
      protected override Configuration CreateConfiguration(object configuration) {
        return new CPlusPlusConfiguration(configuration, this);
      }
    }

    // https://msdn.microsoft.com/en-us/library/microsoft.visualstudio.vcproject.vcprojectconfigurationproperties.aspx
    class CPlusPlusConfiguration : Configuration {
      public CPlusPlusConfiguration(object configuration, Project project)
          : base(configuration, project) {
        var vcconfig = project.Configuration(this.Name);
        this._VCDebugSettings = vcconfig.DebugSettings;
      }

      public override string DebugStartArguments {
        get { return this._VCDebugSettings.CommandArguments; }
        set { this._VCDebugSettings.CommandArguments = value; }
      }

      public override string DebugStartProgram {
        get { return this._VCDebugSettings.Command; }
        set { this._VCDebugSettings.Command = value; }
      }

      private dynamic _VCDebugSettings;
    }

    partial class CSharpProject {
      protected override Configuration CreateConfiguration(object configuration) {
        return new CSharpConfiguration(configuration, this);
      }
    }

    class CSharpConfiguration : Configuration {
      public CSharpConfiguration(object configuration, Project parent)
          : base(configuration, parent) { }

      public override string DebugStartArguments {
        get { return this.PropertyValue<string>("StartArguments"); }
        set { this.SetPropertyValue<string>("StartArguments", value); }
      }

      public override string DebugStartProgram {
        get {
          int action = PropertyValue<int>("StartAction");
          return action == 0 ? null : PropertyValue<string>("StartProgram");
        }

        set {
          if (string.IsNullOrEmpty(value)) {
            SetPropertyValue("StartAction", 0);
            return;
          }
          if (File.Exists(value)) {
            SetPropertyValue("StartAction", 1);
            SetPropertyValue("StartProgram", Path.GetFullPath(value));
            return;
          }
          SetPropertyValue("StartAction", 1);
          SetPropertyValue("StartProgram", value);
        }
      }
    }
  }
}
