using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vscmd {
  partial class VisualStudio {
    IEnumerable<dynamic> ProjectObjectsRecursive() {
      foreach (var item in this.Solution.Projects) {
        foreach (var project in ProjectObjectsFromItem(item))
          yield return project;
      }
    }

    static IEnumerable<dynamic> ProjectObjectsFromItem(dynamic item, int level = 0) {
      //Console.WriteLine(new string(' ', level * 4) + item.Name + "=" + item.Kind);
      if (item.ConfigurationManager != null) {
        yield return item;
        yield break;
      }

      var projectItems = item.ProjectItems;
      if (projectItems != null) {
        foreach (var projectItem in projectItems) {
          var subProjectItem = projectItem.SubProject;
          if (subProjectItem != null) {
            foreach (var subProject in ProjectObjectsFromItem(subProjectItem, level + 1))
              yield return subProject;
          }
        }
      }
    }

    Project ProjectByName(string name) {
      foreach (var item in this.ProjectObjectsRecursive()) {
        if (name == item.Name)
          return Project.FromObject(item);
      }
      throw new KeyNotFoundException(name);
    }
    string StartupProjectName {
      get { return this.Solution.Properties.Item("StartupProject").Value; }
    }

    public Project StartupProject {
      get { return this.ProjectByName(this.StartupProjectName); }
    }

    public partial class Project : ComObject {
      protected Project(object project)
          : base(project) { }

      internal static Project FromObject(dynamic project) {
        switch ((string)project.Kind) {
          case "{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}":
            return new CPlusPlusProject(project);
          case "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}":
            return new CSharpProject(project);
          default:
            return new Project(project);
        }
      }

      public string Name { get { return this.Object.Name; } }

      public dynamic Configuration(string name) {
        return this.Object.Object.Configurations[name];
      }

      internal NotSupportedException KindNotSupportedException() {
        return new NotSupportedException(string.Format(
            "The operation is not suuported for this project type {0}.", this.Object.Kind));
      }
    }

    partial class CPlusPlusProject : Project {
      internal CPlusPlusProject(object project)
          : base(project) { }
    }

    partial class CSharpProject : Project {
      internal CSharpProject(object project)
          : base(project) { }
    }
  }
}
