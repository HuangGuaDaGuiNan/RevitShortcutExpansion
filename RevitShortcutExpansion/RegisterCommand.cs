using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;

namespace RevitShortcutExpansion
{
    public class RegisterCommand : IExternalApplication
    {
        static readonly string commandToDisable = "ID_KEYBOARD_SHORTCUT_DIALOG";
        static RevitCommandId commandId;

        public Result OnStartup(UIControlledApplication application)
        {
            commandId = RevitCommandId.LookupCommandId(commandToDisable);
            if (!commandId.CanHaveBinding)
            {
                ShowDialog("Error", "命令" + commandToDisable + "重写失败");
                return Result.Failed;
            }
            try
            {
                AddInCommandBinding commandBinding = application.CreateAddInCommandBinding(commandId);
                commandBinding.Executed += CommandBinding_Executed;
            }
            catch (Exception)
            {
                ShowDialog("Error", "命令" + commandToDisable + "重写失败");
            }
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            if (commandId.HasBinding) application.RemoveAddInCommandBinding(commandId);
            return Result.Succeeded;
        }

        private void CommandBinding_Executed(object sender, Autodesk.Revit.UI.Events.ExecutedEventArgs e)
        {
            Assembly assembly = Assembly.Load(Assembly.GetExecutingAssembly().Location);
            IExternalCommand externalCommand = assembly.CreateInstance("RevitShortcutExpansion.ShortcutTemplate") as IExternalCommand;
            if (externalCommand != null)
            {
                string message = "ShortcutExpansion";
                externalCommand.Execute(null, ref message, null);
            }
        }

        private static void ShowDialog(string title, string message)
        {
            TaskDialog taskDialog = new TaskDialog(title)
            {
                MainInstruction = message,
                TitleAutoPrefix = false
            };
            taskDialog.Show();
        }

    }
}
