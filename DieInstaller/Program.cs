using Microsoft.Deployment.WindowsInstaller;
using System;
using System.IO;
using WixSharp;
using WixSharp.Controls;

namespace Die_Setup
{
    internal class Program
    {
        static void Main()
        {
            var buildOutputDir = @"..\release\die";

            var project = new ManagedProject("Detect-It-Easy",
                new Dir(@"%ProgramFiles%\Detect-It-Easy",
                    new DirFiles((Path.Combine(buildOutputDir, "*.*"))),
                    new ExeFileShortcut("Detect-It-Easy", "[INSTALLDIR]\\Die.exe", "")
                    {
                        WorkingDirectory = "[INSTALLDIR]",
                    }));

            project.GUID = new Guid("6fe30b47-2577-43ad-9095-1861ba25889b");

            // Define custom dialogs
            project.ManagedUI = new ManagedUI();
            project.ManagedUI.InstallDialogs.Add<WelcomeDialog>();
            project.ManagedUI.InstallDialogs.Add<ProgressDialog>();
            project.ManagedUI.InstallDialogs.Add<ExitDialog>();

            // Define custom actions
            var msixInstallAction = new ManagedAction(CustomActions.InstallMSIX, Return.ignore, When.Before, Step.InstallFiles, Condition.NOT_Installed);
            project.Actions = new WixSharp.Action[] { msixInstallAction };

            project.WixSourceGenerated += (doc) =>
            {
                var welcomeDialog = doc.Root.Select("UI/Dialog[@Id='WelcomeDlg']");
                if (welcomeDialog != null)
                {
                    var titleControl = welcomeDialog.Select("Control[@Id='Title']");
                    if (titleControl != null)
                    {
                        titleControl.SetAttribute("Text", "Welcome to the Die Installer!");
                    }
                }
            };

            project.BuildMsi();
        }
    }

    public class CustomActions
    {
        [CustomAction]
        public static ActionResult InstallMSIX(Session session)
        {
            session.Log("Installing MSIX package...");
            return ActionResult.Success;
        }
    }

    public class WelcomeDialog : WixSharp.IDialog
    {
        public void OnExecute(Dialog dialog)
        {
            // Show a custom welcome form
            var welcomeForm = new WelcomeForm();
            welcomeForm.ShowDialog();
        }

        public bool OnNext(Dialog dialog)
        {
            return true;
        }
    }



    public class ProgressDialog : WixSharp.IDialog
    {
        public void OnExecute(Dialog dialog)
        {
            // Implement dialog execution logic
        }

        public bool OnNext(Dialog dialog)
        {
            return true;
        }
    }

    public class ExitDialog : WixSharp.IDialog
    {
        public void OnExecute(Dialog dialog)
        {
            // Implement dialog execution logic
        }

        public bool OnNext(Dialog dialog)
        {
            return true;
        }
    }
}
