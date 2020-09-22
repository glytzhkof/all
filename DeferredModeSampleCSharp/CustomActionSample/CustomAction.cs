using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using System.Windows.Forms;

namespace CustomAction2
{
    public class CustomActions
    {
        [CustomAction]
        public static ActionResult TestCustomAction(Session session)
        {
            // System.Diagnostics.Debugger.Launch();
            session.Log("Begin TestCustomAction");

            string data = session["CustomActionData"];

            MessageBox.Show (data);

            return ActionResult.Success;
        }
    }
}
