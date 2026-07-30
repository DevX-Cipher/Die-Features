using System;
using System.Windows.Forms;

namespace Die_Setup
{
    public class WelcomeForm : Form
    {
        private Button continueButton;
        private Label welcomeLabel;

        public WelcomeForm()
        {
            continueButton = new Button();
            welcomeLabel = new Label();

            continueButton.Text = "Continue";
            continueButton.Dock = DockStyle.Bottom;
            continueButton.Click += new EventHandler(ContinueButton_Click);

            welcomeLabel.Text = "Welcome to the DiE Installer!";
            welcomeLabel.Dock = DockStyle.Top;

            this.Controls.Add(continueButton);
            this.Controls.Add(welcomeLabel);

            this.Text = "Welcome";
        }

        private void ContinueButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
