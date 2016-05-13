using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace MeetingAddIn
{
    public partial class ThisAddIn
    {
        Office.CommandBar newToolBar;
        Office.CommandBarButton firstButton;
        Office.CommandBarButton secondButton;
        Outlook.Explorers selectExplorers;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var selectExplorers = this.Application.Explorers;
            selectExplorers.NewExplorer += new Outlook
                .ExplorersEvents_NewExplorerEventHandler(newExplorer_Event);
            AddToolbar();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
       
        private void newExplorer_Event(Outlook.Explorer new_Explorer)
        {
            ((Outlook._Explorer)new_Explorer).Activate();
            newToolBar = null;
            AddToolbar();
        }

        private void AddToolbar()
        {

            if (newToolBar == null)
            {
                Office.CommandBars cmdBars =
                    this.Application.ActiveExplorer().CommandBars;
                newToolBar = cmdBars.Add("Meet-O",
                    Office.MsoBarPosition.msoBarTop, false, true);
            }
            try
            {
                Office.CommandBarButton button_1 =
                    (Office.CommandBarButton)newToolBar.Controls
                    .Add(1, missing, missing, missing, missing);
                button_1.Style = Office
                    .MsoButtonStyle.msoButtonCaption;
                button_1.Caption = "       New Meet-O";
                button_1.Tag = "New Meet-O";
                
                if (this.firstButton == null)
                {
                    this.firstButton = button_1;
                    firstButton.Click += new Office.
                        _CommandBarButtonEvents_ClickEventHandler
                        (ButtonClick);
                }

                
                newToolBar.Visible = true;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonClick(Office.CommandBarButton ctrl,
                ref bool cancel)
        {
            var outlookApp = new Outlook.Application();
            Outlook.AppointmentItem newAppointment = outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem);
            newAppointment.Start = DateTime.Now.AddHours(2);
            newAppointment.End = DateTime.Now.AddHours(2).AddMinutes(20);
            newAppointment.Location = "ConferenceRoom #2345";
            newAppointment.Body = "We will discuss progress on the group project.<span style=\"color: red\">Expensive meeting: 4200€</span>";
            newAppointment.AllDayEvent = false;
            newAppointment.Subject = "Meet-o Project";
            newAppointment.Recipients.Add("romain.linsolas@sgcib.com");
            Outlook.Recipients sentTo = newAppointment.Recipients;
            Outlook.Recipient sentInvite = null;
            sentInvite = sentTo.Add("Holly Holt");
            sentInvite.Type = (int)Outlook.OlMeetingRecipientType
                .olRequired;
            sentInvite = sentTo.Add("David Junca ");
            sentInvite.Type = (int)Outlook.OlMeetingRecipientType
                .olOptional;
            sentTo.ResolveAll();
            newAppointment.Save();
            newAppointment.Display(true);
        }

        #region Code généré par VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
